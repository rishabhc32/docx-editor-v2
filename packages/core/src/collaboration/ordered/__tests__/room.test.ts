// The room core: one order, stamps from the login and the clock, duplicates answered with their
// first order number, a change that throws taken back exactly, and a room rebuilt from its log.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { ORDERED_CHANGE_LIMITS, type OrderedStampedChange } from '../change/types.ts';
import { allowWritersPolicy } from '../room/policy.ts';
import {
  OrderedRoom,
  OrderedRoomOpenError,
  orderedStampDate,
  type OrderedApplyChange,
} from '../room/room.ts';
import { actor, BODY, loadRoom, paragraphOf, snapshot, realApply } from './room-harness.ts';
import { testAttribution, testPerson } from './test-persons.ts';

const ANA = actor('Ana', 'ana-1', 'organiser');
const BEN = actor('Ben', 'ben-1', 'organiser');
const T0 = Date.UTC(2026, 9, 28, 9, 30, 0);

function room(apply: OrderedApplyChange = realApply, log?: readonly OrderedStampedChange[]) {
  let tick = 0;
  return new OrderedRoom({
    document: loadRoom(),
    applyChange: apply,
    policy: allowWritersPolicy,
    now: () => T0 + 1500 * tick++,
    ...(log === undefined ? {} : { log }),
  });
}

function insert(id: string, offset: number, text: string, extra: Record<string, unknown> = {}) {
  const doc = loadRoom();
  const { text: space } = paragraphOf(doc, 0);
  return {
    v: 1,
    id,
    kind: 'insert',
    part: BODY,
    mode: 'suggest',
    place: { space, offset },
    text,
    ...extra,
  };
}

describe('the room', () => {
  test('orders changes and stamps n, by and at from the login and the clock', () => {
    const r = room();
    const first = r.submit(ANA, [insert('ana-1:1', 4, 'new ')]);
    const second = r.submit(BEN, [insert('ben-1:1', 0, 'A '), insert('ben-1:2', 8, 'x')]);
    expect(first.refused).toEqual([]);
    expect(second.refused).toEqual([]);
    expect([...first.accepted, ...second.accepted].map((c) => [c.id, c.n, c.by, c.at])).toEqual([
      ['ana-1:1', 1, testPerson('Ana'), '2026-10-28T09:30:00Z'],
      ['ben-1:1', 2, testPerson('Ben'), '2026-10-28T09:30:01Z'],
      ['ben-1:2', 3, testPerson('Ben'), '2026-10-28T09:30:01Z'],
    ]);
    expect(second.acks).toEqual([
      { id: 'ben-1:1', n: 2 },
      { id: 'ben-1:2', n: 3 },
    ]);
    expect(r.order).toBe(3);
    expect(r.document.suggestions.get('S:ana-1:1.1')?.attribution).toEqual(
      testAttribution('Ana', '2026-10-28T09:30:00Z')
    );
  });

  test('a repeated id gets the order number of its first copy and applies nothing', () => {
    const r = room();
    const batch = [insert('ana-1:1', 4, 'new '), insert('ana-1:2', 0, 'A ')];
    r.submit(ANA, batch);
    const before = snapshot(r.document);
    const again = r.submit(ANA, [...batch, insert('ana-1:3', 2, 'z')]);
    expect(again.acks).toEqual([
      { id: 'ana-1:1', n: 1 },
      { id: 'ana-1:2', n: 2 },
      { id: 'ana-1:3', n: 3 },
    ]);
    expect(again.accepted.map((c) => c.id)).toEqual(['ana-1:3']);
    // The same room without the repeated copies: the same document
    const once = room();
    once.submit(ANA, batch);
    once.submit(ANA, [insert('ana-1:3', 2, 'z')]);
    expect(snapshot(r.document)).toEqual(snapshot(once.document));
    expect(before).not.toEqual(snapshot(r.document));
  });

  test('a change that throws is taken back exactly and refused by name', () => {
    const errors: unknown[] = [];
    const halfway =
      (error: () => Error): OrderedApplyChange =>
      (doc, change, ctx) => {
        realApply(doc, change, ctx);
        throw error();
      };
    for (const [apply, code, reason] of [
      [halfway(() => new Error('boom')), 'apply-failed', undefined],
      [halfway(() => new OrderedRefusal('not-text', 'test')), 'apply-refused', 'not-text'],
    ] as const) {
      let tick = 0;
      const r = new OrderedRoom({
        document: loadRoom(),
        applyChange: apply,
        policy: allowWritersPolicy,
        now: () => T0 + tick++,
        onApplyError: (error) => errors.push(error),
      });
      const items = r.document.stories.get(BODY)?.tree.toArray() ?? [];
      const before = snapshot(r.document);
      const result = r.submit(ANA, [insert('ana-1:1', 4, 'new ')]);
      expect(result.accepted).toEqual([]);
      expect(result.refused).toEqual([
        {
          id: 'ana-1:1',
          code,
          ...(reason === undefined ? { detail: 'Error' } : { reason, detail: 'test' }),
        },
      ]);
      expect(snapshot(r.document)).toEqual(before);
      // The same item objects, not copies
      const after = r.document.stories.get(BODY)?.tree.toArray() ?? [];
      expect(after.length).toBe(items.length);
      expect(after.every((item, k) => item === items[k])).toBe(true);
      expect(r.document.undo).toBeNull();
      expect(r.order).toBe(0);
      expect(r.orderOf('ana-1:1')).toBeUndefined();
    }
    expect(errors.length).toBe(1);
  });

  test('a refused change refuses only itself; later ones in the batch still apply', () => {
    const r = room();
    const result = r.submit(ANA, [
      insert('ana-1:1', 4, 'a'),
      { ...insert('ana-1:2', 0, 'b'), place: { space: 'b:nowhere', offset: 0 } },
      insert('ana-1:3', 2, 'c', { into: 'S:ana-9:9.1' }),
      insert('ana-1:4', 6, 'd', { into: 'S:ana-1:1.1' }),
    ]);
    expect(result.accepted.map((c) => [c.id, c.n])).toEqual([
      ['ana-1:1', 1],
      ['ana-1:4', 2],
    ]);
    expect(result.refused.map((r) => [r.id, r.code])).toEqual([
      ['ana-1:2', 'unknown-address'],
      ['ana-1:3', 'unknown-suggestion'],
    ]);
  });

  test('a malformed batch is refused whole and applies nothing', () => {
    const r = room();
    const result = r.submit(ANA, [
      insert('ana-1:1', 4, 'a'),
      { ...insert('ana-1:2', 0, 'b'), n: 7 },
    ]);
    expect(result.accepted).toEqual([]);
    expect(result.refused.map((x) => [x.id, x.code])).toEqual([
      ['ana-1:1', 'server-field'],
      ['ana-1:2', 'server-field'],
    ]);
    expect(r.order).toBe(0);
  });

  test('an oversized batch is refused with a named code and applies nothing', () => {
    const r = room();
    const batch = Array.from({ length: ORDERED_CHANGE_LIMITS.batch + 1 }, (_, k) =>
      insert(`ana-1:${k + 1}`, 0, 'x')
    );
    const result = r.submit(ANA, batch);
    expect(result).toEqual({
      accepted: [],
      acks: [],
      refused: [{ id: null, code: 'too-many', detail: `over ${ORDERED_CHANGE_LIMITS.batch}` }],
    });
    expect(r.order).toBe(0);
  });

  test('ids outside the sender namespace, and a namespace another person used, are refused', () => {
    const r = room();
    expect(r.submit(ANA, [insert('ben-1:1', 0, 'x')]).refused.map((x) => x.code)).toEqual([
      'foreign-id',
    ]);
    r.submit(ANA, [insert('ana-1:1', 0, 'x')]);
    expect(r.namespaceOwner('ana-1')).toBe(testPerson('Ana').id);
    const thief = { ...BEN, namespace: 'ana-1' };
    const stolen = r.submit(thief, [insert('ana-1:2', 0, 'y')]);
    expect(stolen.refused).toEqual([
      {
        id: 'ana-1:2',
        code: 'stamp',
        reason: 'stamp',
        detail: 'the namespace belongs to another person',
      },
    ]);
    expect(r.order).toBe(1);
  });

  test('the join snapshot is the changes after an order number', () => {
    const r = room();
    r.submit(ANA, [
      insert('ana-1:1', 0, 'a'),
      insert('ana-1:2', 0, 'b'),
      insert('ana-1:3', 0, 'c'),
    ]);
    expect(r.joinSnapshot(0).map((c) => c.n)).toEqual([1, 2, 3]);
    expect(r.joinSnapshot(1).map((c) => c.n)).toEqual([2, 3]);
    expect(r.joinSnapshot(3)).toEqual([]);
    expect(r.joinSnapshot(9)).toEqual([]);
    expect(() => r.joinSnapshot(-1)).toThrow(RangeError);
    expect(() => r.joinSnapshot(1.5)).toThrow(RangeError);
  });

  test('a room rebuilt from its log has the same document, ids and namespaces', () => {
    const a = room();
    a.submit(ANA, [insert('ana-1:1', 4, 'new ')]);
    a.submit(BEN, [
      {
        v: 1,
        id: 'ben-1:1',
        kind: 'decide',
        cards: [{ suggestions: orderedCardSeen(a.document, 'S:ana-1:1.1') }],
        status: 'accepted',
      },
      insert('ben-1:2', 0, 'A '),
    ]);
    const b = room(realApply, a.joinSnapshot(0));
    expect(snapshot(b.document)).toEqual(snapshot(a.document));
    expect(b.order).toBe(3);
    expect(b.orderOf('ben-1:2')).toBe(3);
    expect(b.namespaceOwner('ben-1')).toBe(testPerson('Ben').id);
    // A change resent after a restart is a duplicate
    expect(b.submit(BEN, [insert('ben-1:2', 0, 'A ')]).acks).toEqual([{ id: 'ben-1:2', n: 3 }]);
    expect(
      b.submit({ ...ANA, namespace: 'ben-1' }, [insert('ben-1:3', 0, 'q')]).refused[0]?.code
    ).toBe('stamp');
  });

  test('a log that does not rebuild the room refuses to open', () => {
    const a = room();
    a.submit(ANA, [
      insert('ana-1:1', 0, 'a'),
      insert('ana-1:2', 0, 'b'),
      insert('ana-1:3', 0, 'c'),
    ]);
    const log = a.joinSnapshot(0);
    const [one, two, three] = log;
    if (!one || !two || !three) throw new Error('no log');
    const opens = (entries: readonly unknown[]) => {
      try {
        room(realApply, entries as OrderedStampedChange[]);
        return 'opened';
      } catch (error) {
        return error instanceof OrderedRoomOpenError ? `${error.code}@${error.at}` : 'other';
      }
    };
    expect(opens(log)).toBe('opened');
    expect(opens([one, three])).toBe('log-gap@3');
    expect(opens([one, { ...two, id: one.id }])).toBe('log-repeated-id@2');
    expect(opens([one, { ...two, by: testPerson('Ben') }])).toBe('log-stamp@2');
    expect(opens([one, { ...two, at: 'yesterday' }])).toBe('log-invalid@2');
    expect(opens([one, { ...two, into: 'S:none:1.1' }])).toBe('log-apply@2');
  });

  test('a bad actor or a failing clock throws before anything applies', () => {
    const r = room();
    expect(() => r.submit({ ...ANA, namespace: 'a b' }, [insert('ana-1:1', 0, 'x')])).toThrow(
      RangeError
    );
    const blank = { ...ANA.person, name: ' ' };
    expect(() => r.submit({ ...ANA, person: blank }, [insert('ana-1:1', 0, 'x')])).toThrow(
      RangeError
    );
    const broken = new OrderedRoom({
      document: loadRoom(),
      applyChange: realApply,
      policy: allowWritersPolicy,
      now: () => Number.NaN,
    });
    const before = snapshot(broken.document);
    expect(() => broken.submit(ANA, [insert('ana-1:1', 0, 'x')])).toThrow(RangeError);
    expect(snapshot(broken.document)).toEqual(before);
    expect(broken.order).toBe(0);
  });

  test('the stamp date is Word form, to the second', () => {
    expect(orderedStampDate(Date.UTC(2026, 0, 2, 3, 4, 5, 678))).toBe('2026-01-02T03:04:05Z');
    expect(() => orderedStampDate(Number.NaN)).toThrow(RangeError);
  });
});
