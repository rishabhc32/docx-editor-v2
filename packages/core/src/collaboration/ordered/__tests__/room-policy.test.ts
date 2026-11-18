// The suggestions-only policy gives the host's earlier verdicts, rule by rule.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { actor, BODY, HEADER, loadRoom, paragraphOf, snapshot, realApply } from './room-harness.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const BEN = actor('Ben', 'ben-1', 'party');
const RITA = actor('Rita', 'rita-1', 'reader');

/** A room where Ana typed "XYZ" at offset 4 and struck "Due" (two suggestions). */
function setup() {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadRoom(),
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
  });
  const p0 = paragraphOf(room.document, 0);
  const p1 = paragraphOf(room.document, 1);
  const h0 = paragraphOf(room.document, 0, HEADER);
  const counters = new Map<string, number>();
  /** The card holding a suggestion, as the organiser's copy (the room's here) shows it. */
  const seen = (id: string) => orderedCardSeen(room.document, id);
  /** One change in a batch. */
  const send = (who: typeof ANA, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const batch = [{ v: 1, id: `${who.namespace}:${k}`, ...change }];
    const result = room.submit(who, batch);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  const edit = { part: BODY, mode: 'suggest' };
  expect(
    send(ANA, { kind: 'insert', ...edit, place: { space: p0.text, offset: 4 }, text: 'XYZ' })
  ).toBe('accepted');
  expect(
    send(ANA, {
      kind: 'delete',
      ...edit,
      from: { space: p1.text, offset: 0 },
      to: { space: p1.text, offset: 3 },
    })
  ).toBe('accepted');
  const typed = 'i:ana-1:1.1';
  return {
    room,
    send,
    edit,
    p0,
    p1,
    h0,
    typed,
    seen,
    anaInsert: 'S:ana-1:1.1',
    anaDelete: 'S:ana-1:2.1',
  };
}

describe('the suggestions-only policy', () => {
  test('a party may not decide; an organiser may', () => {
    const { room, send, seen, anaInsert } = setup();
    const before = snapshot(room.document);
    expect(
      send(ANA, { kind: 'decide', cards: [{ suggestions: seen(anaInsert) }], status: 'accepted' })
    ).toBe('policy-decide');
    expect(
      send(BEN, { kind: 'decide', cards: [{ suggestions: seen(anaInsert) }], status: 'rejected' })
    ).toBe('policy-decide');
    expect(snapshot(room.document)).toEqual(before);
    expect(
      send(OLGA, { kind: 'decide', cards: [{ suggestions: seen(anaInsert) }], status: 'accepted' })
    ).toBe('accepted');
    expect(room.document.suggestions.get(anaInsert)?.status).toBe('accepted');
  });

  test('a party may not send a revert; an organiser reverts only their own direct edits', () => {
    const { room, send, p0, anaInsert } = setup();
    // Olga types directly: her suggestion is accepted at once and marked direct
    expect(
      send(OLGA, {
        kind: 'insert',
        part: BODY,
        mode: 'edit',
        place: { space: p0.text, offset: 0 },
        text: 'Q',
      })
    ).toBe('accepted');
    const olgaTyped = 'S:olga-1:1.1';
    expect(room.document.suggestions.get(olgaTyped)).toMatchObject({
      status: 'accepted',
      direct: true,
    });
    const before = snapshot(room.document);
    // A party's revert is a direct edit: refused by name, whoever's suggestion it names
    for (const id of [olgaTyped, anaInsert])
      expect(send(ANA, { kind: 'revert', suggestions: [id], status: 'rejected' })).toBe(
        'policy-direct/text'
      );
    expect(snapshot(room.document)).toEqual(before);
    // The organiser takes back only her own direct edit: not a party's suggestion, not twice
    expect(send(OLGA, { kind: 'revert', suggestions: [anaInsert], status: 'rejected' })).toBe(
      'apply-refused/not-own-suggestion'
    );
    expect(send(OLGA, { kind: 'revert', suggestions: [olgaTyped], status: 'rejected' })).toBe(
      'accepted'
    );
    expect(room.document.suggestions.get(olgaTyped)?.status).toBe('rejected');
    expect(send(OLGA, { kind: 'revert', suggestions: [olgaTyped], status: 'rejected' })).toBe(
      'apply-refused/stale-revert'
    );
    expect(send(OLGA, { kind: 'revert', suggestions: [olgaTyped], status: 'accepted' })).toBe(
      'accepted'
    );
  });

  test('an organiser may not revert a suggestion a decision accepted, even her own', () => {
    const { send, seen, edit, p0 } = setup();
    // Olga suggests, then decides her own suggestion: decided, so it has no inverse
    expect(
      send(OLGA, { kind: 'insert', ...edit, place: { space: p0.text, offset: 0 }, text: 'Q' })
    ).toBe('accepted');
    const olgaSuggested = 'S:olga-1:1.1';
    expect(
      send(OLGA, {
        kind: 'decide',
        cards: [{ suggestions: seen(olgaSuggested) }],
        status: 'accepted',
      })
    ).toBe('accepted');
    expect(send(OLGA, { kind: 'revert', suggestions: [olgaSuggested], status: 'rejected' })).toBe(
      'apply-refused/not-direct'
    );
  });

  // Today's check refuses a party's revision or a changed part in a header or footer
  // (`furniture`); a format press that changes no value writes neither
  test('a party’s header change that changes nothing passes; one that changes a value does not', () => {
    const { room, send, edit, h0 } = setup();
    const format = (properties: readonly Record<string, unknown>[], mode = 'suggest') => ({
      kind: 'format',
      ...edit,
      mode,
      part: HEADER,
      from: { space: h0.text, offset: 0 },
      to: { space: h0.text, offset: 6 },
      properties,
    });
    expect(send(OLGA, format([{ localName: 'i' }], 'edit'))).toBe('accepted');
    const header = () => serializeOoxmlPart(composeOrderedPart(room.document, HEADER).part);
    const before = header();
    expect(send(ANA, format([{ localName: 'i' }]))).toBe('accepted');
    expect(header()).toBe(before);
    expect(send(ANA, format([{ localName: 'b' }]))).toBe('policy-furniture/furniture');
    expect(header()).toBe(before);
  });

  test('a party may not change a header or a footer; an organiser may', () => {
    const { send, edit, h0 } = setup();
    const change = {
      kind: 'insert',
      ...edit,
      part: HEADER,
      place: { space: h0.text, offset: 0 },
      text: 'x',
    };
    expect(send(ANA, change)).toBe('policy-furniture/furniture');
    expect(send(OLGA, change)).toBe('accepted');
  });

  test('a party may not change anything inside another person suggestion', () => {
    const { send, edit, p0, p1, typed } = setup();
    // Typing inside Ana's typed text, and deleting or formatting it
    expect(
      send(BEN, { kind: 'insert', ...edit, place: { space: typed, offset: 1 }, text: 'b' })
    ).toBe('policy-nested/nested');
    const overTyped = { from: { space: p0.text, offset: 2 }, to: { space: p0.text, offset: 6 } };
    expect(send(BEN, { kind: 'delete', ...edit, ...overTyped })).toBe('policy-nested/nested');
    expect(
      send(BEN, { kind: 'format', ...edit, ...overTyped, properties: [{ localName: 'b' }] })
    ).toBe('policy-nested/nested');
    // A format on text Ana struck sits inside her deletion
    const struck = { from: { space: p1.text, offset: 1 }, to: { space: p1.text, offset: 2 } };
    expect(
      send(BEN, { kind: 'format', ...edit, ...struck, properties: [{ localName: 'i' }] })
    ).toBe('policy-nested/nested');
    // Beside Ana's text is not inside it; a second deletion of struck text is allowed (policy 1)
    expect(
      send(BEN, { kind: 'insert', ...edit, place: { space: typed, offset: 0 }, text: 'b' })
    ).toBe('accepted');
    expect(
      send(BEN, { kind: 'insert', ...edit, place: { space: p0.text, offset: 4 }, text: 'c' })
    ).toBe('accepted');
    expect(send(BEN, { kind: 'delete', ...edit, ...struck })).toBe('accepted');
    // Ana herself, and the organiser, may
    expect(
      send(ANA, { kind: 'insert', ...edit, place: { space: typed, offset: 1 }, text: 'a' })
    ).toBe('accepted');
    expect(
      send(OLGA, { kind: 'insert', ...edit, place: { space: typed, offset: 2 }, text: 'o' })
    ).toBe('accepted');
  });

  test('a person may take back their own suggestion, not another person', () => {
    const { send, anaInsert, anaDelete } = setup();
    expect(send(BEN, { kind: 'retract', suggestions: [anaInsert] })).toBe(
      'policy-not-own/others_suggestions'
    );
    expect(send(OLGA, { kind: 'retract', suggestions: [anaDelete] })).toBe(
      'policy-not-own/others_suggestions'
    );
    expect(send(ANA, { kind: 'retract', suggestions: [anaInsert, anaDelete] })).toBe('accepted');
  });

  test('an edit may add to the sender own pending suggestion only', () => {
    const { send, edit, p0, seen, anaInsert } = setup();
    const change = { kind: 'insert', ...edit, place: { space: p0.text, offset: 0 }, text: 'q' };
    expect(send(BEN, { ...change, into: anaInsert })).toBe('policy-not-own/others_suggestions');
    expect(send(OLGA, { ...change, into: anaInsert })).toBe('policy-not-own/others_suggestions');
    expect(send(ANA, { ...change, into: anaInsert })).toBe('accepted');
    expect(
      send(OLGA, { kind: 'decide', cards: [{ suggestions: seen(anaInsert) }], status: 'accepted' })
    ).toBe('accepted');
    expect(send(ANA, { ...change, into: anaInsert })).toBe('policy-direct/text');
    expect(send(ANA, { kind: 'retract', suggestions: [anaInsert] })).toBe('policy-direct/text');
  });

  test('a party may not edit directly; an organiser may', () => {
    const { send, p0 } = setup();
    const direct = { part: BODY, mode: 'edit' };
    const range = { from: { space: p0.text, offset: 0 }, to: { space: p0.text, offset: 3 } };
    expect(
      send(ANA, { kind: 'insert', ...direct, place: { space: p0.text, offset: 0 }, text: 'x' })
    ).toBe('policy-direct/text');
    expect(
      send(ANA, { kind: 'format', ...direct, ...range, properties: [{ localName: 'b' }] })
    ).toBe('policy-direct/package');
    expect(send(OLGA, { kind: 'delete', ...direct, ...range })).toBe('accepted');
  });

  test('a read-only person sends no change', () => {
    const { send, edit, p0 } = setup();
    expect(
      send(RITA, { kind: 'insert', ...edit, place: { space: p0.text, offset: 0 }, text: 'x' })
    ).toBe('policy-read-only');
    expect(send(RITA, { kind: 'retract', suggestions: ['S:rita-1:1.1'] })).toBe(
      'unknown-suggestion'
    );
  });

  test('a change under another person namespace is refused as a stamp', () => {
    const { send, edit, p0 } = setup();
    const change = { kind: 'insert', ...edit, place: { space: p0.text, offset: 0 }, text: 'x' };
    expect(send({ ...BEN, namespace: 'ana-1' }, change)).toBe('stamp/stamp');
  });

  test('a policy that throws refuses the change and applies nothing', () => {
    const doc = loadRoom();
    const room = new OrderedRoom({
      document: doc,
      applyChange: realApply,
      policy: {
        check: () => {
          throw new Error('the host lookup is down');
        },
      },
      now: () => Date.UTC(2026, 9, 28),
    });
    const p0 = paragraphOf(doc, 0);
    const change = {
      v: 1,
      id: 'ana-1:1',
      kind: 'insert',
      part: BODY,
      mode: 'suggest',
      place: { space: p0.text, offset: 0 },
      text: 'x',
    };
    expect(room.submit(ANA, [change]).refused.map((refusal) => refusal.code)).toEqual([
      'policy-failed',
    ]);
    expect(room.order).toBe(0);
  });

  test('an apply that does not pass the context on is not judged: the change is refused', () => {
    const doc = loadRoom();
    const room = new OrderedRoom({
      document: doc,
      applyChange: (d, change, { person, date, ids }) =>
        void realApply(d, change, { person, date, ids }),
      policy: suggestionsOnlyPolicy(),
      now: () => Date.UTC(2026, 9, 28),
    });
    const p0 = paragraphOf(doc, 0);
    const before = snapshot(doc);
    const insert = { kind: 'insert', part: BODY, mode: 'suggest', text: 'X' };
    const result = room.submit(ANA, [
      { v: 1, id: 'ana-1:1', ...insert, place: { space: p0.text, offset: 4 } },
    ]);
    expect(result.refused.map((r) => r.code)).toEqual(['policy-failed']);
    expect(snapshot(doc)).toEqual(before);
  });
});
