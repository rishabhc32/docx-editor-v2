// A change after a close (design D14). The room's file keeps the item list with its addresses, so
// an address no longer tells a late change apart. The document keeps the highest applied counter
// of each namespace (`lastCounters`), and the room on the next file refuses a change at or under
// it by name (`other-file`). Here the next room opens the item list the close saved: the same
// document, as the open of a saved list gives it back.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange, OrderedStampedChange } from '../change/types.ts';
import { openOrderedRoomFile, openOrderedWordFile } from '../client/open.ts';
import { OrderedReplica } from '../client/replica.ts';
import { isHidden } from '../model/document.ts';
import type { OrderedDocument } from '../model/types.ts';
import { decodeOrderedClientMessage, encodeOrderedServerMessage } from '../protocol/codec.ts';
import { allowWritersPolicy } from '../room/policy.ts';
import { OrderedRoom, OrderedRoomOpenError } from '../room/room.ts';
import { checkpointBytes, mediaStore } from './fold-check.ts';
import { TestClient, type ChangeBody } from './room-client.ts';
import { actor, BODY, loadRoom, paragraphOf, ROOM_FILE, snapshot } from './room-harness.ts';

const ANA = actor('Ana', 'ana-1', 'organiser');
const BEN = actor('Ben', 'ben-1', 'organiser');

/** A page whose host keeps one namespace across its reloads. */
const PAGE = actor('Ana', 'ana-page', 'organiser');
const T0 = Date.UTC(2026, 10, 26, 9, 30, 0);

function roomOver(document: OrderedDocument, log?: readonly OrderedStampedChange[]): OrderedRoom {
  return new OrderedRoom({
    document,
    applyChange: applyOrderedChange,
    policy: allowWritersPolicy,
    now: () => T0,
    ...(log === undefined ? {} : { log }),
  });
}

/** An insert at a unit of the first paragraph, as a client fills it in. */
const body = (doc: OrderedDocument, offset: number, extra: { into?: string } = {}): ChangeBody => ({
  kind: 'insert',
  part: BODY,
  mode: 'suggest',
  place: { space: paragraphOf(doc, 0).text, offset },
  text: 'Late ',
  ...extra,
});

const insert = (doc: OrderedDocument, id: string, offset: number, extra?: { into?: string }) =>
  ({ v: 1, id, ...body(doc, offset, extra) }) as OrderedChange;

/** The text the story shows. */
const textOf = (doc: OrderedDocument): string =>
  (doc.stories.get(BODY)?.tree.toArray() ?? [])
    .filter((item) => item.kind === 'text' && !isHidden(doc, item))
    .map((item) => item.text)
    .join('');

/** The changes of a client's frame, as the room reads them. */
function changesOf(frame: string | null, namespace: string): readonly OrderedChange[] {
  const decoded = decodeOrderedClientMessage(frame ?? '', { namespace });
  if (!decoded.ok || decoded.value.type !== 'changes') {
    throw new Error('not a changes frame');
  }

  return decoded.value.changes;
}

/** The client's frame through the room; the room's answers back to the client. */
function deliver(room: OrderedRoom, client: TestClient, frame: string | null) {
  const result = room.submit(client.actor, changesOf(frame, client.actor.namespace));
  if (result.accepted.length > 0) {
    const changes = result.accepted;
    client.receive(encodeOrderedServerMessage({ v: 3, type: 'changes', changes }));
  }

  if (result.refused.length > 0) {
    const refusals = result.refused;
    client.receive(encodeOrderedServerMessage({ v: 3, type: 'refuse', refusals }));
  }

  return result;
}

describe('the highest applied counter of each namespace', () => {
  test('an applied change raises it; a refused one leaves it as it was', () => {
    const room = roomOver(loadRoom());
    const doc = room.document;
    expect(doc.lastCounters).toEqual(new Map());
    expect(room.submit(ANA, [insert(doc, 'ana-1:4', 0)]).refused).toEqual([]);
    expect(room.submit(BEN, [insert(doc, 'ben-1:2', 2)]).refused).toEqual([]);
    expect(doc.lastCounters).toEqual(
      new Map([
        ['ana-1', 4],
        ['ben-1', 2],
      ])
    );

    // The model refuses it once it applies: its undo record lowers the counter again
    const before = snapshot(doc);
    const into = insert(doc, 'ana-1:9', 4, { into: 'S:ben-1:2.1' });
    expect(room.submit(ANA, [into]).refused).toEqual([
      { id: 'ana-1:9', code: 'apply-refused', reason: 'not-own-suggestion', detail: 'S:ben-1:2.1' },
    ]);
    expect(doc.lastCounters).toEqual(
      new Map([
        ['ana-1', 4],
        ['ben-1', 2],
      ])
    );
    expect(snapshot(doc)).toEqual(before);
  });

  test("a take-back of one's own suggestion is a new change: the counter goes up", () => {
    const room = roomOver(loadRoom());
    const doc = room.document;
    expect(room.submit(ANA, [insert(doc, 'ana-1:1', 0)]).refused).toEqual([]);
    const retract = { v: 1, id: 'ana-1:2', kind: 'retract', suggestions: ['S:ana-1:1.1'] } as const;
    expect(room.submit(ANA, [retract]).refused).toEqual([]);
    expect(textOf(doc)).not.toContain('Late');
    expect(doc.lastCounters.get('ana-1')).toBe(2);
  });
});

describe('a change after a close', () => {
  test('a change the room applied before the close, sent again after it, is refused by name', () => {
    const first = roomOver(loadRoom());
    const ana = new TestClient(loadRoom(), ANA);
    ana.local(body(ana.doc, 0));
    ana.local(body(ana.doc, 4));
    const frame = ana.flush();

    // The room applies both and closes before the client hears of them
    expect(first.submit(ANA, changesOf(frame, 'ana-1')).refused).toEqual([]);
    const shown = textOf(first.document);
    expect(shown.match(/Late/g)).toHaveLength(2);

    // The client sends its pending changes again to the room on the next file
    const next = roomOver(first.document);
    const before = snapshot(next.document);
    const result = deliver(next, ana, frame);
    expect(result.accepted).toEqual([]);
    expect(result.refused).toEqual([
      { id: 'ana-1:1', code: 'other-file', detail: "the room's file holds ana-1:2" },
      { id: 'ana-1:2', code: 'other-file', detail: "the room's file holds ana-1:2" },
    ]);
    expect(snapshot(next.document)).toEqual(before);
    expect(textOf(next.document)).toBe(shown);
  });

  test('a counter at or under the saved one is refused; a higher one and other namespaces apply', () => {
    const first = roomOver(loadRoom());
    const doc = first.document;
    expect(first.submit(ANA, [insert(doc, 'ana-1:3', 0)]).refused).toEqual([]);
    expect(first.submit(ANA, [insert(doc, 'ana-1:5', 0)]).refused).toEqual([]);

    const next = roomOver(doc);
    expect(next.order).toBe(0);
    const late = next.submit(ANA, [insert(doc, 'ana-1:4', 0), insert(doc, 'ana-1:5', 0)]);
    expect(late.refused.map(({ id, code }) => ({ id, code }))).toEqual([
      { id: 'ana-1:4', code: 'other-file' },
      { id: 'ana-1:5', code: 'other-file' },
    ]);
    expect(late.acks).toEqual([]);

    // Order numbers start at 1 again; a repeated id of this file keeps its first order number
    expect(next.submit(ANA, [insert(doc, 'ana-1:6', 0)]).acks).toEqual([{ id: 'ana-1:6', n: 1 }]);
    expect(next.submit(ANA, [insert(doc, 'ana-1:6', 0)]).acks).toEqual([{ id: 'ana-1:6', n: 1 }]);
    expect(next.submit(BEN, [insert(doc, 'ben-1:1', 0)]).refused).toEqual([]);
    expect(doc.lastCounters).toEqual(
      new Map([
        ['ana-1', 6],
        ['ben-1', 1],
      ])
    );
  });

  test('a log entry made before the room file does not rebuild the room', () => {
    const first = roomOver(loadRoom());
    const doc = first.document;
    expect(first.submit(ANA, [insert(doc, 'ana-1:1', 0)]).refused).toEqual([]);
    const log = first.joinSnapshot(0);

    // The next file holds the change, so the old log replays it never
    expect(() => roomOver(doc, log)).toThrow(OrderedRoomOpenError);
    expect(() => roomOver(doc, log)).toThrow('log-apply at 1: other-file');
  });
});

describe("the organiser's revert after a close", () => {
  const direct = (doc: OrderedDocument, id: string) =>
    ({ ...insert(doc, id, 0), mode: 'edit' }) as OrderedChange;
  const revert = (id: string, suggestion: string) =>
    ({ v: 1, id, kind: 'revert', suggestions: [suggestion], status: 'rejected' }) as const;

  test('of a direct edit made before the file is refused by name; of a later one it applies', () => {
    const first = roomOver(loadRoom());
    const doc = first.document;
    expect(first.submit(ANA, [direct(doc, 'ana-1:1')]).refused).toEqual([]);

    // An editor that opens the next file has no undo of it either
    const next = roomOver(doc);
    const before = snapshot(doc);
    expect(next.submit(ANA, [revert('ana-1:2', 'S:ana-1:1.1')]).refused).toEqual([
      { id: 'ana-1:2', code: 'other-file', detail: "the room's file holds ana-1:1" },
    ]);
    expect(snapshot(doc)).toEqual(before);
    expect(textOf(doc)).toContain('Late');

    expect(next.submit(ANA, [direct(doc, 'ana-1:3')]).refused).toEqual([]);
    expect(next.submit(ANA, [revert('ana-1:4', 'S:ana-1:3.1')]).refused).toEqual([]);
    expect(textOf(doc).match(/Late/g)).toHaveLength(1);
  });
});

describe("an author's undo of own edits after a close", () => {
  const undo = (id: string, change: string) =>
    ({ v: 1, id, kind: 'retract', suggestions: [], undo: [change] }) as const;

  test('of a change made before the file is refused by name; of a later one it applies', () => {
    const first = roomOver(loadRoom());
    const doc = first.document;
    const into = { into: 'S:ana-1:1.1' };
    expect(first.submit(ANA, [insert(doc, 'ana-1:1', 0)]).refused).toEqual([]);
    expect(first.submit(ANA, [insert(doc, 'ana-1:2', 0, into)]).refused).toEqual([]);

    // An editor that opens the next file has no undo of it
    const next = roomOver(doc);
    const before = snapshot(doc);
    expect(next.submit(ANA, [undo('ana-1:3', 'ana-1:2')]).refused).toEqual([
      { id: 'ana-1:3', code: 'other-file', detail: "the room's file holds ana-1:2" },
    ]);
    expect(snapshot(doc)).toEqual(before);
    expect(textOf(doc).match(/Late/g)).toHaveLength(2);

    expect(next.submit(ANA, [insert(doc, 'ana-1:4', 0, into)]).refused).toEqual([]);
    expect(textOf(doc).match(/Late/g)).toHaveLength(3);
    expect(next.submit(ANA, [undo('ana-1:5', 'ana-1:4')]).refused).toEqual([]);
    expect(textOf(doc).match(/Late/g)).toHaveLength(2);
  });
});

describe('a page that keeps its namespace across a close', () => {
  const opened = () => openOrderedWordFile(ROOM_FILE, { file: 1 });
  const page = (doc: OrderedDocument) =>
    new OrderedReplica({ doc, person: PAGE.person, namespace: PAGE.namespace, now: () => T0 });

  /** The page types at the start of the first paragraph; returns what it sends. */
  const typed = (replica: OrderedReplica) => {
    replica.applyLocal({ v: 1, id: replica.nextId(), ...body(replica.doc, 0) } as OrderedChange);
    return replica.takeUnsent();
  };

  test('its new copy counts on past the counters the file holds, and the room applies it', async () => {
    const first = roomOver(opened().document);
    const before = page(opened().document);
    expect(first.submit(PAGE, typed(before)).refused).toEqual([]);
    expect(first.submit(PAGE, typed(before)).refused).toEqual([]);

    // The close: the saved item list holds the counters the room held
    const { base } = opened();
    const bytes = await checkpointBytes(first.document, base, mediaStore(base));
    const reopened = () => openOrderedRoomFile(bytes, { file: 2 }).document;
    expect(reopened().lastCounters).toEqual(first.document.lastCounters);
    const next = roomOver(reopened());

    // The page loads again with the same namespace, and types
    const sent = typed(page(reopened()));
    expect(sent.map((change) => change.id)).toEqual(['ana-page:3']);
    const result = next.submit(PAGE, sent);
    expect(result.refused).toEqual([]);
    expect(result.acks).toEqual([{ id: 'ana-page:3', n: 1 }]);
    expect(textOf(next.document).match(/Late/g)).toHaveLength(3);
  });
});
