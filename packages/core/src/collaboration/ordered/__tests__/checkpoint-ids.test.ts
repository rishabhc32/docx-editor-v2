// A suggestion keeps its id across a fold (design D14): the checkpoint lists each suggestion under
// its id, and the open makes it under that id. An id the engine mints passes over every kept one,
// a change id the room ordered before the fold is refused by name (design D14), and a session on
// the older file is refused by the file number, as before. A plain Word file gets new ids.

import { describe, expect, test } from 'bun:test';
import { orderedSeenCards } from '../change/apply-decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { OrderedLocalRoom } from '../client/local-room.ts';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { newSuggestion } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { decodeOrderedServerMessage, encodeOrderedClientMessage } from '../protocol/codec.ts';
import { ORDERED_PROTOCOL_VERSION, type OrderedServerMessage } from '../protocol/messages.ts';
import { allowWritersPolicy } from '../room/policy.ts';
import { OrderedRoom } from '../room/room.ts';
import { apply, APPLY_FILE, at, BODY } from './change-apply-fixture.ts';
import { STYLE_RECORDS, withRoomFacts } from './checkpoint-fixture.ts';
import { checkFold, checkpointBytes, folded, mediaStore } from './fold-check.ts';
import { actor, snapshot } from './room-harness.ts';
import { STYLED } from './style-revision-fixture.ts';

const ANA = actor('Ana', 'ana-1', 'organiser');
const T0 = Date.UTC(2026, 9, 28, 9, 30, 0);

/** Each suggestion's id, status and maker. */
const suggestionsOf = (doc: OrderedDocument) =>
  [...doc.suggestions.values()].map(({ id, status, attribution }) => ({ id, status, attribution }));

/** STYLED with Ana's insertion (`S1`) accepted: a fold lists `S2` to `S7`, and no `S1`. */
function styledWithS1Accepted() {
  const room = openOrderedWordFile(STYLE_RECORDS.source, { file: 1 });
  STYLE_RECORDS.edit(room.document);
  return room;
}

describe('a suggestion id across a fold', () => {
  test('every suggestion a fold keeps has its id, status and maker, after one fold and two', async () => {
    const room = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const doc = withRoomFacts(room.document);
    const blobs = mediaStore(room.base);
    const once = await folded(doc, room.base, 2, blobs);
    const twice = await folded(once.document, once.base, 3, blobs);
    const live = new Map(suggestionsOf(doc).map((s) => [s.id, s]));
    for (const copy of [once.document, twice.document]) {
      const kept = suggestionsOf(copy);
      expect(kept.map(({ id }) => live.get(id))).toEqual(kept);
    }

    // Every pending one but Ana's row height in her own new table, which a fold writes as part of
    // the table, and Ben's accepted control value, which the untracked record names
    expect(suggestionsOf(twice.document).map(({ id, status }) => `${id} ${status}`)).toEqual([
      'S:quin:1.1 pending',
      'S:pia:1.1 pending',
      'S:rae:1.1 pending',
      'S:ana:1.1 pending',
      'S:ana:3.1 pending',
      'S:ben:1.1 accepted',
      'S:pia:2.1 pending',
    ]);
    await checkFold(doc, room.base, blobs, 'every section');
  });

  test("the styles part's records keep their ids, beside a body mark of the same w:id", async () => {
    const { document: doc, base } = openOrderedWordFile(STYLED, { file: 1 });
    await checkFold(doc, base, mediaStore(base), 'styles');
    const once = await folded(doc, base, 2, mediaStore(base));
    expect([...once.document.partRevisions.keys()]).toEqual([...doc.partRevisions.keys()]);
  });

  test('a plain Word file gets new ids, as before', () => {
    const room = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const doc = withRoomFacts(room.document);
    const saved = openOrderedWordFile(orderedPackageBytes(doc, room.base), { file: 2 });
    const ids = [...saved.document.suggestions.keys()];
    expect(ids.length).toBeGreaterThan(0);
    expect(ids.every((id) => /^S[0-9]+$/.test(id))).toBe(true);
  });
});

describe('an id the engine mints', () => {
  test('after the open passes over every kept id', async () => {
    const room = styledWithS1Accepted();
    const once = await folded(room.document, room.base, 2, mediaStore(room.base));
    const kept = [...once.document.suggestions.keys()];
    const minted = newSuggestion(once.document, { author: 'Zed', person: null });
    expect(kept).not.toContain(minted.id);
  });
});

/** An insert at the start of a body paragraph, in the addresses of `doc`'s file. */
const insertAt = (doc: OrderedDocument, id: string, paragraph: number): OrderedChange => ({
  v: 1,
  id,
  kind: 'insert',
  part: BODY,
  mode: 'suggest',
  place: at(doc, paragraph, 0),
  text: 'Late ',
});

function roomOver(doc: OrderedDocument, errors: unknown[] = []): OrderedRoom {
  return new OrderedRoom({
    document: doc,
    applyChange: applyOrderedChange,
    policy: allowWritersPolicy,
    now: () => T0,
    onApplyError: (error) => errors.push(error),
  });
}

describe('a change after a fold', () => {
  test('a change id the room ordered before the fold is refused by name, never a throw', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const first = roomOver(doc);
    expect(first.submit(ANA, [insertAt(doc, 'ana-1:1', 0)]).refused).toEqual([]);
    const once = await folded(doc, base, 2, mediaStore(base));
    expect(once.document.suggestions.has('S:ana-1:1.1')).toBe(true);

    const errors: unknown[] = [];
    const next = roomOver(once.document, errors);
    const before = snapshot(once.document);
    const result = next.submit(ANA, [insertAt(once.document, 'ana-1:1', 1)]);
    expect(result.refused).toEqual([
      { id: 'ana-1:1', code: 'other-file', detail: "the room's file holds ana-1:1" },
    ]);
    expect(errors).toEqual([]);
    expect(snapshot(once.document)).toEqual(before);

    // The addresses stay: a later change made against the older file applies at its place
    expect(next.submit(ANA, [insertAt(doc, 'ana-1:2', 1)]).refused).toEqual([]);
    expect(once.document.suggestions.has('S:ana-1:2.1')).toBe(true);
  });

  test('a session on the older file must reload, and its decision on a kept id never applies', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(
      doc,
      'pia:1',
      { kind: 'insert', part: BODY, mode: 'suggest', place: at(doc, 0, 0), text: 'Late ' },
      'Pia'
    );
    const bytes = await checkpointBytes(doc, base, mediaStore(base));
    const room = new OrderedLocalRoom({ room: 'r', wordFile: bytes, file: 2, code: 'c' });
    const decide = (copy: OrderedDocument) =>
      ({
        v: 1,
        id: 'ana-1:1',
        kind: 'decide',
        cards: orderedSeenCards(copy, ['S:pia:1.1']),
        status: 'accepted',
      }) as const;

    // The session's copy is file 1, and it holds changes the room saved
    const received: OrderedServerMessage[] = [];
    const connection = room.transport({ ...ANA }).connect({
      open: () => undefined,
      close: () => undefined,
      message: (frame) => {
        const read = decodeOrderedServerMessage(frame);
        if (read.ok) {
          received.push(read.value);
        }
      },
    });
    const hello = { room: 'r', file: 1, last: 1, code: 'c', namespace: 'ana-1' };
    const v = ORDERED_PROTOCOL_VERSION;
    connection.send(encodeOrderedClientMessage({ v, type: 'hello', ...hello }));
    connection.send(encodeOrderedClientMessage({ v, type: 'decide', change: decide(doc) }));
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(received).toEqual([{ v, type: 'reload', reason: 'file' }]);
    expect(room.room.order).toBe(0);
    expect(room.room.document.suggestions.get('S:pia:1.1')?.status).toBe('pending');

    // The same decision from a copy of the new file applies to the kept id
    const fresh = openOrderedCheckpoint(bytes, { file: 2 }).document;
    expect(room.room.submit(ANA, [decide(fresh)]).refused).toEqual([]);
    expect(room.room.document.suggestions.get('S:pia:1.1')?.status).toBe('accepted');
  });
});
