// Addresses carry the room's Word file number (design D8). A change made against Word file k
// names spaces of file k; the room on file k+1 loaded the same positions under other spaces, so
// the change never applies there, even where the same paragraph sits at the same position. It
// is refused by name, before the policy or the edit runs.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { orderedChangeTargets } from '../change/targets.ts';
import type { OrderedChange } from '../change/types.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { resolveOrderedPlace } from '../edits/places.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { loadedSpace } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { allowWritersPolicy } from '../room/policy.ts';
import { OrderedRoom } from '../room/room.ts';
import { actor, BODY, paragraphOf, ROOM_FILE, snapshot } from './room-harness.ts';
import { packageOf } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const ANA = actor('Ana', 'ana-1', 'organiser');
const T0 = Date.UTC(2026, 9, 28, 9, 30, 0);

function roomOn(file: number, bytes: Uint8Array = ROOM_FILE): OrderedRoom {
  return new OrderedRoom({
    document: loadOrderedPackage(packageOf(bytes), { file }),
    applyChange: applyOrderedChange,
    policy: allowWritersPolicy,
    now: () => T0,
  });
}

/** An insert at the start of the first paragraph, in the addresses of `file`'s load. */
function insertOn(file: number, id: string): OrderedChange {
  const { text } = paragraphOf(loadOrderedPackage(packageOf(ROOM_FILE), { file }), 0);
  return {
    v: 1,
    id,
    kind: 'insert',
    part: BODY,
    mode: 'suggest',
    place: { space: text, offset: 0 },
    text: 'Late ',
  };
}

describe('Word file numbers in addresses', () => {
  test('a loaded space names its Word file; a minted one names none', () => {
    const doc = loadOrderedPackage(packageOf(ROOM_FILE), { file: 7 });
    expect(doc.file).toBe(7);
    const { text, mark } = paragraphOf(doc, 0);
    expect(loadedSpace(text)).toEqual({ kind: 'b', file: 7, node: '/word/document.xml#0.0.0' });
    expect(loadedSpace(mark.space)?.file).toBe(7);
    expect(loadedSpace('i:ana-1:1.1')).toBeNull();
    expect(() => loadOrderedPackage(packageOf(ROOM_FILE), { file: -1 })).toThrow(RangeError);
  });

  test('the room on file k+1 refuses a change made against file k as other-file', () => {
    // File 2 is file 1 with one change folded in, as the room's close step writes it
    const first = roomOn(1);
    expect(first.submit(ANA, [insertOn(1, 'ana-1:1')]).refused).toEqual([]);
    const next = roomOn(2, orderedPackageBytes(first.document, packageOf(ROOM_FILE)));
    const before = snapshot(next.document);
    const result = next.submit(ANA, [insertOn(1, 'ana-1:2')]);
    expect(result.accepted).toEqual([]);
    expect(result.refused).toEqual([
      { id: 'ana-1:2', code: 'other-file', detail: expect.stringContaining('b:1:') },
    ]);
    expect(snapshot(next.document)).toEqual(before);
    // The same position in the new file's own addresses applies
    expect(next.submit(ANA, [insertOn(2, 'ana-1:3')]).refused).toEqual([]);
  });

  test('the model refuses an old address by name too, and changes nothing', () => {
    const doc = loadOrderedPackage(packageOf(ROOM_FILE), { file: 3 });
    const old = insertOn(2, 'ana-1:1');
    expect(orderedChangeTargets(doc, old)).toMatchObject({ ok: false, code: 'other-file' });
    const before = snapshot(doc);
    const ctx = {
      person: testPerson('Ana'),
      date: '2026-10-28T09:30:00Z',
      ids: orderedIdMinter(old.id),
    };
    let reason: string | null = null;
    try {
      withUndoRecord(doc, () => applyOrderedChange(doc, old, ctx));
    } catch (error) {
      if (error instanceof OrderedRefusal) reason = error.reason;
    }
    expect(reason).toBe('other-file');
    expect(snapshot(doc)).toEqual(before);
    if (old.kind !== 'insert' || 'mark' in old.place) throw new Error('an insert at a unit');
    const place = old.place;
    expect(() => resolveOrderedPlace(doc, BODY, place)).toThrow('other-file');
  });
});
