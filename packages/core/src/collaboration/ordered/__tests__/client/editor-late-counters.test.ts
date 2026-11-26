// The highest applied counter of each namespace (design D7), with real editors and one ordered
// room. Typing, undo and redo each send new changes, so a person's counter only goes up, and
// every copy holds the room's counters: the room's file keeps them at a close.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import type { OrderedDocument } from '../../model/types.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  typeKeys,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const countersOf = (doc: OrderedDocument) =>
  [...doc.lastCounters].sort(([a], [b]) => (a < b ? -1 : 1));

/** Each copy equals the room, with the room's counters; returns the person's own counter. */
function counterAfter(room: OrderedLocalRoom, person: Person, ...others: Person[]): number {
  for (const copy of [person, ...others]) {
    const doc = copy.handle.session.orderedDocument();
    expect({ person: copy.name, same: sameAsRoom(copy, room), refusals: copy.refusals }).toEqual({
      person: copy.name,
      same: true,
      refusals: [],
    });
    expect({ person: copy.name, counters: countersOf(doc) }).toEqual({
      person: copy.name,
      counters: countersOf(room.room.document),
    });
  }

  return room.room.document.lastCounters.get(person.handle.session.namespace) ?? 0;
}

describe('the counters of real sessions', () => {
  test('typing, undo and redo raise the counter; every copy holds the room counters', async () => {
    const room = localRoom();
    const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana');

    caretAt(ana, 0, 4);
    await typeKeys(ana, 'abc');
    await settle();
    const typed = counterAfter(room, ana, org);
    expect(typed).toBeGreaterThan(0);
    expect(editorTexts(org)[0]).toBe('The abcfee is 100 USD.');

    surface(ana).undo();
    await settle();
    const undone = counterAfter(room, ana, org);
    expect(undone).toBeGreaterThan(typed);
    expect(editorTexts(org)[0]).not.toBe('The abcfee is 100 USD.');
    expect(editorTexts(org)).toEqual(editorTexts(ana));

    surface(ana).redo();
    await settle();
    const redone = counterAfter(room, ana, org);
    expect(redone).toBeGreaterThan(undone);
    expect(editorTexts(org)[0]).toBe('The abcfee is 100 USD.');

    // The organiser's direct edit and its undo count in the organiser's own namespace
    caretAt(org, 0, 0);
    await typeKeys(org, 'X');
    await settle();
    const direct = counterAfter(room, org, ana);
    surface(org).undo();
    await settle();
    expect(counterAfter(room, org, ana)).toBeGreaterThan(direct);
    expect(room.room.document.lastCounters.get(ana.handle.session.namespace)).toBe(redone);
  });
});
