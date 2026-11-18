// A change that throws something other than a refusal when the person's own copy applies it (an
// engine defect) leaves nothing of its batch behind: the changes applied before it in the batch
// are taken back and never sent, and an undo puts its step back. No change on this base throws
// there, so the test makes one apply throw.

import { afterEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { OrderedReplica } from '../../client/replica.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(() => {
  mock.restore();
  closeAll();
});

/** Until restored, the person's `nth` apply from now throws a defect. */
function throwOnApply(nth: number) {
  const applyLocal = OrderedReplica.prototype.applyLocal;
  let calls = 0;
  return spyOn(OrderedReplica.prototype, 'applyLocal').mockImplementation(function (
    this: OrderedReplica,
    change
  ) {
    calls += 1;
    if (calls === nth) {
      throw new RangeError('An engine defect.');
    }
    return applyLocal.call(this, change);
  });
}

/** Type over "fee" in the first paragraph: a deletion and an insertion, one batch. */
function typeOverFee(person: Person): void {
  caretAt(person, 0, 4, 7);
  surface(person).insertPlainText('sum');
}

/** Nothing pending, the copy is the room, and the screen is the list. */
function clean(room: ReturnType<typeof localRoom>, person: Person): void {
  expect(person.handle.session.unsentChanges()).toEqual([]);
  expect(sameAsRoom(person, room)).toBe(true);
  expect(screenFollowsList(person)).toBe(true);
}

describe("a defect while a person's own change applies", () => {
  test('in an edit: the changes before it are taken back and never sent', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    const defect = throwOnApply(2);
    expect(() => typeOverFee(ana)).toThrow('An engine defect.');
    defect.mockRestore();
    await settle();
    expect(room.room.order).toBe(0);
    clean(room, ana);
    expect(editorTexts(ben)[0]).toBe('The fee is 100 USD.');
  });

  test('in an undo: nothing is sent, and the step stays to undo', async () => {
    const room = localRoom();
    const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER });
    typeOverFee(org);
    await settle();
    const order = room.room.order;
    // The undo is one change, a `revert`
    const defect = throwOnApply(1);
    expect(() => surface(org).undo()).toThrow('An engine defect.');
    defect.mockRestore();
    await settle();
    expect(room.room.order).toBe(order);
    clean(room, org);
    expect(org.handle.session.canUndo()).toBe(true);
    surface(org).undo();
    await settle();
    clean(room, org);
    expect(editorTexts(org)[0]).toBe('The fee is 100 USD.');
  });

  test('in a redo: nothing is sent', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    typeOverFee(ana);
    await settle();
    surface(ana).undo();
    await settle();
    const order = room.room.order;
    const defect = throwOnApply(2);
    expect(() => surface(ana).redo()).toThrow('An engine defect.');
    defect.mockRestore();
    await settle();
    expect(room.room.order).toBe(order);
    clean(room, ana);
  });
});
