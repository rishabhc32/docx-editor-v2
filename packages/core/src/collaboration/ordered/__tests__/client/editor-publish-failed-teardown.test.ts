// A copy whose list cannot be composed (an engine defect, as an acceptance run met at round 41:
// `RangeError: An inline drawing without a box` in compose) ends its session with
// `remote-apply-failed`. The page must still close it: before, the editor's detach and the
// session's `destroy` composed the same paragraphs again and threw each time, so every new
// join of the page failed at once (eight in a row in about a second). No list on this base
// fails to compose, so the test makes compose throw for one copy while it is armed.

import { afterEach, describe, expect, mock, test } from 'bun:test';
import * as compose from '../../compose/compose.ts';
import type { CollaborationStatus } from '../../../failure.ts';

// Captured before mocking: the namespace binding is live and would point at the wrapper
const composeOrderedParagraph = compose.composeOrderedParagraph;
const composeOrderedPart = compose.composeOrderedPart;
const original = { ...compose };
/** The copy whose list does not compose, while armed. */
let broken: unknown = null;
const defect = () => new RangeError('An inline drawing without a box.');
mock.module('../../compose/compose.ts', () => ({
  ...original,
  composeOrderedParagraph: (...args: Parameters<typeof composeOrderedParagraph>) => {
    if (args[0] === broken) throw defect();
    return composeOrderedParagraph(...args);
  },
  composeOrderedPart: (...args: Parameters<typeof composeOrderedPart>) => {
    if (args[0] === broken) throw defect();
    return composeOrderedPart(...args);
  },
}));

const {
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
} = await import('./editor-harness.ts');
type Person = Awaited<ReturnType<typeof join>>;

afterEach(() => {
  broken = null;
  closeAll();
});

async function key(person: Person, paragraph: number, text: string): Promise<void> {
  const length = editorTexts(person)[paragraph]?.length ?? 0;
  caretAt(person, paragraph, length);
  surface(person).insertPlainText(text);
  await settle(2);
}

describe('a copy that cannot compose its list', () => {
  test('ends with remote-apply-failed, and the page can still close it and join again', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    const statuses: [CollaborationStatus, string | undefined, string | undefined][] = [];
    ana.handle.session.subscribeStatus((status, reason, detail) =>
      statuses.push([status, reason, detail])
    );
    await key(ana, 0, ' One.');
    broken = ana.handle.session.orderedDocument();
    // Ben's change reaches Ana's copy, which cannot show it
    await key(ben, 1, ' Ben.');
    await settle();
    expect(ana.handle.session.status()).toBe('error');
    expect(ana.handle.session.statusSnapshot().reason?.code).toBe('remote-apply-failed');

    // The page closes the editor and the session, as a host does before it joins again
    expect(() => ana.editor.destroy()).not.toThrow();
    expect(() => ana.handle.destroy()).not.toThrow();
    expect(ana.handle.session.status()).toBe('destroyed');
    expect(statuses.map(([status, reason]) => [status, reason])).toEqual([
      ['error', 'remote-apply-failed'],
      ['destroyed', undefined],
    ]);

    // Nothing of Ana's was lost: her change reached the room before the copy failed
    broken = null;
    const again = await join(room, 'Ana');
    expect(again.handle.session.status()).toBe('ready');
    expect(sameAsRoom(again, room)).toBe(true);
    expect(screenFollowsList(again)).toBe(true);
    expect(editorTexts(again)).toEqual([
      'The fee is 100 USD. One.',
      'Due in 30 days. Ben.',
      'Signed here.',
    ]);
  });

  test('heals when a later publish goes out: the paragraphs it could not show are shown then', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    broken = ana.handle.session.orderedDocument();
    await key(ben, 1, ' Ben.');
    await settle();
    expect(ana.handle.session.statusSnapshot().reason).toEqual({
      code: 'remote-apply-failed',
      detail: 'compose: RangeError',
    });
    expect(editorTexts(ana)[1]).toBe('Due in 30 days.');
    broken = null;
    await key(ben, 2, ' Later.');
    await settle();
    expect(ana.handle.session.status()).toBe('ready');
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(screenFollowsList(ana)).toBe(true);
    expect(editorTexts(ana)).toEqual([
      'The fee is 100 USD.',
      'Due in 30 days. Ben.',
      'Signed here. Later.',
    ]);
  });
});
