// A pending change that throws something other than a refusal when the client applies it again
// (an engine defect, as an acceptance run met at round 41: `RangeError: An inline drawing without a
// box`) is refused by name and taken back; the changes that need it are refused too, the
// others are sent, and the session stays ready. Before, the session ended with
// `remote-apply-failed` and kept the change, so the person could not join again without losing
// every unsent edit. No change on this base throws there, so the test makes one change id throw
// once it is armed, wherever it applies: on the copy and in the room.

import { afterEach, describe, expect, mock, test } from 'bun:test';
import * as apply from '../../change/apply.ts';
import type { OrderedChange } from '../../change/types.ts';
import type { CollaborationStatus } from '../../../failure.ts';
import { testPerson } from '../test-persons.ts';

// Captured before mocking: the namespace binding is live and would point at the wrapper
const applyOrderedChange = apply.applyOrderedChange;
const original = { ...apply };
/** The change id that throws, once armed. */
let failing: string | null = null;
mock.module('../../change/apply.ts', () => ({
  ...original,
  applyOrderedChange: (...args: Parameters<typeof applyOrderedChange>) => {
    if (args[1].id === failing) throw new RangeError('An inline drawing without a box.');
    return applyOrderedChange(...args);
  },
}));

const {
  caretAt,
  closeAll,
  editorTexts,
  HeldTransport,
  join,
  localRoom,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
} = await import('./editor-harness.ts');
type Person = Awaited<ReturnType<typeof join>>;

afterEach(() => {
  failing = null;
  closeAll();
});

/** One keystroke run at the end of a paragraph: one change, one user action. */
async function key(person: Person, paragraph: number, text: string): Promise<void> {
  const length = editorTexts(person)[paragraph]?.length ?? 0;
  caretAt(person, paragraph, length);
  surface(person).insertPlainText(text);
  await settle(2);
}

/** The person's pending change ids, in order. */
const pendingIds = (person: Person): string[] =>
  person.handle.session.unsentChanges().map((change) => change.id);

/** The person joins and logs every status and every defect the client reports. */
async function watched(room: ReturnType<typeof localRoom>, name: string) {
  const errors: { readonly error: unknown; readonly change: string }[] = [];
  const person = await join(room, name, {
    onApplyError: (error: unknown, change: OrderedChange) =>
      errors.push({ error, change: change.id }),
  });
  const statuses: CollaborationStatus[] = [];
  person.handle.session.subscribeStatus((status) => statuses.push(status));
  return { person, errors, statuses };
}

function inStep(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ who: person.name, status: person.handle.session.status() }).toEqual({
      who: person.name,
      status: 'ready',
    });
    expect({ who: person.name, room: sameAsRoom(person, room) }).toEqual({
      who: person.name,
      room: true,
    });
    expect({ who: person.name, screen: screenFollowsList(person) }).toEqual({
      who: person.name,
      screen: true,
    });
  }
}

describe('a pending change that throws when it is applied again', () => {
  test('after a reconnect: refused by name, the others sent, the session ready (no lock-out)', async () => {
    const room = localRoom();
    const { person: ana, errors, statuses } = await watched(room, 'Ana');
    const ben = await join(room, 'Ben');
    room.disconnect(ana.handle.session.namespace);
    await settle();
    await key(ana, 0, ' One.');
    await key(ana, 1, ' Two.');
    await key(ana, 2, ' Three.');
    await key(ben, 1, ' Ben.');
    const [one, two, three] = pendingIds(ana);
    expect(pendingIds(ana)).toHaveLength(3);
    failing = two ?? null;
    ana.handle.session.reconnect();
    await settle(12);

    // Before: `error`, `remote-apply-failed`, and no way back in with the unsent changes
    expect(ana.handle.session.statusSnapshot().reason).toBeUndefined();
    expect(statuses).not.toContain('error');
    expect(ana.refusals).toEqual([
      { code: 'ordered-refused', reason: 'apply-failed', detail: 'RangeError', change: two },
    ]);
    expect(errors.map(({ error, change }) => [String(error), change])).toEqual([
      ['RangeError: An inline drawing without a box.', two],
    ]);
    inStep(room, ana, ben);
    // The others reached the room in their order; the refused one never left the copy
    expect(room.room.joinSnapshot(0).map((change) => change.id)).toEqual([
      expect.any(String),
      one,
      three,
    ]);
    expect(editorTexts(ben)).toEqual([
      'The fee is 100 USD. One.',
      'Due in 30 days. Ben.',
      'Signed here. Three.',
    ]);
    expect(editorTexts(ana)).toEqual(editorTexts(ben));
    // The session goes on: the next edit is sent and ordered
    await key(ana, 0, ' Four.');
    await settle();
    inStep(room, ana, ben);
    expect(editorTexts(ben)[0]).toBe('The fee is 100 USD. One. Four.');
  });

  test('a change that needs the failed one is refused by name too; the rest are kept', async () => {
    const room = localRoom();
    const { person: ana, errors } = await watched(room, 'Ana');
    const ben = await join(room, 'Ben');
    room.disconnect(ana.handle.session.namespace);
    await settle();
    await key(ana, 0, ' One.');
    await key(ana, 1, ' Two.');
    // Typed inside the text the failed change inserts: it needs that change
    caretAt(ana, 1, 'Due in 30 days. Tw'.length);
    surface(ana).insertPlainText('o-o');
    await settle(2);
    await key(ana, 2, ' Four.');
    await key(ben, 1, ' Ben.');
    const [one, two, inside, four] = pendingIds(ana);
    expect(pendingIds(ana)).toHaveLength(4);
    failing = two ?? null;
    ana.handle.session.reconnect();
    await settle(12);

    expect(ana.refusals).toEqual([
      { code: 'ordered-refused', reason: 'apply-failed', detail: 'RangeError', change: two },
      expect.objectContaining({
        code: 'ordered-refused',
        reason: 'pending-refused',
        change: inside,
      }),
    ]);
    expect(errors.map(({ change }) => change)).toEqual([two]);
    inStep(room, ana, ben);
    expect(room.room.joinSnapshot(0).map((change) => change.id)).toEqual([
      expect.any(String),
      one,
      four,
    ]);
    expect(editorTexts(ben)).toEqual([
      'The fee is 100 USD. One.',
      'Due in 30 days. Ben.',
      'Signed here. Four.',
    ]);
    expect(editorTexts(ana)).toEqual(editorTexts(ben));
  });

  test("a sent change the room ordered is kept by the room's order when a later unsent one fails", async () => {
    const room = localRoom();
    const held = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
    const errors: string[] = [];
    const ana = await join(room, 'Ana', {
      transport: held,
      onApplyError: (_error: unknown, change: OrderedChange) => errors.push(change.id),
    });
    const ben = await join(room, 'Ben');
    // Ana's first change is sent and ordered, but the room's answer does not reach her copy
    // before the connection drops
    held.held = true;
    await key(ana, 0, ' One.');
    await settle();
    const [one] = pendingIds(ana);
    expect(room.room.order).toBe(1);
    room.disconnect(ana.handle.session.namespace);
    await settle();
    held.held = false;
    await key(ana, 2, ' Two.');
    await key(ben, 1, ' Ben.');
    const [, two] = pendingIds(ana);
    expect(pendingIds(ana)).toEqual([one, two]);
    failing = two ?? null;
    ana.handle.session.reconnect();
    await settle(12);

    expect(ana.refusals).toEqual([
      { code: 'ordered-refused', reason: 'apply-failed', detail: 'RangeError', change: two },
    ]);
    expect(errors).toEqual([two]);
    inStep(room, ana, ben);
    expect(room.room.joinSnapshot(0).map((change) => change.id)).toEqual([one, expect.any(String)]);
    expect(editorTexts(ana)).toEqual([
      'The fee is 100 USD. One.',
      'Due in 30 days. Ben.',
      'Signed here.',
    ]);
    expect(editorTexts(ben)).toEqual(editorTexts(ana));
  });

  test("after another person's change: a sent change that throws is refused, and the room refuses it too", async () => {
    const room = localRoom();
    const held = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
    const errors: string[] = [];
    const ana = await join(room, 'Ana', {
      transport: held,
      onApplyError: (_error: unknown, change: OrderedChange) => errors.push(change.id),
    });
    const ben = await join(room, 'Ben');
    const statuses: CollaborationStatus[] = [];
    ana.handle.session.subscribeStatus((status) => statuses.push(status));
    held.held = true;
    await key(ben, 1, ' Ben.');
    caretAt(ana, 0, editorTexts(ana)[0]?.length ?? 0);
    surface(ana).insertPlainText(' One.');
    const [one] = pendingIds(ana);
    // Armed before the change is sent: the room applies Ben's change, then Ana's, which throws
    // there too
    failing = one ?? null;
    await settle();
    expect(room.room.order).toBe(1);
    held.held = false;
    held.release();
    await settle(12);

    expect(statuses).not.toContain('error');
    // One report for the one user action: the copy's own, which names the change
    expect(ana.refusals).toEqual([
      { code: 'ordered-refused', reason: 'apply-failed', detail: 'RangeError', change: one },
    ]);
    expect(errors).toEqual([one]);
    inStep(room, ana, ben);
    expect(editorTexts(ana)).toEqual([
      'The fee is 100 USD.',
      'Due in 30 days. Ben.',
      'Signed here.',
    ]);
  });
});
