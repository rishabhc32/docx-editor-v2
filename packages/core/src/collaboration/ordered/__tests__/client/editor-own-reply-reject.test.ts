// The organiser's own change inside a party's pending insertion, taken back alone, with real
// editors and one room (the host's "Undo my edit", found by the live run of the 2.x flows). In
// suggesting mode her change is a reply on the party's card: the editor's reject of it is a
// `retract` of her suggestion only, as the editor alone and Word reject it, so the party's words
// stay his pending insertion. In edit mode her words join the party's insertion and have no
// revision of their own: her undo takes back exactly her words. Deciding the card still decides
// it whole. Her undo of the reject gives her change back, as in the editor alone.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
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
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** Each revision card the editor shows: its author, kind and words. */
const revisions = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .flatMap((placement) => (placement.kind === 'revision' ? [placement.item] : []))
    .map((item) => `${item.addresses[0]?.author} ${item.revisionKind} ${item.text}`);

/** The key of the first revision card of `author`. */
function cardOf(person: Person, author: string): string {
  const found = person.editor
    .getReviewItems()
    .find((placement) => placement.kind === 'revision' && placement.author === author);
  if (!found) throw new Error(`no card of ${author}`);
  return found.key;
}

const kinds = (room: OrderedLocalRoom): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

/** Ana suggests "big new " in the first paragraph; the organiser joins in `mode`. */
async function anaSuggests(mode: 'suggesting' | 'edit') {
  const room = localRoom(undefined, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { review });
  const olive = await join(room, 'Olive', { mode, login: ORGANISER, review });
  caretAt(ana, 0, 4);
  surface(ana).insertPlainText('big new ');
  await settle();
  return { room, ana, olive };
}

const ANA_ONLY = ['The big new fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

describe("the organiser's own change inside a party's insertion, taken back alone", () => {
  test("suggesting: rejecting her inserted words takes back hers only; Ana's stay pending", async () => {
    const { room, ana, olive } = await anaSuggests('suggesting');
    caretAt(olive, 0, 8);
    surface(olive).insertPlainText('very ');
    await settle();
    expect(editorTexts(ana)[0]).toBe('The big very new fee is 100 USD.');
    expect(olive.editor.rejectReviewItem(cardOf(olive, 'Olive')).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'retract']);
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorTexts(person)).toEqual(ANA_ONLY);
      expect(revisions(person)).toEqual(['Ana insert big new ']);
    }
    // Her undo gives her words back, as the editor's undo of the reject does; redo takes them back
    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(ana)[0]).toBe('The big very new fee is 100 USD.');
    expect(revisions(ana)).toEqual(['Ana insert big ', 'Olive insert very ', 'Ana insert new ']);
    expect(olive.handle.session.redo()).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(ana)).toEqual(ANA_ONLY);
    expect(revisions(ana)).toEqual(['Ana insert big new ']);
  });

  test("suggesting: rejecting her replacement of Ana's word gives Ana's word back", async () => {
    const { room, ana, olive } = await anaSuggests('suggesting');
    caretAt(olive, 0, 8, 11);
    surface(olive).insertPlainText('old');
    await settle();
    // Her deletion of "new" shows struck, beside her "old"
    expect(editorTexts(ana)[0]).toBe('The big newold fee is 100 USD.');
    // As the card's Undo my edit does: each of her changes rejected
    while (revisions(olive).some((card) => card.startsWith('Olive '))) {
      expect(olive.editor.rejectReviewItem(cardOf(olive, 'Olive')).ok).toBe(true);
      await settle();
    }
    expect(kinds(room)).toEqual(['insert', 'delete', 'insert', 'retract']);
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorTexts(person)).toEqual(ANA_ONLY);
      expect(revisions(person)).toEqual(['Ana insert big new ']);
    }
  });

  test('edit mode: her words join Ana’s insertion, and her undo takes back hers only', async () => {
    const { room, ana, olive } = await anaSuggests('edit');
    caretAt(olive, 0, 8);
    surface(olive).insertPlainText('very ');
    await settle();
    expect(revisions(olive)).toEqual(['Ana insert big very new ']);
    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorTexts(person)).toEqual(ANA_ONLY);
      expect(revisions(person)).toEqual(['Ana insert big new ']);
    }
  });

  test("rejecting Ana's card still rejects it whole, with the organiser's words in it", async () => {
    const { room, ana, olive } = await anaSuggests('suggesting');
    caretAt(olive, 0, 8);
    surface(olive).insertPlainText('very ');
    await settle();
    expect(olive.editor.rejectReviewItem(cardOf(olive, 'Ana')).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorTexts(person)[0]).toBe('The fee is 100 USD.');
      expect(revisions(person)).toEqual([]);
    }
  });
});
