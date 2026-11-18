// Review actions from the editor (task 9.3, design D15), with real editors and one room: the
// organiser's accept and reject, of one card and of all, become one `decide` change with the
// versions the organiser's copy shows; a party's is refused by the party rule (`policy-decide`);
// a card that changed first is refused as stale; a new comment becomes a `comment` change (more
// in `editor-comments-room.test.ts`); a decision has no undo; and what the editors save reopens
// as the room's.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
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

/** The keys of the revision cards an editor shows. */
const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

/** Every room change's kind, in order. */
const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

const fingerprints = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].map((name) =>
    canonicalOoxmlFingerprint(composeOrderedPart(doc, name).part)
  );

/** A party's two suggestions, in two paragraphs, seen by an organiser. */
async function twoSuggestions(policy = false) {
  const room = localRoom(undefined, policy ? { policy: suggestionsOnlyPolicy() } : {});
  const ana = await join(room, 'Ana', { review });
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
  caretAt(ana, 0, 4);
  surface(ana).insertPlainText('new ');
  await settle();
  caretAt(ana, 2, 0);
  surface(ana).insertPlainText('Not ');
  await settle();
  return { room, ana, olive };
}

describe('review actions from the editor (task 9.3)', () => {
  test('the organiser accepts one card from the editor: a decide change through the room', async () => {
    const { room, ana, olive } = await twoSuggestions();
    const [first] = cards(olive);
    if (first === undefined) throw new Error('no card');
    expect(olive.editor.acceptReviewItem(first).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)).toEqual([
        'The new fee is 100 USD.',
        'Due in 30 days.',
        'Not Signed here.',
      ]);
    }
    // The other card is still pending, on both editors
    expect(cards(olive)).toHaveLength(1);
    expect(editorXml(ana)).toContain('<w:ins ');
  });

  test('the organiser rejects one card from the editor', async () => {
    const { room, ana, olive } = await twoSuggestions();
    const [first] = cards(olive);
    if (first === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(first).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)[0]).toBe('The fee is 100 USD.');
    }
  });

  test('accept all and reject all: one decide change for every card', async () => {
    for (const action of ['accept', 'reject'] as const) {
      const { room, ana, olive } = await twoSuggestions();
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle();
      expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
      for (const person of [ana, olive]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        expect(editorXml(person)).not.toContain('<w:ins ');
        expect(editorTexts(person)).toEqual(
          action === 'accept'
            ? ['The new fee is 100 USD.', 'Due in 30 days.', 'Not Signed here.']
            : ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.']
        );
      }
      closeAll();
    }
  });

  test("a party's accept is refused by the party rule, and the card shows again", async () => {
    const { room, ana, olive } = await twoSuggestions(true);
    const [first] = cards(ana);
    if (first === undefined) throw new Error('no card');
    ana.editor.acceptReviewItem(first);
    await settle();
    expect(ana.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'policy-decide' }),
    ]);
    expect(kinds(room)).toEqual(['insert', 'insert']);
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(cards(person)).toHaveLength(2);
    }
  });

  test('a card that changed before the decision arrives is refused as stale, and shows again', async () => {
    const { room, ana, olive } = await twoSuggestions();
    const [first] = cards(olive);
    if (first === undefined) throw new Error('no card');
    // Ana types into her card; the room orders it before Olive's decision
    caretAt(ana, 0, 7);
    surface(ana).insertPlainText('x');
    ana.handle.session.flushPendingJournals();
    olive.editor.acceptReviewItem(first);
    await settle();
    // The copy finds it stale when it applies its pending decision again, and the room refuses it
    // too: one report for the one decision, naming the first cause
    expect(olive.refusals).toEqual([
      expect.objectContaining({ reason: 'pending-refused', detail: 'stale-decision' }),
    ]);
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)[0]).toBe('The newx fee is 100 USD.');
    }
    expect(cards(olive)).toHaveLength(2);
  });

  test('a decision has no undo: it is refused by name', async () => {
    const { room, olive } = await twoSuggestions();
    const [first] = cards(olive);
    if (first === undefined) throw new Error('no card');
    olive.editor.acceptReviewItem(first);
    await settle();
    expect(olive.handle.session.undo()).toBe(false);
    expect(olive.refusals).toEqual([
      expect.objectContaining({ reason: 'no-inverse', detail: 'decided' }),
    ]);
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
  });

  test("the editor's comment becomes a comment change with its text, whatever the text", async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(olive, 1, 0, 3);
    expect(olive.editor.addComment('thread:7').ok).toBe(true);
    caretAt(olive, 2, 0, 3);
    expect(olive.editor.addComment('Please check this.').ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment']);
    const texts = room.room
      .joinSnapshot(0)
      .map((change) => (change.kind === 'comment' ? change.text : ''));
    expect(texts).toEqual(['thread:7', 'Please check this.']);
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorXml(person)).toContain('<w:commentRangeStart ');
    }
  });

  test("the editors' save after decisions reopens as the room's document", async () => {
    const { room, ana, olive } = await twoSuggestions();
    const [first] = cards(olive);
    if (first === undefined) throw new Error('no card');
    olive.editor.acceptReviewItem(first);
    await settle();
    const expected = fingerprints(room.room.document);
    for (const person of [ana, olive]) {
      const saved = new Uint8Array(await person.editor.save());
      expect(readOoxmlPackage(saved).ok).toBe(true);
      const reopened = openOrderedWordFile(saved, { file: 2 });
      expect(fingerprints(reopened.document)).toEqual(expected);
    }
  });
});
