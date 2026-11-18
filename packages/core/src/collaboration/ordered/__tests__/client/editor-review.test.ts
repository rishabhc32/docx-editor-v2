// Decisions, comments and the saved file (task 6.5, "Collaboration verification"): an organiser
// decides a party's card through the room; a card that changed after the organiser saw it is
// refused as `stale-decision` and nothing changes; a comment goes through the room;
// and what the editors show survives save and reopen as the room's own document.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
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
} from './editor-harness.ts';

afterEach(closeAll);

const fingerprints = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].map((name) =>
    canonicalOoxmlFingerprint(composeOrderedPart(doc, name).part)
  );

describe('decisions, comments and the saved file', () => {
  test("an organiser accepts a party's card through the room", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    const [card] = orderedCards(olive.handle.session.orderedDocument());
    const first = card?.suggestions[0];
    if (first === undefined) throw new Error('no card');
    expect(olive.handle.session.requestDecision(first, 'accepted')).not.toBeNull();
    await settle();
    for (const person of [ana, olive]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorXml(person)).not.toContain('<w:ins ');
    }
    expect(editorTexts(ana)[0]).toBe('The new fee is 100 USD.');
    expect(room.room.joinSnapshot(1)[0]?.kind).toBe('decide');
  });

  test('a card that changed after the organiser saw it is refused as stale, and nothing changes', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    const [card] = orderedCards(olive.handle.session.orderedDocument());
    const first = card?.suggestions[0];
    if (first === undefined) throw new Error('no card');
    // Ana types into her suggestion; the room orders it before Olive's decision
    surface(ana).insertPlainText('x');
    ana.handle.session.flushPendingJournals();
    const request = olive.handle.session.requestDecision(first, 'accepted');
    await settle();
    expect(olive.refusals).toEqual([
      expect.objectContaining({
        code: 'ordered-refused',
        reason: 'stale-decision',
        change: request,
      }),
    ]);
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorXml(person)).toContain('<w:ins ');
    }
    expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert', 'insert']);
  });

  test('a comment goes through the room', async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana');
    const id = surface(olive).session.paragraphIds()[1];
    if (id === undefined) throw new Error('no paragraph');
    expect(
      olive.handle.session.addComment('Why?', { nodeId: id, offset: 0 }, { nodeId: id, offset: 3 })
    ).not.toBeNull();
    await settle();
    expect(room.room.joinSnapshot(0)[0]?.kind).toBe('comment');
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(editorXml(person)).toContain('<w:commentRangeStart ');
    }
  });

  test("the editors' save reopens as the room's document", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    caretAt(ben, 1, 6);
    surface(ben).splitParagraph();
    await settle();
    caretAt(ben, 2, 0, 3);
    surface(ben).toggleRunProperty('b');
    caretAt(ana, 0, 0, 3);
    surface(ana).deleteSelection();
    await settle();
    for (const person of [ana, ben]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
    }
    const expected = fingerprints(room.room.document);
    for (const person of [ana, ben]) {
      const saved = new Uint8Array(await person.editor.save());
      expect(readOoxmlPackage(saved).ok).toBe(true);
      // The room's next Word file is what an editor saves now: it opens to the same document
      const reopened = openOrderedWordFile(saved, { file: 2 });
      expect(fingerprints(reopened.document)).toEqual(expected);
    }
    // And a fresh join on that file shows the same text
    expect(editorTexts(ana)).toEqual(editorTexts(ben));
  });
});
