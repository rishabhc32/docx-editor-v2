// Typing inside a text box through real editors and one room (group 9.9): a party's typing,
// Enter and Backspace there are suggestions in the text box's own story; the organiser types
// directly and takes it back with undo; two people type in one text box at once; every copy
// ends as the room, and the screen follows the list.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { storyParagraphs } from '../../../../store/package/story-blocks.ts';
import { textboxStoriesInPart } from '../../../../store/package/textbox-stories.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedViewParagraph } from '../../compose/compose.ts';
import { isTextBoxStoryName } from '../../model/text-boxes.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS } from '../word-fixtures.ts';
import {
  BODY,
  closeAll,
  FILE as PLAIN,
  editorXml,
  HeldTransport,
  join,
  listTexts,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'textbox');
if (!item) throw new Error('no text box item');
const FILE = docx(bodyOf(item), MATRIX_EXTRAS);

/** The text box story's name in a copy. */
function boxOf(person: Person): string {
  const names = [...person.handle.session.orderedDocument().stories.keys()].filter(
    isTextBoxStoryName
  );
  expect(names).toHaveLength(1);
  return names[0] ?? '';
}

/** Enter the text box at its text, as Find's frame scope does. */
function enterBox(person: Person, at = 'Box text'): void {
  const match = person.editor.findMatches(at)[0];
  if (!match?.scope || match.scope.kind !== 'frame') throw new Error('no text box match');
  expect(surface(person).setActiveScope(match.scope)).toBe(true);
}

/** Put the caret in the text box's paragraph `index` at `offset`. */
function caretInBox(person: Person, index: number, offset: number): void {
  const s = surface(person);
  const [story] = textboxStoriesInPart(s.session.part());
  const ids = story ? storyParagraphs(story.root).map((paragraph) => paragraph.id) : [];
  const paragraphId = ids[index];
  if (paragraphId === undefined) throw new Error(`no text box paragraph ${index}`);
  s.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

const type = (person: Person, text: string) => surface(person).insertPlainText(text);

function checked(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(editorXml(person)).toBe(listXml(person));
    expect(sameAsRoom(person, room)).toBe(true);
  }
}

async function two(): Promise<{ room: OrderedLocalRoom; ana: Person; olive: Person }> {
  const room = localRoom(FILE);
  const ana = await join(room, 'Ana', { wordFile: FILE });
  const olive = await join(room, 'Olive', { wordFile: FILE, mode: 'edit', login: ORGANISER });
  return { room, ana, olive };
}

describe('typing inside a text box (group 9.9)', () => {
  test("a party's typing, Enter and Backspace are suggestions in the text box", async () => {
    const { room, ana, olive } = await two();
    const box = boxOf(ana);
    enterBox(ana);
    caretInBox(ana, 0, 3);
    type(ana, ' new');
    await settle();
    expect(room.room.order).toBeGreaterThan(0);
    const [first] = room.room.joinSnapshot(0);
    expect(first && 'part' in first ? first.part : null).toBe(box);
    expect(listTexts(ana, box)).toEqual(['Box new text']);
    expect(listTexts(olive, box)).toEqual(['Box new text']);
    checked(room, ana, olive);
    expect(editorXml(olive)).toMatch(/<w:txbxContent>[\s\S]*<w:ins [^>]*w:author="Ana"/);
    caretInBox(ana, 0, 7);
    surface(ana).splitParagraph();
    await settle();
    expect(listTexts(olive, box)).toEqual(['Box new', ' text']);
    checked(room, ana, olive);
    caretInBox(ana, 1, 0);
    surface(ana).deleteBackward();
    await settle();
    expect(listTexts(olive, box)).toEqual(['Box new text']);
    checked(room, ana, olive);
  });

  test("the organiser's direct typing, its undo and its redo", async () => {
    const { room, ana, olive } = await two();
    const box = boxOf(olive);
    const before = listXml(olive);
    enterBox(olive);
    caretInBox(olive, 0, 0);
    type(olive, 'My ');
    await settle();
    expect(listTexts(ana, box)).toEqual(['My Box text']);
    expect(editorXml(ana)).not.toMatch(/<w:txbxContent>[\s\S]*<w:ins /);
    checked(room, ana, olive);
    olive.editor.exec({ type: 'undo' });
    await settle();
    expect(listXml(olive)).toBe(before);
    expect(listTexts(ana, box)).toEqual(['Box text']);
    checked(room, ana, olive);
    olive.editor.exec({ type: 'redo' });
    await settle();
    expect(listTexts(ana, box)).toEqual(['My Box text']);
    checked(room, ana, olive);
  });

  test('two people type in one text box at once, and both copies end as the room', async () => {
    const room = localRoom(FILE);
    const held = new HeldTransport(
      room.transport({
        person: testPerson('Ben'),
        suggests: true,
        decides: false,
        readOnly: false,
        comments: true,
      })
    );
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const ben = await join(room, 'Ben', { wordFile: FILE, transport: held });
    const box = boxOf(ana);
    held.held = true;
    enterBox(ana);
    caretInBox(ana, 0, 0);
    type(ana, 'A');
    enterBox(ben);
    caretInBox(ben, 0, 8);
    type(ben, '!');
    await settle();
    held.held = false;
    held.release();
    await settle();
    expect(listTexts(ana, box)).toEqual(['ABox text!']);
    expect(listTexts(ben, box)).toEqual(['ABox text!']);
    checked(room, ana, ben);
  });

  test('a save holds the suggestion in both copies of the text box, and reopens', async () => {
    const { ana, olive } = await two();
    enterBox(ana);
    caretInBox(ana, 0, 0);
    type(ana, 'The ');
    await settle();
    for (const person of [ana, olive]) {
      const bytes = new Uint8Array(await person.editor.save());
      const read = readOoxmlPackage(bytes, {});
      if (!read.ok) throw new Error(read.reason);
      const saved = serializeOoxmlPart(read.package.parts.get(BODY)!);
      // The DrawingML copy and the VML fallback copy hold it, as Word writes both
      expect(
        saved.match(/<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t xml:space="preserve">The /g)
      ).toHaveLength(2);
      const reopened = openOrderedWordFile(bytes, { file: 2 });
      const box = [...reopened.document.stories.keys()].find(isTextBoxStoryName) ?? '';
      expect(reopened.document.suggestions.size).toBe(1);
      expect(
        orderedViewParagraph(reopened.document, box, 0)
          ?.chars.map((c) => c.ch)
          .join('')
      ).toBe('The Box text');
    }
  });
});

describe("a party's new text box (groups 9.4 and 9.9)", () => {
  /** Enter the only text box of the editor's main part, as a click on it does. */
  function enterNewBox(person: Person): void {
    const [story] = textboxStoriesInPart(surface(person).session.part());
    if (!story) throw new Error('no text box');
    const scope = {
      kind: 'frame' as const,
      id: story.root.id,
      drawingNodeId: story.drawingNodeId,
      hostParagraphId: story.hostParagraphId,
    };
    expect(surface(person).setActiveScope(scope)).toBe(true);
  }

  /** The party suggests a text box in the first paragraph and types in it. */
  async function suggestBox(ana: Person): Promise<void> {
    const s = surface(ana);
    const paragraphId = s.session.paragraphIds()[0] ?? '';
    s.setSelection({ anchor: { paragraphId, offset: 3 }, head: { paragraphId, offset: 3 } });
    expect(ana.editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
    await settle();
    enterNewBox(ana);
    caretInBox(ana, 0, 0);
    type(ana, 'Party text');
    await settle();
  }

  /** The party's pending suggestions in the room, in the order they were made. */
  const pendingOf = (room: OrderedLocalRoom, author: string): string[] =>
    [...room.room.document.suggestions.values()]
      .filter((s) => s.status === 'pending' && s.attribution.author === author)
      .map((s) => s.id);

  /** A copy's save, read back: the main part's XML, and the copy the file opens to. */
  async function saved(person: Person) {
    const bytes = new Uint8Array(await person.editor.save());
    const read = readOoxmlPackage(bytes, {});
    if (!read.ok) throw new Error(read.reason);
    const reopened = openOrderedWordFile(bytes, { file: 2 }).document;
    const box = [...reopened.stories.keys()].find(isTextBoxStoryName);
    const text =
      box === undefined
        ? null
        : orderedViewParagraph(reopened, box, 0)
            ?.chars.map((c) => c.ch)
            .join('');
    return { xml: serializeOoxmlPart(read.package.parts.get(BODY)!), reopened, text };
  }

  test('typed, then rejected and accepted on both copies, saved and opened again', async () => {
    const room = localRoom(PLAIN, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: PLAIN });
    const olive = await join(room, 'Olive', { wordFile: PLAIN, mode: 'edit', login: ORGANISER });
    await suggestBox(ana);
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      const xml = listXml(person);
      expect(xml).toMatch(/<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:drawing[ >][\s\S]*<wps:txbx>/);
      expect(xml).toMatch(
        /<w:txbxContent><w:p[^>]*><w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t>Party text<\/w:t>/
      );
    }
    // The organiser rejects the text box: it goes with what the party typed in it
    olive.handle.session.requestDecision(pendingOf(room, 'Ana')[0] ?? '', 'rejected');
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) expect(listXml(person)).not.toContain('<wps:txbx>');
    // Suggested again: a save while it is pending holds both suggestions, and opens again
    await suggestBox(ana);
    checked(room, ana, olive);
    const pending = await saved(olive);
    expect(pending.xml).toMatch(/<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:drawing[ >]/);
    expect(pending.reopened.suggestions.size).toBe(2);
    expect(pending.text).toBe('Party text');
    // The organiser accepts both: plain content on every copy, and in each save
    for (const id of pendingOf(room, 'Ana')) olive.handle.session.requestDecision(id, 'accepted');
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      const xml = listXml(person);
      expect(xml).toMatch(
        /<w:txbxContent><w:p w14:paraId="[0-9A-F]{8}"[^>]*><w:r><w:t>Party text<\/w:t><\/w:r><\/w:p>/
      );
      expect(xml).not.toContain('<w:ins');
      const accepted = await saved(person);
      expect(accepted.xml).not.toContain('<w:ins');
      expect(accepted.reopened.suggestions.size).toBe(0);
      expect(accepted.text).toBe('Party text');
    }
  });
});
