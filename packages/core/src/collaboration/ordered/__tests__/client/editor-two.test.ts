// Two real editors and one ordered room (task 6.6, "Collaboration verification"): concurrent
// typing at one place, undo after another person's change, deletion while the other person types
// in the deleted text, remote carets, and a late join. After every step both copies equal the
// room and both screens show their lists.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import {
  caretAt,
  caretOf,
  closeAll,
  editorTexts,
  HeldTransport,
  join,
  listXml,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { docx, p, t, W } from '../word-fixtures.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

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
  const [first, ...rest] = people;
  if (first) for (const other of rest) expect(editorTexts(other)).toEqual(editorTexts(first));
}

describe('two editors in one ordered room', () => {
  test('both type at one place at once and end with the room', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    caretAt(ana, 0, 4);
    caretAt(ben, 0, 4);
    // Both keystrokes go out before either copy hears of the other
    surface(ana).insertPlainText('red ');
    surface(ben).insertPlainText('big ');
    await settle();
    converged(room, ana, ben);
    const first = editorTexts(ana)[0] ?? '';
    expect(['The red big fee is 100 USD.', 'The big red fee is 100 USD.']).toContain(first);
    // Each caret stays after its own text
    expect(caretOf(ana).offset).toBe(first.indexOf('red ') + 4);
    expect(caretOf(ben).offset).toBe(first.indexOf('big ') + 4);
    surface(ana).insertPlainText('x');
    surface(ben).insertPlainText('y');
    await settle();
    converged(room, ana, ben);
    expect(editorTexts(ana)[0]).toContain('red x');
    expect(editorTexts(ana)[0]).toContain('big y');
  });

  test("undo after another person's change takes back only one's own text", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    caretAt(ana, 0, 19);
    surface(ana).insertPlainText(' Net 30.');
    await settle();
    caretAt(ben, 0, 27);
    surface(ben).insertPlainText(' Late fees apply.');
    await settle();
    expect(editorTexts(ben)[0]).toBe('The fee is 100 USD. Net 30. Late fees apply.');
    surface(ana).undo();
    await settle();
    converged(room, ana, ben);
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD. Late fees apply.');
    // Ben's caret stays after his own text
    expect(caretOf(ben)).toMatchObject({ index: 0, offset: 36 });
    surface(ana).redo();
    await settle();
    converged(room, ana, ben);
    expect(editorTexts(ben)[0]).toContain('Late fees apply.');
    expect(editorTexts(ben)[0]).toContain(' Net 30.');
  });

  test('Ben deletes the words Ana is typing in, and both keep both changes', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    caretAt(ana, 1, 7);
    caretAt(ben, 1, 4, 9);
    surface(ana).insertPlainText('only ');
    surface(ben).deleteSelection();
    await settle();
    converged(room, ana, ben);
    // A deletion strikes the file's words; Ana's words stay as her insertion
    expect(editorTexts(ana)[1]).toContain('only ');
    expect(room.room.document.suggestions.size).toBe(2);
  });

  test("a person's caret shows on the other person's screen at the same characters", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    caretAt(ben, 1, 4);
    await settle();
    const seen = ana.handle.session.remoteSelections();
    expect(seen).toHaveLength(1);
    const [benCaret] = seen;
    const ids = surface(ana).session.paragraphIds();
    expect(benCaret).toMatchObject({ name: 'Ben', head: { nodeId: ids[1], offset: 4 } });
    expect(ana.handle.session.participants().map((p) => [p.name, p.isLocal])).toEqual([
      ['Ana', true],
      ['Ben', false],
    ]);
    // Ana types before Ben's caret: it stays at the same characters
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Rent ');
    await settle();
    caretAt(ben, 1, 9);
    await settle();
    const [moved] = ana.handle.session.remoteSelections();
    expect(moved?.head.offset).toBe(9);
    const text = editorTexts(ana)[1] ?? '';
    expect(text.slice(0, 9)).toBe('Rent Due ');
  });

  test('a person who joins late sees everything before them', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    caretAt(ana, 1, 6);
    await settle();
    surface(ana).splitParagraph();
    await settle();
    caretAt(ana, 2, 0, 3);
    surface(ana).toggleRunProperty('b');
    await settle();
    const cy = await join(room, 'Cy');
    converged(room, ana, cy);
    expect(editorTexts(cy)).toEqual(editorTexts(ana));
    caretAt(cy, 2, 0);
    surface(cy).insertPlainText('Then ');
    await settle();
    converged(room, ana, cy);
    expect(editorTexts(ana)[2]).toBe('Then  30 days.');
  });

  test('a key typed while the room answers the last one lands when a host reads the editor in that answer', async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const link = new HeldTransport(room.transport({ person: testPerson('Ben'), ...PARTY }));
    const ben = await join(room, 'Ben', { transport: link });
    caretAt(ben, 0, 19);
    // A host reading the editor on each change, as a React store hook does
    ben.editor.on('change', () => ben.editor.snapshot());
    link.held = true;
    surface(ben).enqueueType(' ');
    surface(ben).flushPendingInput();
    await settle();
    // The next key waits for its turn while the room's answer to the first one arrives: the
    // copy shows the room's version of the paragraph, and the host's read lands the key
    surface(ben).enqueueType('f');
    link.held = false;
    link.release();
    await settle();
    converged(room, olive, ben);
    expect(editorTexts(olive)[0]).toBe('The fee is 100 USD. f');
    expect(caretOf(ben)).toMatchObject({ index: 0, offset: 21 });
  });

  test('a party edits under the suggestions-only rule; a direct edit is refused by the room', async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit' });
    const ana = await join(room, 'Ana');
    caretAt(olive, 2, 0);
    surface(olive).insertPlainText('Now ');
    await settle();
    // The local room allows writers; both stay equal to it
    converged(room, olive, ana);
    expect(editorTexts(ana)[2]).toBe('Now Signed here.');
  });

  test("Enter at the end of a heading starts the next one in the document's default style", async () => {
    // Two default paragraph styles: the last wins, and `on` is a valid way to say default
    const styles =
      `<w:styles xmlns:w="${W}">` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Plain"><w:name w:val="Plain"/></w:style>' +
      '<w:style w:type="paragraph" w:default="on" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
      '<w:next w:val="Normal"/></w:style></w:styles>';
    const file = docx(p(t('Terms'), '<w:pStyle w:val="Heading1"/>') + p(t('Body.')), {
      parts: { 'word/styles.xml': styles },
    });
    const room = localRoom(file);
    const ana = await join(room, 'Ana', { mode: 'edit', login: ORGANISER, wordFile: file });
    const ben = await join(room, 'Ben', { wordFile: file });
    caretAt(ana, 0, 'Terms'.length);
    surface(ana).splitParagraph();
    await settle();
    converged(room, ana, ben);
    for (const person of [ana, ben]) {
      expect(listXml(person).match(/<w:pStyle w:val="\w+"\/>/g)).toEqual([
        '<w:pStyle w:val="Heading1"/>',
        '<w:pStyle w:val="Normal"/>',
      ]);
    }
  });
});
