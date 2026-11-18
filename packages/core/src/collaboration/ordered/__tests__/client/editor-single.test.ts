// One real editor in suggesting mode over an ordered room (task 6.6): typing, Backspace, Enter,
// bold, plain paste, undo and redo. After each: the room ordered the change, the editor shows
// the list's version exactly, the copy equals the room, and the caret is where the edit leaves
// it. An op the change format cannot say, and a write the gate cannot see, are refused by name
// before the editor commits.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  caretAt,
  caretOf,
  closeAll,
  editorTexts,
  editorXml,
  join,
  listTexts,
  localRoom,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { docx, p, t } from '../word-fixtures.ts';
import { copiedBetween } from './paste-harness.ts';

/** A VML picture linked to a file by address, as older Word writes one. */
const VML_LINKED_SOURCE = docx(
  p(
    t('Plan ') +
      '<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office"><v:shape style="width:36pt;height:36pt">' +
      '<v:imagedata o:href="file:///C:/seal.png" o:title=""/></v:shape></w:pict></w:r>'
  )
);

afterEach(closeAll);

/** After an edit: the room has `order` changes, the screen is the list, the copy is the room. */
function checked(person: Person, room: OrderedLocalRoom, order: number): void {
  expect(person.refusals).toEqual([]);
  expect(room.room.order).toBe(order);
  expect(screenFollowsList(person)).toBe(true);
  expect(sameAsRoom(person, room)).toBe(true);
  expect(editorTexts(person)).toEqual(listTexts(person));
}

describe('one editor in an ordered room', () => {
  test('typing is one suggestion the room orders, with the caret after the text', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    // The editor attached its port to the session the collaboration module gave it
    expect(ana.handle.session.attached).toBe(true);
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    checked(ana, room, 1);
    expect(editorTexts(ana)[0]).toBe('The new fee is 100 USD.');
    expect(caretOf(ana)).toMatchObject({ index: 0, offset: 8 });
    surface(ana).insertPlainText('x');
    await settle();
    checked(ana, room, 2);
    expect(caretOf(ana)).toMatchObject({ index: 0, offset: 9 });
    // Typing beside one's own text extends one suggestion: one card
    expect([...room.room.document.suggestions.keys()]).toEqual([
      'S:' + room.room.joinSnapshot(0)[0]?.id + '.1',
    ]);
    expect(editorXml(ana)).toContain('<w:ins ');
  });

  test("Backspace strikes the file's text and takes back one's own", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 1, 3);
    surface(ana).deleteBackward();
    await settle();
    checked(ana, room, 1);
    // The struck letter stays in the text; the caret goes before it
    expect(editorTexts(ana)[1]).toBe('Due in 30 days.');
    expect(editorXml(ana)).toContain('<w:delText>e</w:delText>');
    expect(caretOf(ana)).toMatchObject({ index: 1, offset: 2 });
    caretAt(ana, 0, 3);
    surface(ana).insertPlainText('ab');
    await settle();
    surface(ana).deleteBackward();
    await settle();
    checked(ana, room, 3);
    expect(editorTexts(ana)[0]).toBe('Thea fee is 100 USD.');
    expect(caretOf(ana)).toMatchObject({ index: 0, offset: 4 });
  });

  test('Enter splits the paragraph, and the caret starts the new one', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 1, 6);
    surface(ana).splitParagraph();
    await settle();
    checked(ana, room, 1);
    expect(editorTexts(ana)).toEqual([
      'The fee is 100 USD.',
      'Due in',
      ' 30 days.',
      'Signed here.',
    ]);
    expect(caretOf(ana)).toMatchObject({ index: 2, offset: 0 });
    surface(ana).insertPlainText('X');
    await settle();
    checked(ana, room, 2);
    expect(editorTexts(ana)[2]).toBe('X 30 days.');
    expect(caretOf(ana)).toMatchObject({ index: 2, offset: 1 });
    // Backspace at the start of the new paragraph takes back the person's own Enter
    caretAt(ana, 2, 0);
    surface(ana).deleteBackward();
    await settle();
    checked(ana, room, 3);
    expect(editorTexts(ana)[1]).toBe('Due inX 30 days.');
    expect(caretOf(ana)).toMatchObject({ index: 1, offset: 6 });
  });

  test("bold formats a range, and a whole paragraph's mark with it", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 0, 4, 7);
    surface(ana).toggleRunProperty('b');
    await settle();
    checked(ana, room, 1);
    expect(caretOf(ana)).toMatchObject({ index: 0, offset: 7 });
    // The whole last paragraph: its runs and its mark, in one suggestion
    caretAt(ana, 2, 0, 12);
    surface(ana).toggleRunProperty('b');
    await settle();
    checked(ana, room, 3);
    const kinds = room.room.joinSnapshot(0).map((change) => change.kind);
    expect(kinds).toEqual(['format', 'format', 'mark-format']);
    const last = room.room.joinSnapshot(1);
    expect(last[1]).toMatchObject({ into: `S:${last[0]?.id}.1` });
    expect(editorXml(ana)).toMatch(/<w:pPr><w:rPr><w:b\/><w:rPrChange/);
    // Across two paragraphs: the mark between is formatted by the range
    caretAt(ana, 0, 10, 3, 1);
    surface(ana).toggleRunProperty('i');
    await settle();
    expect(ana.refusals).toEqual([]);
    expect(screenFollowsList(ana)).toBe(true);
    expect(sameAsRoom(ana, room)).toBe(true);
  });

  test('a plain paste of several lines types the text and its breaks', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 2, 2);
    surface(ana).insertPlainText('one\ntwo\nthree');
    await settle();
    checked(ana, room, 3);
    expect(editorTexts(ana)).toEqual([
      'The fee is 100 USD.',
      'Due in 30 days.',
      'Sione',
      'two',
      'threegned here.',
    ]);
    expect(caretOf(ana)).toMatchObject({ index: 4, offset: 5 });
    // One suggestion: the text and both breaks are one card
    const changes = room.room.joinSnapshot(0);
    expect(new Set(room.room.document.suggestions.keys())).toEqual(
      new Set([`S:${changes[0]?.id}.1`])
    );
  });

  test("undo takes back the person's last step, and redo makes it again", async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Rent ');
    await settle();
    caretAt(ana, 2, 6);
    surface(ana).undo();
    await settle();
    checked(ana, room, 3);
    expect(room.room.joinSnapshot(2)[0]?.kind).toBe('retract');
    expect(editorTexts(ana)).toEqual([
      'The new fee is 100 USD.',
      'Due in 30 days.',
      'Signed here.',
    ]);
    // The caret goes back where the undone step started
    expect(caretOf(ana)).toMatchObject({ index: 1, offset: 0 });
    surface(ana).undo();
    await settle();
    checked(ana, room, 4);
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD.');
    expect(caretOf(ana)).toMatchObject({ index: 0, offset: 4 });
    surface(ana).redo();
    await settle();
    checked(ana, room, 5);
    expect(editorTexts(ana)[0]).toBe('The new fee is 100 USD.');
    surface(ana).redo();
    await settle();
    checked(ana, room, 6);
    expect(editorTexts(ana)[1]).toBe('Rent Due in 30 days.');
    expect(caretOf(ana).index).toBe(1);
    // Nothing left to redo; undo again takes back the redone step
    expect(ana.handle.session.canRedo()).toBe(false);
    surface(ana).undo();
    await settle();
    checked(ana, room, 7);
    expect(editorTexts(ana)[1]).toBe('Due in 30 days.');
  });

  test('a redone Enter and the text typed after it come back together', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    caretAt(ana, 0, 7);
    surface(ana).splitParagraph();
    await settle();
    surface(ana).insertPlainText('Q');
    await settle();
    expect(editorTexts(ana).slice(0, 2)).toEqual(['The fee', 'Q is 100 USD.']);
    surface(ana).undo();
    await settle();
    surface(ana).undo();
    await settle();
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD.');
    surface(ana).redo();
    await settle();
    surface(ana).redo();
    await settle();
    checked(ana, room, 6);
    expect(editorTexts(ana).slice(0, 2)).toEqual(['The fee', 'Q is 100 USD.']);
  });

  test('a paste the room refuses by name is refused before the editor commits', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const before = editorXml(ana);
    caretAt(ana, 0, 4);
    // A VML picture that only an address names: Word loads the address when it opens the
    // file, so no room writes one (design D18). The planned options of the op table are
    // refused at the gate in `op-coverage.test.ts`.
    const flavours = copiedBetween(0, 0, 'Plan \uFFFC'.length, VML_LINKED_SOURCE);
    expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(false);
    await settle();
    expect(surface(ana).state().lastRejection).toBe('ordered-refused');
    expect(ana.refusals).toEqual([
      { code: 'ordered-refused', reason: 'external-picture-link', detail: 'a picture by address' },
    ]);
    expect(editorXml(ana)).toBe(before);
    expect(room.room.order).toBe(0);
    expect(ana.handle.session.unsentChanges()).toEqual([]);
  });

  test('an edit the gate admitted but the editor did not commit is taken back, unsent', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const before = listTexts(ana);
    const paragraphId = surface(ana).session.paragraphIds()[0];
    if (paragraphId === undefined) throw new Error('no paragraph');
    const admitted = ana.handle.session.gateOperations(
      [
        {
          op: 'insertText',
          paragraphId,
          offset: 0,
          text: 'Never ',
          revision: { author: 'Ana', date: '2026-10-28T09:00:00Z' },
        },
      ],
      { kind: 'body' }
    );
    expect(admitted).toBeNull();
    expect(listTexts(ana)[0]).toBe('Never The fee is 100 USD.');
    // The editor never commits it: the flush takes it back before anything is sent
    ana.handle.session.flushPendingJournals();
    await settle();
    expect(listTexts(ana)).toEqual(before);
    expect(ana.handle.session.unsentChanges()).toEqual([]);
    expect(room.room.order).toBe(0);
    expect(screenFollowsList(ana)).toBe(true);
  });

  test('a write the gate cannot see is refused as an opaque write', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    let ran = false;
    surface(ana).commitReviewOps(() => {
      ran = true;
      return { committed: true };
    }, 'comment-add');
    expect(ran).toBe(false);
    expect(surface(ana).state().lastRejection).toBe('ordered-opaque-write');
    expect(ana.refusals).toEqual([{ code: 'ordered-opaque-write', reason: 'opaque-write' }]);
  });
});
