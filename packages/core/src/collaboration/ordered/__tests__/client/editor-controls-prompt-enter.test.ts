// Enter in a block control's prompt, with real editors and one ordered room, against the editor
// alone. Word 16.114 (2026-10-30, `~$p4-enter.docx` -> `~$p4-enter-tracked.docx`, Track
// Changes on) selects the whole prompt, and Return replaces it: `w:showingPlcHdr` goes, the
// prompt's text goes with no `w:del`, and the control holds two empty paragraphs, the first
// mark inserted. Reject All (`~$p4-enter-rejected.docx`) leaves Word's blank prompt. The room
// (policy 4) and the editor alone give the same, wherever the caret is in the prompt, on every
// copy and in the save.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
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

/** Word's sample: a block control whose prompt has two paragraphs, between two lines. */
const FILE = docx(
  p(t('Before.')) +
    '<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/><w:sdtContent>' +
    p(t('First line.')) +
    p(t('Second line.')) +
    '</w:sdtContent></w:sdt>' +
    p(t('After.'))
);

/** The control as a reader of the saved file sees it: no paragraph ids, dates or authors. */
const controlOf = (xml: string): string =>
  (/<w:sdt>.*<\/w:sdt>/s.exec(xml)?.[0] ?? '<none>')
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:(date|author|id)="[^"]*"/g, '');

const TRACKED =
  '<w:sdt><w:sdtPr><w:id w:val="602"/></w:sdtPr><w:sdtEndPr/><w:sdtContent>' +
  '<w:p><w:pPr><w:rPr><w:ins/></w:rPr></w:pPr></w:p><w:p/></w:sdtContent></w:sdt>';
const REJECTED =
  '<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/><w:sdtContent>' +
  '<w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p></w:sdtContent></w:sdt>';

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
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

/** The editor alone in suggesting mode: Enter at the place, then its control. */
function alone(paragraph: number, offset: number): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Ana',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    caretAt(person, paragraph, offset);
    surface(person).splitParagraph();
    return controlOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The organiser decides the only revision card. */
function decide(person: Person, action: 'accept' | 'reject'): void {
  const [card] = person.editor.getReviewItems().filter((item) => item.kind === 'revision');
  if (card === undefined) throw new Error('no card');
  const result =
    action === 'accept'
      ? person.editor.acceptReviewItem(card.key)
      : person.editor.rejectReviewItem(card.key);
  expect(result.ok).toBe(true);
}

async function savedControl(person: Person): Promise<string> {
  const saved = new Uint8Array(await person.editor.save());
  expect(readOoxmlPackage(saved).ok).toBe(true);
  const again = openOrderedWordFile(saved, { file: 2 }).document;
  return controlOf(serializeOoxmlPart(composeOrderedPart(again, BODY).part));
}

async function seats() {
  const review = reviewTestModule();
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: FILE,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  return { room, olive, ana };
}

describe("Enter in a block control's prompt replaces the prompt, as Word does", () => {
  const places = [
    ['the start of its first line', 1, 0],
    ['inside its first line', 1, 3],
    ['the end of its first line', 1, 11],
    ['inside its second line', 2, 6],
  ] as const;
  for (const [where, paragraph, offset] of places)
    test(`a party's Enter at ${where}: the room equals the editor alone`, async () => {
      expect(alone(paragraph, offset)).toBe(TRACKED);
      const { room, olive, ana } = await seats();
      caretAt(ana, paragraph, offset);
      surface(ana).splitParagraph();
      await settle();
      converged(room, olive, ana);
      expect(controlOf(editorXml(ana))).toBe(TRACKED);
      expect(controlOf(editorXml(olive))).toBe(TRACKED);
      expect(await savedControl(ana)).toBe(TRACKED);
    });

  test("the organiser's reject leaves Word's blank prompt; an accept keeps two paragraphs", async () => {
    for (const action of ['reject', 'accept'] as const) {
      const { room, olive, ana } = await seats();
      caretAt(ana, 1, 3);
      surface(ana).splitParagraph();
      await settle();
      decide(olive, action);
      await settle();
      converged(room, olive, ana);
      const expected =
        action === 'reject'
          ? REJECTED
          : TRACKED.replace('<w:pPr><w:rPr><w:ins/></w:rPr></w:pPr>', '');
      expect(controlOf(editorXml(ana))).toBe(expected.replace('<w:p></w:p>', '<w:p/>'));
      expect(await savedControl(olive)).toBe(expected.replace('<w:p></w:p>', '<w:p/>'));
    }
  });

  test("a one-line prompt: the party's Enter, then undo; the organiser's Enter, then undo", async () => {
    const one = docx(
      p(t('Before.')) +
        '<w:sdt><w:sdtPr><w:id w:val="603"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
        p(t('Type here.')) +
        '</w:sdtContent></w:sdt>' +
        p(t('After.'))
    );
    const review = reviewTestModule();
    const room = localRoom(one);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: one,
      review,
    });
    const ana = await join(room, 'Ana', { wordFile: one, review });
    const before = controlOf(editorXml(ana));
    for (const person of [ana, olive]) {
      caretAt(person, 1, 4);
      surface(person).splitParagraph();
      await settle();
      converged(room, olive, ana);
      expect(controlOf(editorXml(ana))).not.toContain('Type here.');
      surface(person).undo();
      await settle();
      converged(room, olive, ana);
      expect(controlOf(editorXml(ana))).toBe(before);
    }
  });

  test("the party's undo brings the prompt back whole", async () => {
    const { room, olive, ana } = await seats();
    const before = controlOf(editorXml(ana));
    caretAt(ana, 1, 3);
    surface(ana).splitParagraph();
    await settle();
    surface(ana).undo();
    await settle();
    converged(room, olive, ana);
    expect(controlOf(editorXml(ana))).toBe(before);
    expect(controlOf(editorXml(olive))).toBe(before);
  });
});
