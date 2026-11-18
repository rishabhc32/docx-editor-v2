// A party's redo of an Enter in a link or a control, after the organiser deleted the words on
// one side of the caret and the party took the Enter back, with real editors and
// one ordered room. The redo makes the break again where its place shows now; each copy and the
// saved file read as the editor alone that opens the document as it was before the redo and
// presses Enter there. Undo of the redo gives that document back.
//
// The redo once put the break inside the control or link, past the close item of the taken-back
// split (the close item it read as the container's own end), with no split. The next paragraph
// then held the container's own close item with no open item, and compose threw `An inline
// drawing without a box`: the party's redo did nothing, its later edits did not go to the room,
// and at the next change from the room its session stopped (`remote-apply-failed`).

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
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

const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const link = (content: string) => `<w:hyperlink r:id="rIdLink">${content}</w:hyperlink>`;
const LINK_FACE = '<w:rStyle w:val="Hyperlink"/>';

/** A paragraph and an offset in it, and the end of a selection there. */
type Place = readonly [line: number, at: number, to?: number];

interface Case {
  readonly file: Uint8Array;
  /** The party's Enter: between "Acme" and " Corp". */
  readonly enter: Place;
  /** The organiser's deletion after the caret (" Corp") and before it ("Acme"). */
  readonly after: Place;
  readonly before: Place;
  /** Where the redo's place shows after each deletion and the undo. */
  readonly redoAfter: Place;
  readonly redoBefore: Place;
  /** The paragraph texts after each redo. */
  readonly textsAfter: readonly string[];
  readonly textsBefore: readonly string[];
}

const CASES: Readonly<Record<string, Case>> = {
  'an inline control': {
    file: docx(
      p(t('Name: ') + sdt('<w:text w:multiLine="1"/>', t('Acme Corp')) + t(' end.')) +
        p(t('After.'))
    ),
    enter: [0, 10],
    after: [1, 0, 5],
    before: [0, 6, 10],
    redoAfter: [0, 10],
    redoBefore: [0, 6],
    textsAfter: ['Name: Acme', ' end.', 'After.'],
    textsBefore: ['Name: ', ' Corp end.', 'After.'],
  },
  'a block control': {
    file: docx(p(t('Before.')) + sdt('<w:id w:val="602"/>', p(t('Acme Corp'))) + p(t('After.'))),
    enter: [1, 4],
    after: [2, 0, 5],
    before: [1, 0, 4],
    redoAfter: [1, 4],
    redoBefore: [1, 0],
    textsAfter: ['Before.', 'Acme', '', 'After.'],
    textsBefore: ['Before.', '', ' Corp', 'After.'],
  },
  'a link': {
    file: docx(p(t('Name: ') + link(t('Acme Corp', LINK_FACE)) + t(' end.')) + p(t('After.'))),
    enter: [0, 10],
    after: [1, 0, 5],
    before: [0, 6, 10],
    redoAfter: [0, 10],
    redoBefore: [0, 6],
    textsAfter: ['Name: Acme', ' end.', 'After.'],
    textsBefore: ['Name: ', ' Corp end.', 'After.'],
  },
};

/** The body as a reader compares it: no ids, dates or paragraph ids. */
const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:(id|date)="[^"]*"/g, '');

/** The body after the party's Enter (or typing) at `place` on the editor alone that opens `file`. */
function aloneEnter(file: Uint8Array, [line, at]: Place, typed?: string): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Ana',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    caretAt(person, line, at);
    if (typed === undefined) surface(person).splitParagraph();
    else surface(person).insertPlainText(typed);
    return bodyOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

async function savedBody(person: Person): Promise<string> {
  const saved = new Uint8Array(await person.editor.save());
  const again = openOrderedWordFile(saved, { file: 2 }).document;
  return bodyOf(serializeOoxmlPart(composeOrderedPart(again, BODY).part));
}

/** Every copy is the room's, the editors show their lists, and both show `body`. */
function converged(people: readonly Person[], room: ReturnType<typeof localRoom>, body: string) {
  for (const person of people) {
    const state = {
      person: person.name,
      refusals: person.refusals,
      room: sameAsRoom(person, room),
      screen: screenFollowsList(person),
      body: bodyOf(editorXml(person)),
    };
    expect(state).toEqual({ ...state, refusals: [], room: true, screen: true, body });
  }
}

describe('redo of an Enter taken back after the organiser deleted words beside it', () => {
  for (const [name, c] of Object.entries(CASES))
    for (const side of ['after', 'before'] as const)
      test(`${name}, the words ${side} the caret deleted`, async () => {
        const room = localRoom(c.file);
        const review = reviewTestModule();
        const olive = await join(room, 'Olive', {
          mode: 'edit',
          login: ORGANISER,
          wordFile: c.file,
          review,
        });
        const ana = await join(room, 'Ana', { wordFile: c.file, review });
        const people = [olive, ana];
        caretAt(ana, c.enter[0], c.enter[1]);
        surface(ana).splitParagraph();
        await settle();
        const [line, at, to] = side === 'after' ? c.after : c.before;
        caretAt(olive, line, at, to);
        surface(olive).deleteBackward();
        await settle();
        surface(ana).undo();
        await settle();
        const undone = bodyOf(editorXml(olive));
        converged(people, room, undone);
        const expected = aloneEnter(
          new Uint8Array(await olive.editor.save()),
          side === 'after' ? c.redoAfter : c.redoBefore
        );

        surface(ana).redo();
        await settle();
        converged(people, room, expected);
        expect(editorTexts(ana)).toEqual([...(side === 'after' ? c.textsAfter : c.textsBefore)]);
        expect(await savedBody(ana)).toBe(expected);

        // Undo of the redo takes the break back again
        surface(ana).undo();
        await settle();
        converged(people, room, undone);
      });
});

describe('typing where an Enter taken back split a control or a link', () => {
  // The same place as the redo's: the words after the caret went, so the caret is at the end of
  // the container, and the typed text goes right after it (placement rule 5)
  for (const name of ['an inline control', 'a link'])
    test(`${name}: the party types at the end of the first part's words`, async () => {
      const c = CASES[name];
      if (!c) throw new Error(`no case ${name}`);
      const room = localRoom(c.file);
      const review = reviewTestModule();
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: c.file,
        review,
      });
      const ana = await join(room, 'Ana', { wordFile: c.file, review });
      caretAt(ana, c.enter[0], c.enter[1]);
      surface(ana).splitParagraph();
      await settle();
      caretAt(olive, c.after[0], c.after[1], c.after[2]);
      surface(olive).deleteBackward();
      await settle();
      surface(ana).undo();
      await settle();
      const expected = aloneEnter(new Uint8Array(await olive.editor.save()), c.redoAfter, 'X');
      caretAt(ana, c.redoAfter[0], c.redoAfter[1]);
      surface(ana).insertPlainText('X');
      await settle();
      converged([olive, ana], room, expected);
      expect(editorTexts(ana)).toEqual(['Name: AcmeX end.', 'After.']);
      expect(await savedBody(ana)).toBe(expected);
    });
});
