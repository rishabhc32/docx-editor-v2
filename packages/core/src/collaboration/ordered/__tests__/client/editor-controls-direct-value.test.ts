// The organiser's direct value in a control that a pending suggestion put in, with real editors
// and one ordered room, against the editor alone: in their own suggested wrap (undone and redone
// too) and in another person's pending control. Each is accepted and rejected live, after a save
// and reopen, and after one and two folds.
//
// A direct edit inside a pending insertion is typed into it as its author's (design D9), as Word
// extends a tracked insertion with untracked typing: the value is part of the control's insertion,
// so a reject takes the control with its value, and an accept keeps both. The room once kept the
// value as settled text in the control and the wrapped words hidden in it: a reject live left the
// value as plain text, where the file, the fold and the editor alone left the words as they were.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { checkpointBytes, folded, mediaStore } from '../fold-check.ts';
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
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const LINE = docx(p(t('Line: Revenue')) + p(t('Signed.')));

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/>' +
  '<w:listItem w:displayText="Net 60" w:value="60"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';
const CHECK_BOX =
  '<w14:checkbox><w14:checked w14:val="0"/>' +
  '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>';

/** "Term: <control> end." with the control inserted by Ana's pending suggestion. */
const anasControl = (props: string, content: string) =>
  docx(
    p(
      t('Term: ') +
        '<w:ins w:id="1" w:author="Ana" w:date="2026-10-28T07:00:00Z">' +
        `<w:sdt><w:sdtPr><w:id w:val="5"/>${props}</w:sdtPr><w:sdtContent>${content}` +
        '</w:sdtContent></w:sdt></w:ins>' +
        t(' end.')
    ) + p(t('Signed.'))
  );

/**
 * Each wrap: its label, the control type, and the value set in it. A date picker the wrap makes
 * has no `w:dateFormat`, where the editor's string date and the room's typed date read apart:
 * `checkpoint-fold-decide.test.ts` folds that one.
 */
const WRAPS = [
  ['a combo box', 'comboBox', 'Net 30'],
  ['a plain text control', 'plainText', 'Acme'],
  ['a rich text control', 'richText', 'Acme'],
] as const;

/** Each control of Ana's: its label, the Word file, and the value set in it. */
const ANAS = [
  ['a date picker', anasControl(`<w:date>${DATE_FORMAT}</w:date>`, t('Soon')), '2026-09-15'],
  ['a drop-down list', anasControl(`<w:dropDownList>${ITEMS}</w:dropDownList>`, t('Net 30')), '60'],
  ['a combo box', anasControl(`<w:comboBox>${ITEMS}</w:comboBox>`, t('Net 30')), 'Net 90'],
  ['a check box', anasControl(CHECK_BOX, t('☐')), 'true'],
  ['a plain text control', anasControl('<w:text/>', t('Acme')), 'Beta'],
  ['a rich text control', anasControl('', t('Acme')), 'Beta'],
] as const;

type Decision = 'accept' | 'reject';
type Act = (person: Person) => void;

/** Wrap "Revenu" in a new control of `type`, as a suggestion. */
const wrap =
  (type: string): Act =>
  (person) => {
    person.editor.setMode('suggesting');
    caretAt(person, 0, 6, 12);
    const command = { type: 'insertContentControl', subtype: type, tag: 'clause.86' };
    expect(person.editor.exec(command as never).ok).toBe(true);
  };

/** Set the first control's value directly. */
const value =
  (text: string): Act =>
  (person) => {
    person.editor.setMode('edit');
    const id = surface(person).layout().contentControls?.[0]?.id ?? '';
    expect(surface(person).contentControls.setValue(id, text)).toBe(true);
  };

const undo: Act = (person) => surface(person).undo();
const redo: Act = (person) => surface(person).redo();

/** The first paragraph as a reader of the file sees it: no ids, dates or revision sessions. */
const lineOf = (xml: string): string =>
  (/<w:p\b.*?<\/w:p>/s.exec(/<w:body>(.*)<\/w:body>/s.exec(xml)?.[1] ?? '')?.[0] ?? 'no line')
    .replace(/ w:(id|date|rsidR|rsidRPr|rsidDel)="[^"]*"| w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '')
    .replace(/<w:placeholder>.*?<\/w:placeholder>/g, '')
    .replace(/ xml:space="preserve"|<w:rPr>.*?<\/w:rPr>/g, '')
    .replace(/<\/w:r><w:r>/g, '')
    .replace(/<\/w:t><w:t>/g, '');

/** Accept or reject every revision, as the organiser. */
function decideAll(person: Person, decision: Decision): void {
  const command = { type: 'resolveAllReviewChanges', action: decision, scope: 'document' };
  expect(person.editor.exec(command as never).ok).toBe(true);
}

/** What the editor alone writes: after the steps, and after the decision. */
function alone(
  file: Uint8Array,
  steps: readonly Act[],
  decision: Decision
): { readonly edited: string; readonly decided: string } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olive',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    for (const act of steps) {
      act(person);
    }

    surface(person).flushPendingInput();
    const edited = lineOf(editorXml(person));
    decideAll(person, decision);
    return { edited, decided: lineOf(editorXml(person)) };
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The organiser, Olive, and Ben, a party, in a room on `file` (a Word file or a checkpoint). */
async function seats(file: Uint8Array, number: number) {
  const review = reviewTestModule();
  const room = localRoom(file, { file: number });
  const olive = await join(room, 'Olive', { login: ORGANISER, wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  return { room, olive, ben };
}

/** The room's line, the same on both copies. */
function roomLine(seated: Awaited<ReturnType<typeof seats>>): string {
  for (const person of [seated.olive, seated.ben]) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, seated.room) }).toEqual({
      person: person.name,
      room: true,
    });
  }

  return lineOf(editorXml(seated.olive));
}

/** The checkpoint of the room's document folded `times` times. */
async function foldedBytes(
  doc: OrderedDocument,
  file: Uint8Array,
  times: number
): Promise<Uint8Array> {
  const { base } = openOrderedWordFile(file, { file: 1 });
  const blobs = mediaStore(base);
  let copy = { document: doc, base };
  for (let fold = 1; fold < times; fold += 1) {
    copy = await folded(copy.document, copy.base, fold + 1, blobs);
  }

  return checkpointBytes(copy.document, copy.base, blobs);
}

/**
 * The organiser's steps in a room on `file`, decided live, on the save reopened and on the
 * room folded once and twice, each as the editor alone.
 */
async function check(
  label: string,
  file: Uint8Array,
  steps: readonly Act[],
  decision: Decision
): Promise<void> {
  const expected = alone(file, steps, decision);
  const seated = await seats(file, 1);
  for (const act of steps) {
    act(seated.olive);
    await settle();
  }

  expect({ label, edited: roomLine(seated) }).toEqual({ label, edited: expected.edited });

  // A save reopened, and the room folded once and twice, each decided by the organiser
  const saved = new Uint8Array(await seated.olive.editor.save());
  const copies = [
    ['saved', saved],
    ['folded once', await foldedBytes(seated.room.room.document, file, 1)],
    ['folded twice', await foldedBytes(seated.room.room.document, file, 2)],
  ] as const;
  decideAll(seated.olive, decision);
  await settle();
  expect({ label, live: roomLine(seated) }).toEqual({ label, live: expected.decided });
  for (const [copy, bytes] of copies) {
    const reopened = await seats(bytes, 2);
    expect({ label, copy, edited: roomLine(reopened) }).toEqual({
      label,
      copy,
      edited: expected.edited,
    });
    decideAll(reopened.olive, decision);
    await settle();
    const decided = serializeOoxmlPart(composeOrderedPart(reopened.room.room.document, BODY).part);
    expect({ label, copy, decided: lineOf(decided) }).toEqual({
      label,
      copy,
      decided: expected.decided,
    });
  }
}

describe("the organiser's direct value in their own suggested wrap, as the editor alone", () => {
  for (const [label, type, text] of WRAPS) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, LINE, [wrap(type), value(text)], decision);
      });
    }
  }
});

describe("the organiser's direct value in their own suggested wrap, undone, as alone", () => {
  const steps = [wrap('plainText'), value('Acme')];
  for (const [label, more] of [
    ['undone', [undo]],
    ['undone and redone', [undo, redo]],
  ] as const) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, LINE, [...steps, ...more], decision);
      });
    }
  }
});

describe("the organiser's direct value in another person's pending control, as alone", () => {
  for (const [label, file, text] of ANAS) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, file, [value(text)], decision);
      });
    }
  }
});
