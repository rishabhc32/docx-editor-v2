// SUGGESTING MODE PROPOSES COLUMNS, AND TABLES A RANGE DELETION COVERS.
//
// Delete Columns and Insert Column changed the table outright in suggesting mode, and a
// selection from the text before a table to the text after it removed the table with nothing
// to reject. Each now leaves a card that Accept and Reject answer, one undo step, and markup a
// save keeps. Rejecting a table its author proposed and typed into was refused until the typing
// was answered on its own; it now takes the typing with it.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, setSystemTime, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  collectReviewItems as engineCollectReviewItems,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import type { EditorCommand } from '../../contracts/editor.ts';
import { paragraphTextFromLayout } from '@docx-editor.dev/core/layout';
import { HELD_CHANGE_REASON } from '../review-held-changes.ts';
import { docx, selectCellRectangle } from './paginated-surface-fixtures.ts';

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
const row = (prefix: string) =>
  `<w:tr>${cell(`${prefix}1`, 2000)}${cell(`${prefix}2`, 3000)}${cell(`${prefix}3`, 4000)}</w:tr>`;
// Paragraphs in order: before, a1, a2, a3, b1, b2, b3, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `${row('a')}${row('b')}</w:tbl>` +
  p('after');
const A2 = 2;
const AFTER = 7;

/** The engine's own review derivation, wired as a module (core may not import pro). */
function reviewModule(): EditorModule {
  return {
    id: 'review',
    review: {
      displayModes: ['all-markup', 'proposed', 'original'],
      collectReviewItems: engineCollectReviewItems,
      revisionItemsOfParagraph: (part, paragraphId) => {
        const paragraph = findNode(part, paragraphId);
        if (!paragraph || paragraph.kind !== 'paragraph') return [];
        return revisionItemsOf({
          id: part.id,
          name: part.name,
          contentType: part.contentType,
          root: paragraph,
        });
      },
    },
  };
}

const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

function mount(document: Uint8Array, mode: 'editing' | 'suggesting'): DocxEditorInstance {
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document,
    author: 'Grace Hopper',
    modules: [reviewModule()],
  });
  if (!editor.surface) throw new Error('surface failed to mount');
  editors.push(editor);
  expect(editor.exec({ type: 'setEditingMode', mode }).ok).toBe(true);
  return editor;
}

function caret(editor: DocxEditorInstance, paragraph: number, offset = 0): void {
  const position = { paragraphId: editor.surface!.session.paragraphIds()[paragraph]!, offset };
  editor.surface!.setSelection({ anchor: position, head: position });
}

function select(editor: DocxEditorInstance, from: [number, number], to: [number, number]): void {
  const ids = editor.surface!.session.paragraphIds();
  editor.surface!.setSelection({
    anchor: { paragraphId: ids[from[0]]!, offset: from[1] },
    head: { paragraphId: ids[to[0]]!, offset: to[1] },
  });
}

const xmlOf = (editor: DocxEditorInstance): string =>
  serializeOoxmlPart(editor.surface!.session.part());
const cards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind));
const texts = (editor: DocxEditorInstance): string[] =>
  editor
    .surface!.session.paragraphIds()
    .map((id) => paragraphTextFromLayout(editor.surface!.layout(), id));
const gridWidths = (xml: string): number[] =>
  [...xml.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((match) => Number(match[1]));
const resolveAll = (editor: DocxEditorInstance, action: 'accept' | 'reject') =>
  expect(editor.exec({ type: 'resolveAllReviewChanges', action }).ok).toBe(true);

interface ColumnCase {
  readonly name: string;
  readonly command: EditorCommand;
  /** Select columns two and three (grid 1..2) instead of placing the caret in a2. */
  readonly rectangle?: boolean;
  readonly marker: 'cellDel' | 'cellIns';
  /** The markers the proposal writes. */
  readonly markers: number;
  /** Grid widths with the proposal accepted, and in editing mode. */
  readonly accepted: readonly number[];
  readonly editing: readonly number[];
}

const COLUMN_CASES: readonly ColumnCase[] = [
  {
    name: 'Delete Columns',
    command: { type: 'deleteColumn' },
    marker: 'cellDel',
    markers: 2,
    accepted: [5000, 4000],
    editing: [2000, 4000],
  },
  {
    name: 'Delete Columns over two columns',
    command: { type: 'deleteColumn' },
    rectangle: true,
    marker: 'cellDel',
    markers: 4,
    accepted: [9000],
    editing: [2000],
  },
  {
    name: 'Insert Column to the right',
    command: { type: 'insertColumn', where: 'right' },
    marker: 'cellIns',
    markers: 2,
    accepted: [2000, 1500, 1500, 4000],
    editing: [2000, 3000, 3000, 4000],
  },
  {
    name: 'Insert Column to the left',
    command: { type: 'insertColumn', where: 'left' },
    marker: 'cellIns',
    markers: 2,
    accepted: [1000, 1000, 3000, 4000],
    editing: [2000, 3000, 3000, 4000],
  },
];

function run(editor: DocxEditorInstance, entry: ColumnCase): void {
  if (entry.rectangle) {
    selectCellRectangle(editor.surface!, { row: 0, column: 1 }, { row: 1, column: 2 });
  } else caret(editor, A2);
  expect(editor.exec(entry.command)).toEqual({ ok: true, changed: true });
}

describe.each(COLUMN_CASES)('$name', (entry) => {
  test('suggesting leaves one card over the marked column, and Accept applies it', () => {
    const editor = mount(docx(BODY), 'suggesting');
    run(editor, entry);
    expect(cards(editor)).toEqual(['structural']);
    const markers = xmlOf(editor).match(new RegExp(`<w:${entry.marker} [^>]*/>`, 'g')) ?? [];
    expect(markers).toHaveLength(entry.markers);
    expect(new Set(markers).size).toBe(1);
    expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(gridWidths(xmlOf(editor))).toEqual([...entry.accepted]);
    expect(xmlOf(editor)).not.toMatch(/<w:(cellDel|cellIns|gridSpan)\b/);
  });

  test('Reject restores the document exactly', () => {
    const editor = mount(docx(BODY), 'suggesting');
    const original = xmlOf(editor);
    run(editor, entry);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(xmlOf(editor)).toBe(original);
  });

  test('one undo step takes the proposal back', () => {
    const editor = mount(docx(BODY), 'suggesting');
    const original = xmlOf(editor);
    run(editor, entry);
    expect(cards(editor)).toEqual(['structural']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(original);
    expect(cards(editor)).toEqual([]);
  });

  test('a save keeps the proposal, and reopening reads the same card', async () => {
    const editor = mount(docx(BODY), 'suggesting');
    run(editor, entry);
    const reopened = mount(new Uint8Array(await editor.save()), 'editing');
    expect(cards(reopened)).toEqual(['structural']);
    expect(xmlOf(reopened)).toBe(xmlOf(editor));
  });

  test('editing mode applies it outright, where suggesting proposes it', () => {
    const editing = mount(docx(BODY), 'editing');
    run(editing, entry);
    expect(cards(editing)).toEqual([]);
    expect(gridWidths(xmlOf(editing))).toEqual([...entry.editing]);
    expect(xmlOf(editing)).not.toMatch(/<w:(cellDel|cellIns)\b/);
    const suggesting = mount(docx(BODY), 'suggesting');
    run(suggesting, entry);
    expect(cards(suggesting)).toEqual(['structural']);
  });
});

test('deleting a proposed column withdraws it', () => {
  const editor = mount(docx(BODY), 'suggesting');
  const original = xmlOf(editor);
  caret(editor, A2);
  expect(editor.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
  expect(cards(editor)).toEqual(['structural']);
  // The caret lands in the new column, a2's neighbour.
  expect(editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
  expect(cards(editor)).toEqual([]);
  expect(xmlOf(editor)).toBe(original);
});

test('two Delete Column presses in one second are two cards', () => {
  // Both presses take the same revision date, as two presses within a wall-clock second do.
  setSystemTime(new Date('2026-10-01T12:00:00Z'));
  try {
    const editor = mount(docx(BODY), 'suggesting');
    caret(editor, A2);
    expect(editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    caret(editor, A2 - 1);
    expect(editor.exec({ type: 'deleteColumn' }).ok).toBe(true);
    expect(cards(editor)).toEqual(['structural', 'structural']);
  } finally {
    setSystemTime();
  }
});

describe('a range deletion over a whole table proposes the table', () => {
  test('the text is struck, the table is marked, and no paragraph joins across it', () => {
    const editor = mount(docx(BODY), 'suggesting');
    select(editor, [0, 2], [AFTER, 2]);
    editor.surface!.deleteForward();
    expect(cards(editor)).toEqual(['delete', 'structural', 'delete']);
    const xml = xmlOf(editor);
    expect(xml.match(/<w:trPr><w:del [^>]*\/><\/w:trPr>/g)).toHaveLength(2);
    expect(xml).toContain('<w:delText>fore</w:delText>');
    expect(xml).toContain('<w:delText>af</w:delText>');
    expect(xml).not.toMatch(/<w:pPr><w:rPr><w:del\b/);
    // Editing mode still removes the table and joins the paragraphs.
    const editing = mount(docx(BODY), 'editing');
    select(editing, [0, 2], [AFTER, 2]);
    editing.surface!.deleteForward();
    expect(xmlOf(editing)).not.toContain('<w:tbl>');
    expect(texts(editing)).toEqual(['beter']);
  });

  test('Reject All restores it; Accept All removes the table and keeps both paragraphs', () => {
    const editor = mount(docx(BODY), 'suggesting');
    const original = texts(editor);
    select(editor, [0, 2], [AFTER, 2]);
    editor.surface!.deleteBackward();
    resolveAll(editor, 'reject');
    expect(texts(editor)).toEqual(original);
    expect(xmlOf(editor)).not.toMatch(/<w:(del|cellDel)\b/);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    resolveAll(editor, 'accept');
    expect(xmlOf(editor)).not.toContain('<w:tbl>');
    // The untracked deletion joins them; a proposal never proposed the break across the table.
    expect(texts(editor)).toEqual(['be', 'ter']);
  });

  test('typing over the range lands after the struck text, and the table is proposed', () => {
    const editor = mount(docx(BODY), 'suggesting');
    select(editor, [0, 2], [AFTER, 2]);
    editor.surface!.type('X');
    expect(cards(editor)).toEqual(['delete', 'structural', 'replace']);
    expect(xmlOf(editor)).toMatch(
      /<w:delText>af<\/w:delText><\/w:r><\/w:del><w:ins [^>]*><w:r><w:t>X</
    );
  });

  test('a save keeps the proposal', async () => {
    const editor = mount(docx(BODY), 'suggesting');
    select(editor, [0, 2], [AFTER, 2]);
    editor.surface!.deleteForward();
    const reopened = mount(new Uint8Array(await editor.save()), 'editing');
    expect(cards(reopened)).toEqual(['delete', 'structural', 'delete']);
    expect(xmlOf(reopened)).toBe(xmlOf(editor));
  });
});

describe('rejecting a proposed table its author typed into', () => {
  function proposeTableAndType(typist: string): DocxEditorInstance {
    const editor = mount(docx(p('before')), 'suggesting');
    caret(editor, 0, 6);
    expect(editor.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true);
    editor.setAuthor(typist);
    editor.surface!.type('hi');
    expect(cards(editor)).toEqual(['structural', 'insert']);
    return editor;
  }

  test('takes the typing with it, in one undo step', () => {
    const editor = mount(docx(p('before')), 'suggesting');
    const original = xmlOf(editor);
    editors.pop()!.destroy();
    const proposed = proposeTableAndType('Grace Hopper');
    const table = proposed
      .getReviewItems()
      .find((item) => item.kind === 'revision' && item.revisionKind === 'structural')!;
    expect(proposed.rejectReviewItem(table.key)).toEqual({ ok: true, changed: true });
    expect(cards(proposed)).toEqual([]);
    expect(xmlOf(proposed)).toBe(original);
    expect(proposed.exec({ type: 'undo' }).ok).toBe(true);
    expect(cards(proposed)).toEqual(['structural', 'insert']);
  });

  test("is refused while another author's typing is pending, and says why", () => {
    const editor = proposeTableAndType('Ada Lovelace');
    const table = editor
      .getReviewItems()
      .find((item) => item.kind === 'revision' && item.revisionKind === 'structural')!;
    const refused = editor.rejectReviewItem(table.key);
    expect(refused).toEqual({ ok: false, code: 'unsupported', reason: HELD_CHANGE_REASON });
    // Answering the typing first clears the way.
    const typing = editor
      .getReviewItems()
      .find((item) => item.kind === 'revision' && item.revisionKind === 'insert')!;
    expect(editor.rejectReviewItem(typing.key).ok).toBe(true);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(xmlOf(editor)).not.toContain('<w:tbl>');
  });
});
