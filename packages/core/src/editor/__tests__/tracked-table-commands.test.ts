// SUGGESTING MODE PROPOSES TABLE DELETION AND CELL FORMATTING.
//
// Delete Table, Delete Rows or Delete Columns over the whole table, cell fill, vertical
// alignment and cell borders changed the document outright in suggesting mode: no review card,
// nothing to reject. Each now leaves one card that Accept and Reject answer, one undo step, and
// markup a save keeps.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  collectReviewItems as engineCollectReviewItems,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import type { EditorCommand } from '../../contracts/editor.ts';
import { resolveRevisionMarkup } from '../../contracts/revision-markup.ts';
import { selectCellRectangle } from './paginated-surface-fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}</w:tr><w:tr>${cell('b1')}${cell('b2')}</w:tr></w:tbl>` +
  p('after');

function docxOf(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

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
  // The caret in the first cell, where every command below acts.
  const paragraphId = editor.surface.session.paragraphIds()[1]!;
  const caret = { paragraphId, offset: 0 };
  editor.surface.setSelection({ anchor: caret, head: caret });
  return editor;
}

const xmlOf = (editor: DocxEditorInstance): string =>
  serializeOoxmlPart(editor.surface!.session.part());
/** Rejecting a row deletion leaves its emptied `w:trPr` behind; nothing else differs. */
const withoutEmptyRowProperties = (xml: string): string => xml.replaceAll('<w:trPr/>', '');
const cards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind));

interface Case {
  readonly name: string;
  readonly command: EditorCommand;
  /** Whether `command` acts on a rectangle over every cell rather than at the caret. */
  readonly everyCell?: boolean;
  readonly card: 'structural' | 'format';
  /** Whether the document shows the command's effect, accepted or applied outright. */
  readonly applied: (xml: string) => boolean;
}

const tableGone = (xml: string) => !xml.includes('<w:tbl>');
const CASES: readonly Case[] = [
  {
    name: 'Delete Table',
    command: { type: 'deleteTable' },
    card: 'structural',
    applied: tableGone,
  },
  {
    name: 'Delete Rows over every row',
    everyCell: true,
    command: { type: 'deleteRow' },
    card: 'structural',
    applied: tableGone,
  },
  {
    name: 'Delete Columns over every column',
    everyCell: true,
    command: { type: 'deleteColumn' },
    card: 'structural',
    applied: tableGone,
  },
  {
    name: 'cell fill',
    command: { type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } },
    card: 'format',
    applied: (xml) => xml.includes('<w:shd w:fill="FF0000" w:val="clear"/>'),
  },
  {
    name: 'cell vertical alignment',
    command: { type: 'setTableCellVerticalAlignment', alignment: 'center' },
    card: 'format',
    applied: (xml) => xml.includes('<w:vAlign w:val="center"/>'),
  },
  {
    name: 'cell borders',
    command: {
      type: 'setTableBorders',
      scope: 'all',
      spec: { style: 'single', size: 8, color: { kind: 'hex', value: '000000' } },
    },
    card: 'format',
    applied: (xml) => xml.includes('<w:top w:color="000000" w:sz="8" w:val="single"/>'),
  },
];

function run(editor: DocxEditorInstance, entry: Case): void {
  if (entry.everyCell)
    selectCellRectangle(editor.surface!, { row: 0, column: 0 }, { row: 1, column: 1 });
  expect(editor.exec(entry.command)).toEqual({ ok: true, changed: true });
}

describe.each(CASES)('$name', (entry) => {
  test('suggesting leaves one card, and Accept applies it', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    run(editor, entry);
    expect(cards(editor)).toEqual([entry.card]);
    expect(entry.applied(xmlOf(editor))).toBe(entry.card === 'format');
    expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(entry.applied(xmlOf(editor))).toBe(true);
    expect(xmlOf(editor)).not.toMatch(/<w:(del|cellDel|tcPrChange)\b/);
  });

  test('Reject restores the document', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    const original = xmlOf(editor);
    run(editor, entry);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(withoutEmptyRowProperties(xmlOf(editor))).toBe(original);
  });

  test('one undo step takes the proposal back', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    const original = xmlOf(editor);
    run(editor, entry);
    expect(cards(editor)).toEqual([entry.card]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(original);
    expect(cards(editor)).toEqual([]);
  });

  test('a save keeps the proposal, and reopening reads the same card', async () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    run(editor, entry);
    const reopened = mount(new Uint8Array(await editor.save()), 'editing');
    expect(cards(reopened)).toEqual([entry.card]);
    expect(xmlOf(reopened)).toBe(xmlOf(editor));
  });

  test('editing mode applies it outright, where suggesting proposes it', () => {
    const editing = mount(docxOf(BODY), 'editing');
    run(editing, entry);
    expect(cards(editing)).toEqual([]);
    expect(entry.applied(xmlOf(editing))).toBe(true);
    expect(xmlOf(editing)).not.toMatch(/<w:(del|cellDel|tcPrChange)\b/);
    const suggesting = mount(docxOf(BODY), 'suggesting');
    run(suggesting, entry);
    expect(cards(suggesting)).toEqual([entry.card]);
  });
});

test('with formatting changes untracked, cell formatting applies bare and deletion is still proposed', () => {
  const [deleteTable, fill] = ['Delete Table', 'cell fill'].map(
    (name) => CASES.find((entry) => entry.name === name)!
  );
  const editor = mount(docxOf(BODY), 'suggesting');
  editor.surface!.setRevisionMarkup(resolveRevisionMarkup({ trackFormatting: false }));
  run(editor, fill!);
  expect(cards(editor)).toEqual([]);
  expect(fill!.applied(xmlOf(editor))).toBe(true);
  run(editor, deleteTable!);
  expect(cards(editor)).toEqual(['structural']);
});

test('Delete Table withdraws a row the author proposed, and proposes every other row', () => {
  const editor = mount(docxOf(BODY), 'suggesting');
  const original = xmlOf(editor);
  expect(editor.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  const proposed = xmlOf(editor);
  expect(editor.exec({ type: 'deleteTable' })).toEqual({ ok: true, changed: true });
  expect(cards(editor)).toEqual(['structural']);
  expect(xmlOf(editor).match(/<w:tr>/g)).toHaveLength(2);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(xmlOf(editor)).toBe(proposed);
  expect(editor.exec({ type: 'redo' }).ok).toBe(true);
  expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
  expect(withoutEmptyRowProperties(xmlOf(editor))).toBe(original);
});
