// The editor's cell commands from a caret in a vertically merged cell (task 9.13 of
// `ordered-suggestion-collaboration`). Word draws the merge as one cell, and the store's cell
// ops refuse a selection that names only some of its `w:tc` (`invalidArgs`). The command named
// the caret's cell alone, so a fill, a border or an alignment there was refused. It now covers
// the whole merged cell, and a cell outside the merge only itself.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { paragraphTextOf } from '../../store/store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, merge = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/>${merge}</w:tcPr>${text ? p(text) : '<w:p/>'}</w:tc>`;

/** The first column merges down over three rows; the second column does not. */
const TABLE =
  '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
  `<w:tr>${cell('Merged', '<w:vMerge w:val="restart"/>')}${cell('B1')}</w:tr>` +
  `<w:tr>${cell('', '<w:vMerge/>')}${cell('B2')}</w:tr>` +
  `<w:tr>${cell('', '<w:vMerge/>')}${cell('B3')}</w:tr>` +
  `<w:tr>${cell('A4')}${cell('B4')}</w:tr></w:tbl><w:p/>`;

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const mounted: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of mounted.splice(0)) editor.destroy();
});

/** An editor with the caret in the paragraph that reads `text`. */
function caretIn(text: string): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(TABLE) });
  mounted.push(editor);
  const surface = editor.surface;
  if (!surface) throw new Error('no surface');
  const id = surface.session
    .paragraphIds()
    .find((candidate) => paragraphTextOf(surface.session.part(), candidate) === text);
  if (!id) throw new Error(`no paragraph ${text}`);
  surface.setSelection({
    anchor: { paragraphId: id, offset: 1 },
    head: { paragraphId: id, offset: 1 },
  });
  return editor;
}

const count = (editor: DocxEditorInstance, what: string): number =>
  serializeOoxmlPart(editor.surface!.session.part()).split(what).length - 1;

describe('cell commands from a caret in a vertically merged cell', () => {
  test('a fill covers the whole merged cell', () => {
    const editor = caretIn('Merged');
    expect(editor.exec({ type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } }).ok).toBe(
      true
    );
    expect(count(editor, 'w:fill="FF0000"')).toBe(3);
  });

  test('borders and a vertical alignment cover it too', () => {
    const editor = caretIn('Merged');
    expect(
      editor.exec({
        type: 'setTableBorders',
        scope: 'all',
        spec: { style: 'single', size: 4, color: { kind: 'hex', value: '112233' } },
      }).ok
    ).toBe(true);
    // The store writes a merged cell's borders where Word reads them for the merged cell
    expect(count(editor, '<w:tcBorders>')).toBeGreaterThan(0);
    expect(editor.exec({ type: 'setTableCellVerticalAlignment', alignment: 'center' }).ok).toBe(
      true
    );
    expect(count(editor, '<w:vAlign w:val="center"/>')).toBe(3);
  });

  test('a cell outside the merge takes the press alone', () => {
    const editor = caretIn('B2');
    expect(editor.exec({ type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } }).ok).toBe(
      true
    );
    expect(count(editor, 'w:fill="FF0000"')).toBe(1);
  });
});
