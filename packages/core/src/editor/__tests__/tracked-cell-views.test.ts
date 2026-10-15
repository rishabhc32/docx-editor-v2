// THE CARET AND TABLE COMMANDS IN A VIEW THAT HIDES A PROPOSED COLUMN.
//
// No Markup and Simple Markup lay out a column proposed for deletion as Accept would leave the
// table: the column is gone and the cell before it takes its width. A click there, an arrow key
// past it, and a caret drawn for it must all land in a cell the reader can see. The surviving
// cells are drawn in a grid the document does not have yet, so a table command must still
// address the stored column of the cell it starts from, not the column the view draws it in.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  collectReviewItems as engineCollectReviewItems,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { caretAt, hitTestPage } from '@docx-editor.dev/core/layout';
import type { EditorModule } from '../../contracts/modules.ts';
import { docx } from './paginated-surface-fixtures.ts';

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
const A1 = 1;
const A2 = 2;
const A3 = 3;

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
        return revisionItemsOf({ ...part, root: paragraph });
      },
    },
  };
}

const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

/** A suggesting editor whose author proposed deleting the column of paragraph `paragraph`. */
function proposeColumnDeletion(paragraph: number): DocxEditorInstance {
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(BODY),
    author: 'Grace Hopper',
    modules: [reviewModule()],
  });
  editors.push(editor);
  expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
  editor.surface!.setRevisionDisplayMode('all-markup');
  caret(editor, paragraph);
  expect(editor.exec({ type: 'deleteColumn' })).toEqual({ ok: true, changed: true });
  return editor;
}

const idOf = (editor: DocxEditorInstance, paragraph: number): string =>
  editor.surface!.session.paragraphIds()[paragraph]!;

function caret(editor: DocxEditorInstance, paragraph: number, offset = 0): void {
  const position = { paragraphId: idOf(editor, paragraph), offset };
  editor.surface!.setSelection({ anchor: position, head: position });
}

test.each(['proposed', 'simple-markup'] as const)(
  'in %s a click, an arrow key and the caret skip the hidden column',
  (view) => {
    const editor = proposeColumnDeletion(A2);
    const shown = caretAt(editor.surface!.layout(), { paragraphId: idOf(editor, A2), offset: 1 })!;
    expect(shown).not.toBeNull();

    editor.surface!.setRevisionDisplayMode(view);
    const layout = editor.surface!.layout();
    // The column is gone, and the cell before it took its space.
    expect(caretAt(layout, { paragraphId: idOf(editor, A2), offset: 1 })).toBeNull();
    const hit = hitTestPage(layout, shown.pageIndex, { x: shown.x, y: shown.y + 1 });
    expect(hit?.position.paragraphId).toBe(idOf(editor, A1));
    // Right from the end of a1 goes on to a3, past the hidden a2.
    caret(editor, A1, 2);
    editor.surface!.navigate('right');
    expect(editor.surface!.state().selection.head).toEqual({
      paragraphId: idOf(editor, A3),
      offset: 0,
    });
    expect(caretAt(layout, editor.surface!.state().selection.head)).not.toBeNull();
  }
);

test('Insert Column beside a cell that took a hidden column inserts beside that cell', () => {
  // The first column goes, so a2 is drawn in the first column of the resolved grid.
  const editor = proposeColumnDeletion(A1);
  editor.surface!.setRevisionDisplayMode('proposed');
  caret(editor, A2);
  expect(editor.exec({ type: 'insertColumn', where: 'right' })).toEqual({
    ok: true,
    changed: true,
  });
  const firstRow = serializeOoxmlPart(editor.surface!.session.part()).match(
    /<w:tr>.*?<\/w:tr>/
  )![0];
  const cells = [...firstRow.matchAll(/<w:tc>.*?<\/w:tc>/g)].map(([xml]) =>
    xml.includes('<w:cellDel ')
      ? 'deleted'
      : xml.includes('<w:cellIns ')
        ? 'inserted'
        : xml.match(/<w:t(?: [^>]*)?>([^<]*)</)![1]
  );
  expect(cells).toEqual(['deleted', 'a2', 'inserted', 'a3']);
});
