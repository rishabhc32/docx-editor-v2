// A bookmark write the store refuses keeps the store's reason.
//
// The bookmark lane handed `commit` only whether its write committed, so the store's answer
// never reached the report: `exec` answered "insertBookmark was refused" for a write inside a
// locked content control. The reason is now the document's own.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../contracts/editor.ts';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor } from '../docx-editor.ts';
import { docx } from './paginated-surface-fixtures.ts';

const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const BODY =
  `<w:p>${t('Free words.')}</w:p>` +
  '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>' +
  `<w:p><w:bookmarkStart w:id="1" w:name="Inside"/>${t('Locked words.')}` +
  '<w:bookmarkEnd w:id="1"/></w:p></w:sdtContent></w:sdt>';

type Editor = ReturnType<typeof createDocxEditor>;

function withEditor(run: (editor: Editor) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(BODY), author: 'Ada' });
  try {
    run(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

const partXml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());

function select(editor: Editor, index: number, start: number, end: number): void {
  const paragraphId = editor.surface!.session.paragraphIds()[index]!;
  const range = { anchor: { paragraphId, offset: start }, head: { paragraphId, offset: end } };
  expect(editor.exec({ type: 'setSelection', range } as EditorCommand).ok).toBe(true);
}

describe('a refused bookmark write keeps the store reason', () => {
  test('insertBookmark inside a locked content control', () => {
    withEditor((editor) => {
      select(editor, 1, 0, 6);
      const before = partXml(editor);
      expect(editor.exec({ type: 'insertBookmark', name: 'Added' })).toEqual({
        ok: false,
        code: 'unsupported',
        reason: 'this content is locked',
      });
      expect(partXml(editor)).toBe(before);
    });
  });

  test('a bookmark on free text still inserts', () => {
    withEditor((editor) => {
      select(editor, 0, 0, 4);
      expect(editor.exec({ type: 'insertBookmark', name: 'Added' })).toEqual({
        ok: true,
        changed: true,
      });
    });
  });
});
