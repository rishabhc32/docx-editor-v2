// `exec({ type: 'insertText' })` reports every refusal of its write.
//
// It reported only two reasons, a table of contents and a field code, and answered
// `{ ok: true, changed: false }` for any other refused write. Text that never landed then looked
// like a write that changed nothing. A refused edit must say so.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../docx-editor.ts';
import { docx } from './paginated-surface-fixtures.ts';

const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const LOCKED =
  '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="7"/></w:sdtPr>' +
  `<w:sdtContent><w:p>${t('Locked words.')}</w:p></w:sdtContent></w:sdt>` +
  `<w:p>${t('Free words.')}</w:p>`;

function withEditor(run: (editor: ReturnType<typeof createDocxEditor>) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(LOCKED), author: 'Ada' });
  try {
    run(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

const caret = (editor: ReturnType<typeof createDocxEditor>, index: number, offset: number) => {
  const paragraphId = editor.surface!.session.paragraphIds()[index]!;
  const at = { paragraphId, offset };
  expect(editor.exec({ type: 'setSelection', range: { anchor: at, head: at } }).ok).toBe(true);
};

describe('exec insertText reports a refused write', () => {
  test('typing in a control whose content is locked is refused, not "unchanged"', () => {
    withEditor((editor) => {
      caret(editor, 0, 3);
      const result = editor.exec({ type: 'insertText', text: 'X' });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.code).toBe('unsupported');
        expect(result.reason).toBe(editor.surface!.state().lastRejection!);
      }
      expect(editor.surface!.session.storyText({ kind: 'body' })).toBe(
        'Locked words.\nFree words.'
      );
    });
  });

  test('typing in a free paragraph still lands', () => {
    withEditor((editor) => {
      caret(editor, 1, 4);
      expect(editor.exec({ type: 'insertText', text: 'X' })).toEqual({ ok: true, changed: true });
    });
  });
});
