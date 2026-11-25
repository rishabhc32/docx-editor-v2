// Typing the engine refuses is reported, never dropped.
//
// Keys are queued and land on a timer. A snapshot derived while a key was still queued landed
// it from inside the derive (a page read flushes queued typing), and cached the state it read
// before the key as the state after it. The refusal reached the diagnostics, but
// `snapshot().lastRejection` stayed null, so the chrome showed a key that did nothing.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, expect, test } from 'bun:test';
import type { EditorDiagnostic } from '../../contracts/diagnostics.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { docx } from './paginated-surface-fixtures.ts';

const LOCKED =
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r>' +
  '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:text/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>Acme</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:r><w:t xml:space="preserve"> Corp</w:t></w:r></w:p>';
const LOCK_REASON = 'this content is locked';

const editors: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  document.body.replaceChildren();
});

/** An editor with its caret inside the locked control's words, and its diagnostics. */
function openAtLockedWords() {
  const container = document.createElement('div');
  document.body.append(container);
  const rejected: EditorDiagnostic[] = [];
  const editor = createDocxEditor({
    container,
    document: docx(LOCKED),
    modules: [
      {
        id: 'diagnostics',
        onDiagnostic: (event: EditorDiagnostic) => {
          if (event.kind === 'edit.rejected') rejected.push(event);
        },
      },
    ],
  });
  editors.push(editor);
  const surface = editor.surface!;
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: 8 },
    head: { paragraphId, offset: 8 },
  });
  return { editor, surface, container, rejected };
}

async function settle(): Promise<void> {
  for (let k = 0; k < 4; k += 1) await new Promise((resolve) => setTimeout(resolve, 0));
}

test('a refused key is in the snapshot and the diagnostics, and changes nothing', async () => {
  const { editor, surface, container, rejected } = openAtLockedWords();
  container.querySelector('.docx-pages')!.dispatchEvent(
    new InputEvent('beforeinput', {
      bubbles: true,
      cancelable: true,
      inputType: 'insertText',
      data: 'X',
    })
  );
  await settle();
  expect(surface.session.bodyText()).toBe('Name: Acme Corp');
  expect(editor.snapshot().lastRejection).toBe(LOCK_REASON);
  expect(rejected).toEqual([expect.objectContaining({ origin: 'input' })]);
});

test('refused pasted text is in the snapshot and the diagnostics, and changes nothing', () => {
  const { editor, surface, rejected } = openAtLockedWords();
  surface.insertPlainText('X\nY');
  expect(surface.session.bodyText()).toBe('Name: Acme Corp');
  expect(editor.snapshot().lastRejection).toBe(LOCK_REASON);
  expect(rejected).toEqual([expect.objectContaining({ origin: 'input' })]);
});
