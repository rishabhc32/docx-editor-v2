// INSERT › TABLE LEAVES THE USER TYPING IN THE FIRST CELL.
//
// The host control that ran the command holds DOM focus, and a keystroke there never reaches
// the pages, so the command has to take focus as well as place the caret.

import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { semanticSelectionFromDom } from '../dom-selection.ts';
import { docx, paragraph, putCaret } from './paginated-surface-fixtures.ts';

let editor: DocxEditorInstance | undefined;
let container: HTMLElement | undefined;
let hostButton: HTMLButtonElement | undefined;

afterEach(() => {
  editor?.destroy();
  container?.remove();
  hostButton?.remove();
  document.getSelection()?.removeAllRanges();
});

function mountWithHostFocus() {
  container = document.createElement('div');
  document.body.append(container);
  editor = createDocxEditor({ container, document: docx(paragraph('hello')) });
  const surface = editor.surface!;
  surface.focus();
  putCaret(surface, 5);
  // The host's Insert › Table control takes focus on click, as a grid cell or menu item does.
  hostButton = document.createElement('button');
  document.body.append(hostButton);
  hostButton.focus();
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  return { editor, surface, pages };
}

test('a table inserted from host chrome takes the next keystroke in its first cell', async () => {
  const { editor, surface, pages } = mountWithHostFocus();
  expect(document.activeElement).toBe(hostButton);

  expect(editor.exec({ type: 'insertTable', rows: 2, cols: 2 })).toMatchObject({ ok: true });

  const firstCell = surface.session.paragraphIds()[0]!;
  const caret = { paragraphId: firstCell, offset: 0 };
  expect(surface.state().selection).toEqual({ anchor: caret, head: caret });
  expect(document.activeElement).toBe(pages);
  expect(semanticSelectionFromDom(pages, document.getSelection())).toEqual({
    anchor: caret,
    head: caret,
  });

  // The repaint, the focus move and the selection write each queue a `selectionchange`.
  // They arrive before the keystroke and must leave the caret where the insert put it.
  document.dispatchEvent(new Event('selectionchange'));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(surface.state().selection).toEqual({ anchor: caret, head: caret });

  // The browser delivers the keystroke to whatever holds focus.
  document.activeElement!.dispatchEvent(
    new InputEvent('beforeinput', {
      bubbles: true,
      cancelable: true,
      inputType: 'insertText',
      data: 'X',
    })
  );
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(editor.query({ type: 'paragraphs' }).map(({ text }) => text)).toEqual([
    'X',
    '',
    '',
    '',
    'hello',
  ]);
});
