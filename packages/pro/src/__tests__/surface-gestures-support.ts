/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A user's press and keys on a mounted editor, through the same events a browser sends.

import type { DocxEditorInstance } from '@docx-editor.dev/core/editor';

/** The painted pages, which take the pointer and the keyboard. */
export function pagesOf(container: HTMLElement): HTMLElement {
  const pages = container.querySelector<HTMLElement>('.docx-pages');
  if (!pages) throw new Error('no painted pages');
  return pages;
}

/** A press on the first painted fragment of the document's `index`th content control. */
export function pressContentControl(
  editor: DocxEditorInstance,
  container: HTMLElement,
  index = 0
): void {
  const surface = editor.surface!;
  const pages = pagesOf(container);
  // happy-dom lays nothing out, so the one rectangle the pointer controller reads is given.
  Object.defineProperty(pages, 'getBoundingClientRect', {
    configurable: true,
    value: () => ({ left: 0, top: 0, right: 2000, bottom: 2000, width: 2000, height: 2000 }),
  });
  const layout = surface.layout();
  const fragment = layout.contentControls![index]!.fragments[0]!;
  const page = layout.pages[fragment.pageIndex]!;
  const { paintScale, pageOffsetX } = surface.overlayCoordinates();
  const x =
    page.box.x + (pageOffsetX.get(page.index) ?? 0) + page.contentBox.x + fragment.box.x + 10;
  const y = page.box.y + page.contentBox.y + fragment.box.y + fragment.box.height / 2;
  const init = {
    bubbles: true,
    cancelable: true,
    button: 0,
    pointerId: 1,
    pointerType: 'mouse',
    clientX: x * paintScale,
    clientY: y * paintScale,
  };
  pages.dispatchEvent(new PointerEvent('pointerdown', init));
  pages.ownerDocument.dispatchEvent(new PointerEvent('pointerup', init));
}

/** Keys as the browser sends them: one `beforeinput` each, landed before the next one. */
export function typeKeys(editor: DocxEditorInstance, container: HTMLElement, text: string): void {
  for (const data of text) {
    pagesOf(container).dispatchEvent(
      new InputEvent('beforeinput', {
        bubbles: true,
        cancelable: true,
        inputType: 'insertText',
        data,
      })
    );
    editor.surface!.flushPendingInput();
  }
}

/** Backspace, as the browser sends it. */
export function pressBackspace(editor: DocxEditorInstance, container: HTMLElement): void {
  pagesOf(container).dispatchEvent(
    new InputEvent('beforeinput', {
      bubbles: true,
      cancelable: true,
      inputType: 'deleteContentBackward',
    })
  );
  editor.surface!.flushPendingInput();
}
