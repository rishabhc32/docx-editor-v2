// A write over a selection whose end the published pages do not hold.
//
// A deferred layout pass publishes the pages after the screen from the PREVIOUS layout, so a
// paragraph added there — by another participant, or by the API — is in no published order
// until the pass finishes. Formatting, indenting or deleting a range that ends in it did
// nothing and reported nothing. The write now finishes that layout first. A range that ends
// in text the page never shows (a paragraph hidden whole) is refused by name instead.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, beforeEach, describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../contracts/editor.ts';
import { documentOrder } from '../../layout/document-order.ts';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor } from '../docx-editor.ts';
import { setDeferredLayoutForTest } from '../surface-deferred-layout.ts';
import { UNPLACED_SELECTION_REFUSAL } from '../surface-range-edit.ts';
import { docx } from './paginated-surface-fixtures.ts';

const t = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

type Editor = ReturnType<typeof createDocxEditor>;

function withEditor(body: string, run: (editor: Editor) => void): void {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(body), author: 'Ada' });
  try {
    run(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

const partXml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());
const paragraphsXml = (editor: Editor) => partXml(editor).match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];

function selectRange(editor: Editor, from: string, to: string, toOffset: number): void {
  const range = {
    anchor: { paragraphId: from, offset: 0 },
    head: { paragraphId: to, offset: toOffset },
  };
  expect(editor.exec({ type: 'setSelection', range } as EditorCommand).ok).toBe(true);
}

/** Each write, and what it must leave in the last paragraph's XML. */
const WRITES: readonly { name: string; command: EditorCommand; lastHas: RegExp }[] = [
  { name: 'toggleMark', command: { type: 'toggleMark', mark: 'bold' }, lastHas: /<w:b\/>/ },
  { name: 'setIndent', command: { type: 'setIndent', left: 720 }, lastHas: /<w:ind / },
  { name: 'setAlignment', command: { type: 'setAlignment', align: 'center' }, lastHas: /<w:jc / },
  {
    name: 'toggleList',
    command: { type: 'toggleList', kind: 'bullet' } as EditorCommand,
    lastHas: /<w:numPr>/,
  },
];

describe('a range that ends on a page the deferred pass has not finished', () => {
  const BODY = Array.from(
    { length: 400 },
    (_, index) => `<w:p>${t(`Paragraph ${index}`)}</w:p>`
  ).join('');
  let restore: () => void = () => undefined;
  beforeEach(() => {
    restore = setDeferredLayoutForTest(1, 0);
  });
  afterEach(() => restore());

  /** Split the first and the last paragraph in one commit the surface did not make. */
  function addLastParagraph(editor: Editor): { first: string; tail: string } {
    const surface = editor.surface!;
    const ids = surface.session.paragraphIds();
    surface.session.applyTreeOps([
      { op: 'splitParagraph', paragraphId: ids[0]!, offset: 4 },
      { op: 'splitParagraph', paragraphId: ids[ids.length - 1]!, offset: 4 },
    ]);
    const after = surface.session.paragraphIds();
    const tail = after[after.length - 1]!;
    // The precondition the row is about: the new paragraph is in no published order yet.
    expect(documentOrder(surface.layout()).includes(tail)).toBe(false);
    return { first: after[0]!, tail };
  }

  for (const write of WRITES) {
    test(`${write.name} reaches the new paragraph`, () => {
      withEditor(BODY, (editor) => {
        const { first, tail } = addLastParagraph(editor);
        selectRange(editor, first, tail, 3);
        expect(editor.exec(write.command)).toEqual({ ok: true, changed: true });
        expect(paragraphsXml(editor).at(-1)).toMatch(write.lastHas);
      });
    });
  }

  test('deleteText deletes through the new paragraph', () => {
    withEditor(BODY, (editor) => {
      const { first, tail } = addLastParagraph(editor);
      selectRange(editor, first, tail, 3);
      expect(editor.exec({ type: 'deleteText' } as EditorCommand)).toEqual({
        ok: true,
        changed: true,
      });
      // The tail is "graph 399" (the last paragraph split after "Para"); 3 of it are gone.
      expect(editor.surface!.session.storyText({ kind: 'body' })).toBe('ph 399');
    });
  });

  test('typing over the range replaces it', () => {
    withEditor(BODY, (editor) => {
      const { first, tail } = addLastParagraph(editor);
      selectRange(editor, first, tail, 3);
      expect(editor.exec({ type: 'insertText', text: 'X' } as EditorCommand).ok).toBe(true);
      expect(editor.surface!.session.storyText({ kind: 'body' })).toBe('Xph 399');
    });
  });
});

describe('a range that ends in text the page does not show', () => {
  const HIDDEN =
    `<w:p>${t('One')}</w:p>` +
    `<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>${t('Hidden', '<w:vanish/>')}</w:p>` +
    `<w:p>${t('Three')}</w:p>`;

  const refused: readonly { name: string; command: EditorCommand }[] = [
    ...WRITES,
    { name: 'deleteText', command: { type: 'deleteText' } as EditorCommand },
    { name: 'insertText', command: { type: 'insertText', text: 'X' } as EditorCommand },
    { name: 'clearFormatting', command: { type: 'clearFormatting' } as EditorCommand },
  ];
  for (const write of refused) {
    test(`${write.name} is refused by name and changes nothing`, () => {
      withEditor(HIDDEN, (editor) => {
        const ids = editor.surface!.session.paragraphIds();
        expect(documentOrder(editor.surface!.layout()).includes(ids[1]!)).toBe(false);
        selectRange(editor, ids[0]!, ids[1]!, 2);
        const before = partXml(editor);
        const result = editor.exec(write.command);
        expect(result.ok).toBe(false);
        expect(result.ok ? null : result.reason).toBe(UNPLACED_SELECTION_REFUSAL);
        expect(partXml(editor)).toBe(before);
      });
    });
  }

  test('a range inside the shown text still writes', () => {
    withEditor(HIDDEN, (editor) => {
      const ids = editor.surface!.session.paragraphIds();
      selectRange(editor, ids[0]!, ids[2]!, 2);
      expect(editor.exec({ type: 'toggleMark', mark: 'bold' })).toEqual({
        ok: true,
        changed: true,
      });
    });
  });
});
