// `exec` reports every refused write as a refusal.
//
// `insertText` learned this for row 32. The other write verbs still answered
// `{ ok: true, changed: false }` when the document turned them down: a delete, a cut, a paste,
// a break, a mark or a paragraph format inside a locked content control looked exactly like a
// write that had nothing to change. A refused edit must say so, and a write that truly had
// nothing to change must still say THAT (row 25).

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../contracts/editor.ts';
import { createDocxEditor } from '../docx-editor.ts';
import { docx } from './paginated-surface-fixtures.ts';

const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const LOCKED =
  '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="7"/></w:sdtPr>' +
  `<w:sdtContent><w:p>${t('Locked words.')}</w:p></w:sdtContent></w:sdt>` +
  `<w:p>${t('Free words.')}</w:p>`;
const TEXT = 'Locked words.\nFree words.';
const LOCKED_REASON = 'this content is locked';

type Editor = ReturnType<typeof createDocxEditor>;

function withEditor(run: (editor: Editor) => void): void {
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

/** Select `start..end` in the index-th paragraph (a caret when they are equal). */
function select(editor: Editor, index: number, start: number, end = start): void {
  const paragraphId = editor.surface!.session.paragraphIds()[index]!;
  const range = {
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  };
  expect(editor.exec({ type: 'setSelection', range }).ok).toBe(true);
}

const bodyText = (editor: Editor) => editor.surface!.session.storyText({ kind: 'body' });

/** Each write verb that answered `ok: true` for a refused write, with a selection it refuses. */
const REFUSED: readonly {
  readonly name: string;
  readonly command: EditorCommand;
  ranged: boolean;
}[] = [
  { name: 'deleteText', command: { type: 'deleteText' }, ranged: true },
  { name: 'cut', command: { type: 'cut' }, ranged: true },
  { name: 'paste', command: { type: 'paste', text: 'P', html: '<b>P</b>' }, ranged: false },
  { name: 'paste (text)', command: { type: 'paste', text: 'P' }, ranged: true },
  {
    name: 'pasteWithoutFormatting',
    command: { type: 'pasteWithoutFormatting', text: 'P' },
    ranged: false,
  },
  { name: 'insertBreak page', command: { type: 'insertBreak', kind: 'page' }, ranged: false },
  { name: 'insertBreak line', command: { type: 'insertBreak', kind: 'line' }, ranged: false },
  { name: 'toggleMark', command: { type: 'toggleMark', mark: 'bold' }, ranged: true },
  { name: 'setAlignment', command: { type: 'setAlignment', align: 'center' }, ranged: false },
  {
    name: 'setParagraphSpacing',
    command: { type: 'setParagraphSpacing', beforePt: 12, afterPt: 12 },
    ranged: false,
  },
  {
    name: 'setParagraphDirection',
    command: { type: 'setParagraphDirection', direction: 'rtl' },
    ranged: false,
  },
  { name: 'setIndent', command: { type: 'setIndent', left: 720 }, ranged: false },
];

describe('exec reports a refused write', () => {
  for (const { name, command, ranged } of REFUSED) {
    test(`${name} in a locked control is refused, not "unchanged"`, () => {
      withEditor((editor) => {
        select(editor, 0, 2, ranged ? 6 : 2);
        expect(editor.exec(command)).toEqual({
          ok: false,
          code: 'unsupported',
          reason: LOCKED_REASON,
        });
        expect(bodyText(editor)).toBe(TEXT);
      });
    });
  }
});

describe('a write with nothing to change stays ok: true, changed: false', () => {
  test('an alignment the paragraph already has', () => {
    withEditor((editor) => {
      select(editor, 1, 2);
      expect(editor.exec({ type: 'setAlignment', align: 'center' }).changed).toBe(true);
      expect(editor.exec({ type: 'setAlignment', align: 'center' })).toEqual({
        ok: true,
        changed: false,
      });
    });
  });

  test('a mark at a caret arms the typing format and changes nothing', () => {
    withEditor((editor) => {
      select(editor, 1, 2);
      expect(editor.exec({ type: 'toggleMark', mark: 'bold' })).toEqual({
        ok: true,
        changed: false,
      });
    });
  });

  test('clearing formatting from plain text', () => {
    withEditor((editor) => {
      select(editor, 1, 0, 4);
      expect(editor.exec({ type: 'clearFormatting' })).toEqual({ ok: true, changed: false });
    });
  });

  test('an earlier refusal does not turn a later unchanged write into a refusal', () => {
    withEditor((editor) => {
      select(editor, 1, 2);
      expect(editor.exec({ type: 'setAlignment', align: 'center' }).changed).toBe(true);
      select(editor, 0, 2, 6);
      expect(editor.exec({ type: 'deleteText' }).ok).toBe(false);
      // The surface still holds the delete's reason: no command clears it on entry.
      expect(editor.surface!.state().lastRejection).toBe(LOCKED_REASON);
      select(editor, 1, 2);
      expect(editor.exec({ type: 'setAlignment', align: 'center' })).toEqual({
        ok: true,
        changed: false,
      });
    });
  });

  test('the same writes outside the locked control land', () => {
    withEditor((editor) => {
      select(editor, 1, 0, 4);
      expect(editor.exec({ type: 'toggleMark', mark: 'bold' })).toEqual({
        ok: true,
        changed: true,
      });
      expect(editor.exec({ type: 'deleteText' })).toEqual({ ok: true, changed: true });
      expect(bodyText(editor)).toBe('Locked words.\n words.');
    });
  });
});
