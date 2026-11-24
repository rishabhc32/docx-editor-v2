// A page setup over a suggested section break, with real editors in one ordered room and in the
// editor alone. Olive suggests a new paragraph, Ana suggests a landscape page for the whole
// document, and Olive suggests a continuous section break at the new paragraph, so Ana's change
// record is on Olive's new section too. Ana's card holds her page alone: its reject leaves
// Olive's break, as the editor alone and Word (§17.13.5.32) leave it. The room once decided the
// break with Ana's card.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Editor,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './sections-fixtures.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('One.')) +
    p(t('Two.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"' +
    ' w:footer="708" w:gutter="0"/></w:sectPr>'
);

/** The authors of an editor's revision cards. */
const authors = (editor: Editor): string[] =>
  editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [item.author ?? ''] : []))
    .sort();

/** The body an editor saves, without ids and dates. */
async function bodyOf(editor: Editor): Promise<string> {
  const files = unzipSync(new Uint8Array(await editor.save()));
  return strFromU8(files['word/document.xml'] ?? new Uint8Array()).replace(
    / (?:w:id|w:date|w14:paraId|w14:textId)="[^"]*"/g,
    ''
  );
}

/** Reject Ana's card on an editor. */
function rejectAna(editor: Editor): void {
  const card = editor
    .getReviewItems()
    .find((item) => item.kind === 'revision' && item.author === 'Ana');
  if (!card) {
    throw new Error("no card of Ana's");
  }

  expect(editor.rejectReviewItem(card.key).ok).toBe(true);
}

/** The editor alone: the three suggestions and the reject of Ana's card; its body. */
async function alone(): Promise<string> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Olive',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const s = editor.surface;
    const [first] = s?.session.paragraphIds() ?? [];
    if (!s || first === undefined) {
      throw new Error('no paragraph');
    }

    const caret = { paragraphId: first, offset: 2 };
    s.setSelection({ anchor: caret, head: caret });
    s.splitParagraph();
    s.setAuthor('Ana');
    expect(s.setSectionProperties({ orientation: 'landscape' })).toBe(true);
    s.setAuthor('Olive');
    s.setSelection({ anchor: caret, head: caret });
    expect(s.insertSectionBreak('continuous')).toBe(true);
    expect(authors(editor)).toEqual(['Ana', 'Olive']);
    rejectAna(editor);
    expect(authors(editor)).toEqual(['Olive']);
    return await bodyOf(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

describe('a page setup over a suggested section break', () => {
  test("Ana's reject leaves Olive's break, as in the editor alone", async () => {
    const own = await alone();
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const review = reviewTestModule();
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    const olive = await join(room, 'Olive', { login: ORGANISER, wordFile: FILE, review });
    caretAt(olive, 0, 2);
    surface(olive).splitParagraph();
    await settle();
    expect(surface(ana).setSectionProperties({ orientation: 'landscape' })).toBe(true);
    await settle();
    caretAt(olive, 0, 2);
    expect(surface(olive).insertSectionBreak('continuous')).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(authors(olive.editor)).toEqual(['Ana', 'Olive']);
    rejectAna(olive.editor);
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(authors(person.editor)).toEqual(['Olive']);
      expect(await bodyOf(person.editor)).toBe(own);
    }
  });
});
