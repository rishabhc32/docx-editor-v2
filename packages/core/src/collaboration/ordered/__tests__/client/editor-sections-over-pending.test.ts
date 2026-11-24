// The organiser's direct section edits over another person's pending section record, with real
// editors in one ordered room and in the editor alone. Cy suggests margins for the whole
// document; Olive then breaks a section and undoes it, or sets the page size and decides Cy's
// suggestion. Every copy ends as the room, and the room's save says what the editor's own says:
// each section's properties, and who holds its change record.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { collectSectionPropertyNodes } from '../../../../store/package/hf-references.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  type Editor,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './sections-fixtures.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('First paragraph.')) +
    p(t('Second paragraph.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708"' +
    ' w:footer="708" w:gutter="0"/><w:pgNumType w:start="5"/><w:cols w:num="2" w:space="720"/>' +
    '<w:docGrid w:linePitch="360"/></w:sectPr>'
);

/** Each section of a save: its properties, then the author of its change record, if any. */
function sectionsOf(bytes: Uint8Array): string[][] {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  const main = read.package.parts.get(read.package.mainDocumentPart);
  if (!main) throw new Error('no main part');
  return collectSectionPropertyNodes(main.root).map((section) =>
    (section?.children ?? []).flatMap((child) => {
      if (child.kind === 'textValue') {
        return [];
      }

      const attributes = child.attributes
        .filter((a) => child.localName !== 'sectPrChange' || a.localName === 'author')
        .map((a) => `${a.localName}=${a.value}`)
        .sort();
      return [[child.localName, ...attributes].join(' ')];
    })
  );
}

/** Decide the first revision card, as the organiser does. */
function decide(editor: Editor, decision: 'accept' | 'reject'): void {
  const [card] = editor.getReviewItems().filter((item) => item.kind === 'revision');
  if (!card) throw new Error('no card');
  const done =
    decision === 'accept' ? editor.acceptReviewItem(card.key) : editor.rejectReviewItem(card.key);
  expect(done.ok).toBe(true);
}

/** Olive's steps after Cy's suggestion, on her editor. */
const STEPS: Readonly<Record<string, (editor: Editor) => void>> = {
  'a section break, undone': (editor) => {
    expect(editor.surface?.insertSectionBreak('nextPage')).toBe(true);
    editor.surface?.undo();
  },
  "the page size, then Cy's margins rejected": (editor) => {
    const page = { pageWidthTwips: 11906, pageHeightTwips: 16838 };
    expect(editor.surface?.setSectionProperties(page)).toBe(true);
    decide(editor, 'reject');
  },
  "the page size, then Cy's margins accepted": (editor) => {
    const page = { pageWidthTwips: 11906, pageHeightTwips: 16838 };
    expect(editor.surface?.setSectionProperties(page)).toBe(true);
    decide(editor, 'accept');
  },
};

const MARGINS = { marginTopTwips: 2880 };

/** The editor alone: Cy's suggestion, then Olive's steps with tracking off; its save. */
async function alone(step: (editor: Editor) => void): Promise<Uint8Array> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Cy',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    expect(editor.surface?.setSectionProperties(MARGINS)).toBe(true);
    editor.surface?.setAuthor('Olive');
    editor.surface?.setEditingMode('edit');
    const [first] = editor.surface?.session.paragraphIds() ?? [];
    if (first === undefined) throw new Error('no paragraph');
    editor.surface?.setSelection({
      anchor: { paragraphId: first, offset: 5 },
      head: { paragraphId: first, offset: 5 },
    });
    step(editor);
    return new Uint8Array(await editor.save());
  } finally {
    editor.destroy();
    container.remove();
  }
}

describe("the organiser's direct section edits over another person's pending record", () => {
  for (const [label, step] of Object.entries(STEPS)) {
    test(`${label}: the room says what the editor alone says`, async () => {
      const own = sectionsOf(await alone(step));
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const review = reviewTestModule();
      const cy = await join(room, 'Cy', { wordFile: FILE, review });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: FILE,
        review,
      });
      expect(cy.editor.surface?.setSectionProperties(MARGINS)).toBe(true);
      await settle();
      converged(room, cy, olive);
      caretAt(olive, 0, 5);
      step(olive.editor);
      await settle();
      converged(room, cy, olive);
      for (const person of [cy, olive]) {
        expect(sectionsOf(new Uint8Array(await person.editor.save()))).toEqual(own);
      }
    });
  }
});
