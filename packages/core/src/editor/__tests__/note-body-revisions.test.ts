// A deleted note in suggesting mode, as Word writes it: the reference struck, and the note's body
// marked deleted under the same author and date (Word 2016 files `RP050-Deleted-Footnote.docx`
// and `tdf141660.docx`). A new note: the reference inserted, and the note's runs in an insertion
// of their own under the same author and date (Word 16 for Mac, 2026-10-29). One review card
// decides each pair.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { collectReviewItems, findNode, revisionItemsOf } from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const reference = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;

const BODY =
  `<w:p><w:r><w:t>Alpha</w:t></w:r>${reference(1)}<w:r><w:t xml:space="preserve"> one</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>Beta</w:t></w:r>${reference(2)}</w:p>`;

const note = (id: number, text: string) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
  `<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p></w:footnote>`;

const FOOTNOTES =
  `<w:footnotes xmlns:w="${W}">` +
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  `${note(1, 'First note')}${note(2, 'Second note')}</w:footnotes>`;

function docx(body = BODY, footnotes = FOOTNOTES): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
        `<Override PartName="/word/footnotes.xml" ContentType="${WML}.footnotes+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OFFICE}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${OFFICE}/footnotes" Target="footnotes.xml"/></Relationships>`
    ),
    'word/footnotes.xml': strToU8(footnotes),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

function reviewModule(): EditorModule {
  return {
    id: 'review',
    review: {
      displayModes: ['all-markup', 'proposed', 'original'],
      collectReviewItems,
      revisionItemsOfParagraph: (part, paragraphId) => {
        const paragraph = findNode(part, paragraphId);
        if (!paragraph || paragraph.kind !== 'paragraph') return [];
        return revisionItemsOf({ ...part, root: paragraph });
      },
    },
  };
}

async function withEditor(
  bytes: Uint8Array,
  run: (editor: DocxEditorInstance) => Promise<void> | void
): Promise<void> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    author: 'Reviewer',
    modules: [reviewModule()],
  });
  if (!editor.surface) throw new Error('surface failed to mount');
  try {
    await run(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

interface Saved {
  readonly bytes: Uint8Array;
  readonly document: string;
  readonly footnotes: string;
}

async function saved(editor: DocxEditorInstance): Promise<Saved> {
  const bytes = new Uint8Array(await editor.save());
  const files = unzipSync(bytes);
  return {
    bytes,
    document: strFromU8(files['word/document.xml']!),
    footnotes: strFromU8(files['word/footnotes.xml']!),
  };
}

/** The first note's body element, as saved. */
const firstNote = (footnotes: string): string =>
  /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(footnotes)![0];

const attributes = (xml: string, element: string): { author?: string; date?: string }[] =>
  [...xml.matchAll(new RegExp(`<w:${element} ([^>]*?)/?>`, 'g'))].map((m) => ({
    author: /w:author="([^"]*)"/.exec(m[1]!)?.[1],
    date: /w:date="([^"]*)"/.exec(m[1]!)?.[1],
  }));

/** Strike the first note's reference in suggesting mode, as Delete over it does. */
function strikeFirstNote(editor: DocxEditorInstance): void {
  editor.surface!.setEditingMode('suggest');
  expect(editor.surface!.deleteNote('footnote', 1)).toBe(true);
}

const revisionCards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().filter((item) => item.kind === 'revision');

describe('a deleted note in suggesting mode', () => {
  test('marks the note body deleted with the reference, as Word writes it', async () => {
    await withEditor(docx(), async (editor) => {
      strikeFirstNote(editor);
      const after = await saved(editor);
      const [struck] = attributes(after.document, 'del');
      expect(struck?.author).toBe('Reviewer');
      const body = firstNote(after.footnotes);
      // The paragraph mark, then every run (the `w:footnoteRef` run too) in a deletion
      expect(body).toMatch(/<w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/);
      expect(body).toMatch(/<w:del [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/>/);
      expect(body).toContain('<w:delText xml:space="preserve"> First note</w:delText>');
      expect(body).not.toContain('<w:t');
      for (const mark of attributes(body, 'del')) expect(mark).toEqual(struck!);
      // The other note stays as it was
      expect(after.footnotes).toContain(note(2, 'Second note').slice(-80));
      // One card decides the reference and the body
      expect(revisionCards(editor)).toHaveLength(1);
    });
  });

  test('rejecting the card gives the note back whole', async () => {
    await withEditor(docx(), async (editor) => {
      const before = await saved(editor);
      strikeFirstNote(editor);
      const [card] = revisionCards(editor);
      expect(editor.rejectReviewItem(card!.key).ok).toBe(true);
      const after = await saved(editor);
      expect(after.document).toBe(before.document);
      expect(after.footnotes).toBe(before.footnotes);
      expect(revisionCards(editor)).toHaveLength(0);
    });
  });

  test('accepting the card removes the note', async () => {
    await withEditor(docx(), async (editor) => {
      strikeFirstNote(editor);
      const [card] = revisionCards(editor);
      expect(editor.acceptReviewItem(card!.key).ok).toBe(true);
      const after = await saved(editor);
      expect(after.document).not.toContain('w:footnoteReference w:id="1"');
      expect(after.footnotes).not.toContain('w:id="1"');
      expect(after.footnotes).toContain('Second note');
      expect(after.footnotes).not.toContain('w:del ');
      expect(revisionCards(editor)).toHaveLength(0);
    });
  });

  test('Accept All and Reject All decide the note body with its reference', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(docx(), async (editor) => {
        const before = await saved(editor);
        strikeFirstNote(editor);
        const result = editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' });
        expect(result.ok).toBe(true);
        const after = await saved(editor);
        if (action === 'reject') expect(after.footnotes).toBe(before.footnotes);
        else expect(after.footnotes).not.toContain('w:id="1"');
        expect(after.footnotes).not.toContain('w:del ');
      });
    }
  });

  test('undo takes the reference and the body back in one step, and redo strikes both', async () => {
    await withEditor(docx(), async (editor) => {
      const before = await saved(editor);
      strikeFirstNote(editor);
      const struck = await saved(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      expect((await saved(editor)).footnotes).toBe(before.footnotes);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect((await saved(editor)).footnotes).toBe(struck.footnotes);
    });
  });

  test('after save and reopen, the struck note is still one card, and each decision holds', async () => {
    let bytes: Uint8Array | undefined;
    let original: Saved | undefined;
    await withEditor(docx(), async (editor) => {
      original = await saved(editor);
      strikeFirstNote(editor);
      bytes = (await saved(editor)).bytes;
    });
    await withEditor(bytes!, async (editor) => {
      expect(revisionCards(editor)).toHaveLength(1);
      const [card] = revisionCards(editor);
      expect(editor.rejectReviewItem(card!.key).ok).toBe(true);
      // A rejected mark of a loaded file leaves its emptied properties, as for any paragraph
      const footnotes = (await saved(editor)).footnotes.replace('<w:pPr><w:rPr/></w:pPr>', '');
      expect(footnotes).toBe(original!.footnotes);
    });
    await withEditor(bytes!, async (editor) => {
      const [card] = revisionCards(editor);
      expect(editor.acceptReviewItem(card!.key).ok).toBe(true);
      expect((await saved(editor)).footnotes).not.toContain('w:id="1"');
    });
  });
});

/** The saved note `id`'s element. */
const noteNamed = (footnotes: string, id: number): string =>
  new RegExp(`<w:footnote w:id="${id}">[\\s\\S]*?</w:footnote>`).exec(footnotes)?.[0] ?? '';

const ids = (xml: string, element: string): string[] =>
  [...xml.matchAll(new RegExp(`<w:${element} [^>]*?w:id="([^"]*)"`, 'g'))].map((m) => m[1]!);

/** Insert a footnote after "Alpha" in suggesting mode, as Insert Footnote does. */
function suggestNote(editor: DocxEditorInstance): void {
  const surface = editor.surface!;
  surface.setEditingMode('suggest');
  const { paragraphId } = surface.state().selection.anchor;
  surface.setSelection({ anchor: { paragraphId, offset: 5 }, head: { paragraphId, offset: 5 } });
  expect(surface.insertNote('footnote')).toBe(true);
}

describe('a new note in suggesting mode', () => {
  test("writes the note's runs in an insertion with the reference's author and date, as Word does", async () => {
    await withEditor(docx(), async (editor) => {
      suggestNote(editor);
      const after = await saved(editor);
      const [inserted] = attributes(after.document, 'ins');
      expect(inserted?.author).toBe('Reviewer');
      expect(after.document).toMatch(
        /<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="3"\/>/
      );
      const body = noteNamed(after.footnotes, 3);
      // The `w:footnoteRef` run and the text run in one insertion; the paragraph mark unmarked
      expect(body).toMatch(
        /<w:p[^>]*>(<w:pPr>(?:(?!<w:rPr>).)*<\/w:pPr>)?<w:ins [^>]*><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteRef\/><\/w:r>/
      );
      expect(body).not.toContain('<w:rPr><w:ins');
      expect(attributes(body, 'ins')).toEqual([inserted!]);
      // Its own `w:id`, as Word gives it
      expect(ids(body, 'ins')[0]).not.toBe(ids(after.document, 'ins')[0]);
      expect(revisionCards(editor)).toHaveLength(1);
    });
  });

  test('typing in the new note joins its insertion, and one card still decides it', async () => {
    await withEditor(docx(), async (editor) => {
      suggestNote(editor);
      editor.surface!.type('New note');
      const body = noteNamed((await saved(editor)).footnotes, 3);
      expect(body).toMatch(/<w:ins [^>]*>[\s\S]*New note[\s\S]*<\/w:ins>/);
      expect(revisionCards(editor)).toHaveLength(1);
      const [card] = revisionCards(editor);
      expect(editor.acceptReviewItem(card!.key).ok).toBe(true);
      const after = await saved(editor);
      expect(after.document + after.footnotes).not.toContain('<w:ins ');
      expect(noteNamed(after.footnotes, 3)).toContain('New note');
    });
  });

  test('accepting the card accepts the note text; rejecting it removes the note', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(docx(), async (editor) => {
        const before = await saved(editor);
        suggestNote(editor);
        const [card] = revisionCards(editor);
        const decided =
          action === 'accept'
            ? editor.acceptReviewItem(card!.key)
            : editor.rejectReviewItem(card!.key);
        expect(decided.ok).toBe(true);
        const after = await saved(editor);
        expect(revisionCards(editor)).toHaveLength(0);
        expect(after.document + after.footnotes).not.toContain('<w:ins ');
        if (action === 'reject') {
          expect(after.document).toBe(before.document);
          expect(after.footnotes).toBe(before.footnotes);
        } else {
          expect(after.document).toContain('<w:footnoteReference w:id="3"/>');
          expect(noteNamed(after.footnotes, 3)).toContain('<w:footnoteRef/>');
        }
      });
    }
  });

  test('Accept All and Reject All decide the note text with its reference', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(docx(), async (editor) => {
        const before = await saved(editor);
        suggestNote(editor);
        const result = editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' });
        expect(result.ok).toBe(true);
        const after = await saved(editor);
        expect(after.footnotes).not.toContain('<w:ins ');
        if (action === 'reject') expect(after.footnotes).toBe(before.footnotes);
        else expect(noteNamed(after.footnotes, 3)).toContain('<w:footnoteRef/>');
      });
    }
  });

  test('undo takes the reference and the note away in one step, and redo brings both', async () => {
    await withEditor(docx(), async (editor) => {
      const before = await saved(editor);
      suggestNote(editor);
      const made = await saved(editor);
      expect(editor.exec({ type: 'undo' }).ok).toBe(true);
      const undone = await saved(editor);
      expect(undone.document).toBe(before.document);
      expect(undone.footnotes).toBe(before.footnotes);
      expect(editor.exec({ type: 'redo' }).ok).toBe(true);
      expect((await saved(editor)).footnotes).toBe(made.footnotes);
    });
  });

  test('after save and reopen, the new note is still one card, and each decision holds', async () => {
    let bytes: Uint8Array | undefined;
    let original: Saved | undefined;
    await withEditor(docx(), async (editor) => {
      original = await saved(editor);
      suggestNote(editor);
      editor.surface!.type('Kept');
      bytes = (await saved(editor)).bytes;
    });
    await withEditor(bytes!, async (editor) => {
      expect(revisionCards(editor)).toHaveLength(1);
      const [card] = revisionCards(editor);
      expect(editor.rejectReviewItem(card!.key).ok).toBe(true);
      expect((await saved(editor)).footnotes).toBe(original!.footnotes);
    });
    await withEditor(bytes!, async (editor) => {
      const [card] = revisionCards(editor);
      expect(editor.acceptReviewItem(card!.key).ok).toBe(true);
      const after = await saved(editor);
      expect(after.footnotes).not.toContain('<w:ins ');
      expect(noteNamed(after.footnotes, 3)).toContain('Kept');
    });
  });
});

describe("Word's own new note", () => {
  // As Word 16 wrote it (sample `a-new-note-word.docx`): ids 0 and 1, one author and date
  const STAMP = 'w:author="Ana" w:date="2026-10-29T23:17:00Z"';
  const wordBody =
    `<w:p><w:r><w:t>Alpha</w:t></w:r><w:ins w:id="0" ${STAMP}>${reference(1)}</w:ins>` +
    '<w:r><w:t xml:space="preserve"> one</w:t></w:r></w:p>';
  const wordNote = (later = '') =>
    `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:ins w:id="1" ${STAMP}>` +
    '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
    `<w:r><w:t xml:space="preserve"> New note.</w:t></w:r></w:ins>${later}</w:p></w:footnote>`;
  const wordNotes = (later = '') =>
    `<w:footnotes xmlns:w="${W}">` +
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
    `${wordNote(later)}</w:footnotes>`;

  test('is one card, and each decision decides the note text with its reference', async () => {
    for (const action of ['accept', 'reject'] as const) {
      await withEditor(docx(wordBody, wordNotes()), async (editor) => {
        expect(revisionCards(editor)).toHaveLength(1);
        const [card] = revisionCards(editor);
        const decided =
          action === 'accept'
            ? editor.acceptReviewItem(card!.key)
            : editor.rejectReviewItem(card!.key);
        expect(decided.ok).toBe(true);
        const after = await saved(editor);
        expect(after.document + after.footnotes).not.toContain('<w:ins ');
        if (action === 'accept') expect(noteNamed(after.footnotes, 1)).toContain('New note.');
        else expect(after.footnotes).not.toContain('w:id="1"');
      });
    }
  });

  test('words the author typed in it at another time keep a card of their own', async () => {
    const later =
      '<w:ins w:id="2" w:author="Ana" w:date="2026-10-30T09:00:00Z"><w:r><w:t xml:space="preserve"> Later.</w:t></w:r></w:ins>';
    await withEditor(docx(wordBody, wordNotes(later)), async (editor) => {
      expect(revisionCards(editor)).toHaveLength(2);
      const card = revisionCards(editor).find(
        (item) => item.kind === 'revision' && !/Later/.test(JSON.stringify(item))
      );
      expect(editor.acceptReviewItem(card!.key).ok).toBe(true);
      const after = await saved(editor);
      expect(noteNamed(after.footnotes, 1)).toMatch(/<\/w:r><w:ins [^>]*w:id="2"/);
      expect(revisionCards(editor)).toHaveLength(1);
    });
  });
});
