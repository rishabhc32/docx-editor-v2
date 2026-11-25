// Words that hold a note reference, moved or struck (design D20). A move copies the reference,
// and both halves name the one note: a reject keeps the note with the source, an accept keeps it
// with the copy. A fold reads the copy as a reference to the file's own note, so every decision
// after a fold gives the Word file the same decision gives live, and the cards read the same.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile, type OrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, DIRECT, EDIT, FOOTNOTES, markOf } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { docx, MATRIX_EXTRAS, p, t } from './word-fixtures.ts';

const NOTE_EDIT = { part: FOOTNOTES, mode: 'suggest' } as const;
const FOOTNOTE =
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
const ENDNOTE = '<w:r><w:endnoteReference w:id="1"/></w:r>';
const COMMENT_START = '<w:commentRangeStart w:id="0"/>';
const COMMENT_END = '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>';

/** "Payment", the reference, " Terms apply here.", and a second paragraph. */
const fileWith = (reference: string, before = '', after = '') =>
  docx(
    p(t('Payment') + before + reference + t(' Terms') + after + t(' apply here.')) +
      p(t('The second paragraph.')),
    MATRIX_EXTRAS
  );

/** Olive's move of the reference and " Terms" after it into the second paragraph. */
const moveTerms = (doc: OrderedDocument) =>
  apply(
    doc,
    'olive:1',
    { kind: 'move', ...EDIT, from: at(doc, 0, 7), to: at(doc, 0, 14), place: at(doc, 1, 4) },
    'Olive'
  ).suggestions;

/** Each note reference of the main document, and each note with its text, as the file holds. */
function notesOf(file: OrderedWordFile, blobs: ReturnType<typeof mediaStore>): string[] {
  const parts = unzipSync(orderedPackageBytes(file.document, file.base, { blobs }));
  const xml = (name: string) => strFromU8(parts[name] ?? new Uint8Array());
  const references = [
    ...xml('word/document.xml').matchAll(/<w:(footnote|endnote)Reference w:id="(\d+)"/g),
  ];
  const notes = [
    ...xml('word/footnotes.xml').matchAll(/<w:(footnote) w:id="(\d+)">([\s\S]*?)<\/w:footnote>/g),
    ...xml('word/endnotes.xml').matchAll(/<w:(endnote) w:id="(\d+)">([\s\S]*?)<\/w:endnote>/g),
  ];
  return [
    ...references.map((m) => `${m[1]} reference ${m[2]}`),
    ...notes.map((m) => `${m[1]} ${m[2]}:${m[3]!.replace(/<[^>]+>/g, '')}`),
  ];
}

/**
 * The notes after `status` decides every suggestion `edit` made: live, after one fold and after
 * two. Each copy is a fresh room.
 */
async function decidedNotes(
  file: Uint8Array,
  edit: (doc: OrderedDocument) => readonly string[],
  status: 'accepted' | 'rejected'
): Promise<string[][]> {
  const out: string[][] = [];
  for (const folds of [0, 1, 2]) {
    let room = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(room.base);
    const ids = edit(room.document);
    for (let k = 0; k < folds; k += 1) {
      room = await folded(room.document, room.base, room.document.file + 1, blobs);
    }

    for (const id of ids) {
      const seen = orderedCardSeen(room.document, id);
      if (seen) {
        apply(room.document, 'olga:1', { kind: 'decide', cards: [{ suggestions: seen }], status });
      }
    }

    out.push(notesOf(room, blobs));
  }

  return out;
}

/** Check the fold, and that each decision gives `expected` live and after each fold. */
async function checkNotes(
  file: Uint8Array,
  edit: (doc: OrderedDocument) => readonly string[],
  expected: Record<'accepted' | 'rejected', string[]>
): Promise<void> {
  const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
  edit(doc);
  await checkFold(doc, base, mediaStore(base), 'a note reference in moved or struck words');
  for (const status of ['accepted', 'rejected'] as const) {
    const want = expected[status];
    expect({ status, notes: await decidedNotes(file, edit, status) }).toEqual({
      status,
      notes: [want, want, want],
    });
  }
}

// The file's notes: the footnote and the endnote, which no reference cites in the footnote cases
const FEES = 'footnote 1: Fees are due net of tax.';
const ANNEX = 'endnote 1: See the annex.';

describe('a decision on a move of words that hold a note reference, live and after a fold', () => {
  test('a footnote reference: a reject keeps the note, an accept keeps one note', async () => {
    await checkNotes(fileWith(FOOTNOTE), moveTerms, {
      accepted: ['footnote reference 1', FEES, ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });

  test('an endnote reference', async () => {
    await checkNotes(fileWith(ENDNOTE), moveTerms, {
      accepted: ['endnote reference 1', FEES, ANNEX],
      rejected: ['endnote reference 1', FEES, ANNEX],
    });
  });

  test("a paragraph's every word", async () => {
    const file = docx(
      p(t('Payment') + FOOTNOTE + t(' Terms.')) + p(t('The second paragraph.')),
      MATRIX_EXTRAS
    );
    const moveAll = (doc: OrderedDocument) =>
      apply(
        doc,
        'olive:1',
        { kind: 'move', ...EDIT, from: at(doc, 0, 0), to: at(doc, 0, 15), place: at(doc, 1, 0) },
        'Olive'
      ).suggestions;
    await checkNotes(file, moveAll, {
      accepted: ['footnote reference 1', FEES, ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });

  test('a comment anchor around the moved words', async () => {
    await checkNotes(fileWith(FOOTNOTE, COMMENT_START, COMMENT_END), moveTerms, {
      accepted: ['footnote reference 1', FEES, ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });

  // Bob strikes the words around Olive's source: the note shows while the copy does
  test("another person's deletion over the move's source", async () => {
    const struck = (doc: OrderedDocument) => [
      ...moveTerms(doc),
      ...apply(
        doc,
        'bob:1',
        { kind: 'delete', ...EDIT, from: at(doc, 0, 0), to: at(doc, 0, 25) },
        'Bob'
      ).suggestions,
    ];
    await checkNotes(fileWith(FOOTNOTE), struck, {
      accepted: ['footnote reference 1', FEES, ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });
});

describe("the organiser's direct deletion of a moved copy of a note reference", () => {
  // Ana's pending deletion in the note: the move's struck source still names the note, so it
  // stays whole, as the editor alone keeps it, and goes with the source on an accept
  test('the note stays whole while the struck source names it', async () => {
    const edit = (doc: OrderedDocument) => {
      const struck = { from: at(doc, 2, 14, FOOTNOTES), to: at(doc, 2, 18, FOOTNOTES) };
      const anas = apply(doc, 'ana:1', { kind: 'delete', ...NOTE_EDIT, ...struck }, 'Ana');
      const moved = moveTerms(doc);
      apply(doc, 'olive:2', { kind: 'delete', ...DIRECT, from: at(doc, 1, 4), to: at(doc, 1, 5) });
      return [...anas.suggestions, ...moved];
    };
    await checkNotes(fileWith(FOOTNOTE), edit, {
      accepted: [ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });
});

// Ana centres the note's paragraph; the organiser's direct deletion takes the move's struck source,
// so only the copy cites the note. The note stays the file's own (a moved reference names its
// source's note), and Ana's change stays a card of its own, as the editor alone keeps them. A
// reject takes the copy, and the note goes with its last reference.
describe("the organiser's direct deletion of a move's struck source", () => {
  test("the note stays the file's own, and its change a card of its own", async () => {
    const edit = (doc: OrderedDocument) => {
      const centred = apply(
        doc,
        'ana:1',
        {
          kind: 'paragraph-property',
          ...NOTE_EDIT,
          marks: [markOf(doc, 2, FOOTNOTES)],
          properties: [{ localName: 'jc', attributes: { val: 'center' } }],
        },
        'Ana'
      );
      const moved = moveTerms(doc);
      apply(doc, 'olive:2', { kind: 'delete', ...DIRECT, from: at(doc, 0, 7), to: at(doc, 0, 14) });
      return [...centred.suggestions, ...moved];
    };
    await checkNotes(fileWith(FOOTNOTE), edit, {
      accepted: ['footnote reference 1', FEES, ANNEX],
      rejected: [ANNEX],
    });
  });
});

describe('a decision on a deletion of words that hold a note reference, after a fold', () => {
  test('an accept removes the note with its reference, a reject keeps both', async () => {
    const strike = (doc: OrderedDocument) =>
      apply(
        doc,
        'olive:1',
        { kind: 'delete', ...EDIT, from: at(doc, 0, 7), to: at(doc, 0, 14) },
        'Olive'
      ).suggestions;
    await checkNotes(fileWith(FOOTNOTE), strike, {
      accepted: [ANNEX],
      rejected: ['footnote reference 1', FEES, ANNEX],
    });
  });
});
