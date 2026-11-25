// Word's own notes (`w:type` separator, continuationSeparator and continuationNotice): the editor
// alone shows none of them to edit, so a change that would write into one is refused as a place
// in a note that does not show, live and after a fold.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile, type OrderedWordFile } from '../client/open.ts';
import { at, FOOTNOTES, refusal } from './change-apply-fixture.ts';
import { folded, mediaStore } from './fold-check.ts';
import { docx, MATRIX_EXTRAS, p, R, separators, t, W } from './word-fixtures.ts';

const NOTE = { part: FOOTNOTES, mode: 'suggest' } as const;

/** Footnotes: the two separators, a continuation notice, and the note the body cites. */
const FILE = docx(p(t('The fee') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + t(' is due.')), {
  ...MATRIX_EXTRAS,
  parts: {
    ...MATRIX_EXTRAS.parts,
    'word/footnotes.xml':
      `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${separators('footnote')}` +
      `<w:footnote w:type="continuationNotice" w:id="1">${p(t('Continued.'))}</w:footnote>` +
      `<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>${t(' Fees are net.')}</w:p>` +
      '</w:footnote></w:footnotes>',
  },
});

/** The paragraphs of Word's own notes in the footnotes story, and the cited note's. */
const WORD_NOTES = [0, 1, 2];
const CITED = 3;

/** The file opened live, or after one fold. */
async function opened(folds: number): Promise<OrderedWordFile> {
  const file = openOrderedWordFile(FILE, { file: 1 });
  if (folds === 0) {
    return file;
  }

  return folded(file.document, file.base, 2, mediaStore(file.base));
}

describe("a change into Word's own notes", () => {
  for (const folds of [0, 1]) {
    const when = folds === 0 ? 'live' : 'after a fold';

    test(`typing, a deletion and a move into one are refused, ${when}`, async () => {
      const { document: doc } = await opened(folds);
      for (const index of WORD_NOTES) {
        const place = at(doc, index, 0, FOOTNOTES);
        expect(refusal(doc, 'ana:1', { kind: 'insert', ...NOTE, place, text: 'X' })).toBe(
          'hidden-target'
        );
        const to = at(doc, index, 1, FOOTNOTES);
        expect(refusal(doc, 'ana:2', { kind: 'delete', ...NOTE, from: place, to })).toBe(
          'hidden-target'
        );
        const from = at(doc, CITED, 1, FOOTNOTES);
        const end = at(doc, CITED, 5, FOOTNOTES);
        expect(refusal(doc, 'ana:3', { kind: 'move', ...NOTE, from, to: end, place })).toBe(
          'hidden-target'
        );
      }

      // The note the body cites still takes them
      const place = at(doc, CITED, 1, FOOTNOTES);
      expect(refusal(doc, 'ana:4', { kind: 'insert', ...NOTE, place, text: 'X' })).toBe('applied');
    });
  }
});
