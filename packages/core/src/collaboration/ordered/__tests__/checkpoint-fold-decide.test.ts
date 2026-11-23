// A decision after a fold gives the Word file the same decision gives in the room (design D14).
// A pending suggestion splits the run it sits in, and a fold writes each piece as a run of its
// own; a decision that takes the suggestion away draws the room's run whole again. A new list's
// definition goes with its suggestion.

import { describe, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, EDIT, markOf } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { docx, MATRIX_EXTRAS, p, type PackageExtras } from './word-fixtures.ts';

/** A run as Word writes it, with its revision session and its own properties. */
const run = (text: string, rPr = '') =>
  `<w:r w:rsidR="00A96A35"><w:rPr><w:sz w:val="24"/>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;

/** The file's own format record on a run, by another author. */
const FILE_FORMAT =
  '<w:rPrChange w:id="7" w:author="Author" w:date="2026-07-07T18:18:00Z">' +
  '<w:rPr><w:color w:val="000000"/></w:rPr></w:rPrChange>';

async function foldOf(
  body: string,
  edit: (doc: OrderedDocument) => void,
  label: string,
  extras?: PackageExtras
) {
  const { document: doc, base } = openOrderedWordFile(docx(body, extras), { file: 1 });
  edit(doc);
  await checkFold(doc, base, mediaStore(base), label);
}

describe('a decision after a fold, on a suggestion inside a run', () => {
  test('a suggested word', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
      },
      'a word inside a run'
    );
  });

  test('a suggested deletion', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 6), to: at(doc, 0, 10) });
      },
      'a deletion inside a run'
    );
  });

  test("a suggested word in a run with the file's format record", async () => {
    await foldOf(
      p(run('Lorem ipsum dolor.', FILE_FORMAT)),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 7), text: 'X ' });
      },
      'a word inside a formatted run'
    );
  });
});

describe('a decision after a fold, on a new list', () => {
  /** Ana's suggested list on the first paragraph. */
  const newList = (doc: OrderedDocument) => {
    apply(doc, 'ana:1', {
      kind: 'new-list',
      ...EDIT,
      list: {
        abstract: {
          e: 'w:abstractNum',
          c: [
            {
              e: 'w:lvl',
              a: { 'w:ilvl': '0' },
              c: [
                { e: 'w:numFmt', a: { 'w:val': 'lowerLetter' } },
                { e: 'w:lvlText', a: { 'w:val': '(%1)' } },
              ],
            },
          ],
        },
      },
      paragraphs: [{ mark: markOf(doc, 0), level: 0 }],
    });
  };
  const ITEMS = p(run('First item.')) + p(run('Second item.'));

  test('in a file with no lists: the numbering part goes with it', async () => {
    await foldOf(ITEMS, newList, 'a new list, no numbering part');
  });

  test("in a file with lists of its own: the file's stay", async () => {
    await foldOf(ITEMS, newList, 'a new list, a numbering part', MATRIX_EXTRAS);
  });
});
