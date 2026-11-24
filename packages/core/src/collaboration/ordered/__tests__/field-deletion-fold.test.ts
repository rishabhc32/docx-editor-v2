// A suggested deletion of a field, of each shape, alone and with the words before it: the Word
// file holds it, so a save and a reopen, and one and two folds, give back its card, its person,
// its date, its take-back and both decisions (design D14). A simple field with no result has no
// place for a mark: `w:del` and `w:ins` cannot hold a `w:fldSimple`, so the mark goes around its
// result runs, and it has none. A suggestion that strikes or inserts one is refused before it
// applies (`untracked-only`, design D9); a direct edit makes it.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, DATE, DIRECT, EDIT, refusal } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

const BEGIN =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>AUTHOR</w:instrText></w:r>';
const END = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const EMPTY = '<w:fldSimple w:instr="AUTHOR"/>';

/** "Prepared by " and a field: its atom is at offset 12. */
const prepared = (field: string) => docx(p(t('Prepared by ') + field + t('.')));

/** Bob's deletion of the field, or of "by " and the field. */
const deletion = (wide: boolean) => (doc: OrderedDocument) =>
  apply(
    doc,
    'bob:1',
    { kind: 'delete', ...EDIT, from: at(doc, 0, wide ? 9 : 12), to: at(doc, 0, 13) },
    'Bob'
  );

/** Bob's pending suggestions. */
const bobs = (doc: OrderedDocument) =>
  [...doc.suggestions.values()]
    .filter((s) => s.status === 'pending' && s.attribution.author === 'Bob')
    .map((s) => s.id);

describe('a suggested deletion of a field', () => {
  for (const [shape, field] of [
    [
      'a simple field with a result',
      '<w:fldSimple w:instr="AUTHOR">' + t('Bob') + '</w:fldSimple>',
    ],
    ['a complex field with no result', BEGIN + END],
    [
      'a complex field with an empty result',
      BEGIN + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + END,
    ],
  ] as const) {
    for (const wide of [false, true]) {
      const label = `${shape}${wide ? ', with the words before it' : ''}`;
      test(label, async () => {
        const { document: doc, base } = openOrderedWordFile(prepared(field), { file: 1 });
        const blobs = mediaStore(base);
        deletion(wide)(doc);
        expect(orderedCards(doc).map((card) => card.suggestions)).toEqual([bobs(doc)]);

        // A save and a reopen keep the card, under the name and the date Word saved
        const saved = openOrderedWordFile(orderedPackageBytes(doc, base), { file: 2 }).document;
        expect(orderedCards(saved).length).toBe(1);
        expect(bobs(saved).map((id) => saved.suggestions.get(id)?.attribution)).toEqual([
          { author: 'Bob', person: null, date: DATE },
        ]);

        await checkFold(doc, base, blobs, label);
        await checkFold(doc, base, blobs, `${label}, taken back`, (copy) =>
          apply(copy, 'bob:2', { kind: 'retract', suggestions: bobs(copy) }, 'Bob')
        );
      });
    }
  }
});

describe('a suggestion on a simple field with no result', () => {
  for (const wide of [false, true]) {
    test(`a deletion${wide ? ' with the words before it' : ''} is refused`, () => {
      const { document: doc } = openOrderedWordFile(prepared(EMPTY), { file: 1 });
      const from = at(doc, 0, wide ? 9 : 12);
      const to = at(doc, 0, 13);
      expect(refusal(doc, 'bob:1', { kind: 'delete', ...EDIT, from, to }, 'Bob')).toBe(
        'untracked-only'
      );
    });
  }

  test('a paste that holds one is refused', () => {
    const { document: doc } = openOrderedWordFile(prepared(''), { file: 1 });
    const run = { e: 'w:r', c: [{ e: 'w:t', c: ['By '] }] };
    const field = { e: 'w:fldSimple', a: { 'w:instr': 'AUTHOR' } };
    const blocks = [{ e: 'w:p', c: [run, field] }];
    const place = at(doc, 0, 0);
    expect(refusal(doc, 'bob:1', { kind: 'insert-fragment', ...EDIT, place, blocks }, 'Bob')).toBe(
      'untracked-only'
    );
  });

  test('a direct deletion removes it, and a fold keeps that', async () => {
    const { document: doc, base } = openOrderedWordFile(prepared(EMPTY), { file: 1 });
    apply(
      doc,
      'olive:1',
      { kind: 'delete', ...DIRECT, from: at(doc, 0, 12), to: at(doc, 0, 13) },
      'Olive'
    );
    expect(orderedCards(doc)).toEqual([]);
    await checkFold(doc, base, mediaStore(base), 'a direct deletion');
  });
});
