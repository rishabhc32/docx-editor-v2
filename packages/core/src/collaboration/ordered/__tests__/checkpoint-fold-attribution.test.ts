// A suggestion given another date writes that date on each of its marks, also the marks a load
// read from the file: the file, the checkpoint's list and the load agree, so a fold changes nothing
// the room uses (design D14), and the maker's take-back and each decision work as in the room.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, DATE, DIRECT, EDIT, refusal } from './change-apply-fixture.ts';
import { checkFold, checkpointBytes, folded, mediaStore } from './fold-check.ts';
import { anas, FOLD_SHAPES } from './fold-shapes.ts';
import { REVISION_BODIES } from './revision-fixtures.ts';
import { STYLED } from './style-revision-fixture.ts';
import { bodyOf, docx, MATRIX_ITEMS, p, t } from './word-fixtures.ts';

const NEW_DATE = '2026-11-20T08:00:00Z';

/** Ana gives her pending suggestions the new date. */
const redate = (doc: OrderedDocument) =>
  apply(doc, 'ana:4', {
    kind: 'attribution',
    ...EDIT,
    suggestions: anas(doc),
    author: 'Ana',
    date: NEW_DATE,
  });

/** Ana takes her pending suggestions back. */
const retract = (doc: OrderedDocument) =>
  apply(doc, 'ana:5', { kind: 'retract', suggestions: anas(doc) });

describe('a new date on a suggestion a fold loaded', () => {
  // Ana's suggestion on the Word file, folded, so its marks are the file's, then given a new date
  for (const [label, file, suggest] of FOLD_SHAPES) {
    test(label, async () => {
      const first = openOrderedWordFile(file, { file: 1 });
      suggest(first.document);
      const blobs = mediaStore(first.base);
      const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
      redate(doc);
      await checkFold(doc, base, blobs, label);
      await checkFold(doc, base, blobs, `${label}, taken back`, retract);
    });
  }
});

describe('a new date on a suggestion of the Word file', () => {
  // A mark Word saved with its UTC date beside the local one: a new date drops both
  const file = docx(
    p(
      t('Pay ') +
        `<w:ins w:id="7" w:author="Ana" w:date="${DATE}" ` +
        'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" ' +
        `w16du:dateUtc="${DATE}">${t('within thirty days ')}</w:ins>` +
        t('now.')
    )
  );

  test('the fold writes the new date, and the copy reads the same suggestion', async () => {
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(base);
    redate(doc);
    const written = strFromU8(
      unzipSync(await checkpointBytes(doc, base, blobs))['word/document.xml']!
    );
    expect(written.match(/<w:ins\b[^>]*>/g)).toEqual([
      `<w:ins w:author="Ana" w:date="${NEW_DATE}" w:id="7">`,
    ]);
    await checkFold(doc, base, blobs, 'a file insertion');

    // A save holds the new date too, and a reopen reads it
    const saved = openOrderedWordFile(orderedPackageBytes(doc, base), { file: 2 }).document;
    expect([...saved.suggestions.values()].map((s) => s.attribution)).toEqual([
      { author: 'Ana', person: null, date: NEW_DATE },
    ]);
  });

  // The file's insertion is nobody's in the room: Ana's take-back is refused before and after
  test('a take-back after the fold gives what it gives in the room', async () => {
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(base);
    redate(doc);
    const once = await folded(doc, base, 2, blobs);
    const twice = await folded(once.document, once.base, 3, blobs);
    const takeBack = (copy: OrderedDocument) =>
      refusal(copy, 'ana:5', { kind: 'retract', suggestions: anas(copy) });
    expect([takeBack(once.document), takeBack(twice.document)]).toEqual([
      takeBack(doc),
      takeBack(doc),
    ]);
    expect(takeBack(doc)).toBe('not-own-suggestion');
  });
});

describe('a new date on each kind of Word revision', () => {
  // Each pending suggestion of the file given the new date by its own author
  const redateAll = (doc: OrderedDocument) => {
    for (const s of [...doc.suggestions.values()].filter((x) => x.status === 'pending')) {
      const { author } = s.attribution;
      apply(
        doc,
        `${author.toLowerCase()}:4`,
        { kind: 'attribution', ...EDIT, suggestions: [s.id], author, date: NEW_DATE },
        author
      );
    }
  };

  // The marks a fold writes with another date
  const oldDates = (bytes: Uint8Array) =>
    Object.entries(unzipSync(bytes)).flatMap(([name, data]) =>
      [...strFromU8(data).matchAll(/<w:\w+ [^>]*w:date="([^"]*)"[^>]*>/g)]
        .filter(([, date]) => date !== NEW_DATE)
        .map(([mark]) => `${name} ${mark}`)
    );

  for (const [label, file] of [
    ...Object.entries(REVISION_BODIES).map(([name, body]) => [name, docx(body)] as const),
    ['style records', STYLED] as const,
  ]) {
    test(label, async () => {
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      const blobs = mediaStore(base);
      redateAll(doc);
      expect(oldDates(await checkpointBytes(doc, base, blobs))).toEqual([]);
      await checkFold(doc, base, blobs, label);
    });
  }

  // Ana's format on a prompt that Olive's typing filled: the glossary keeps the prompt with its
  // record. The record takes the new date, and comes back with it when the control empties
  test('a prompt record the glossary keeps', async () => {
    const item = MATRIX_ITEMS.find((candidate) => candidate.key === 'placeholder');
    const first = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
    apply(first.document, 'ana:1', {
      kind: 'format',
      ...EDIT,
      from: at(first.document, 1, 7),
      to: at(first.document, 1, 11),
      properties: [{ localName: 'i' }],
    });
    apply(
      first.document,
      'olive:2',
      { kind: 'insert', ...DIRECT, place: at(first.document, 1, 6), text: 'Q' },
      'Olive'
    );
    const blobs = mediaStore(first.base);
    const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
    redate(doc);
    expect(oldDates(await checkpointBytes(doc, base, blobs))).toEqual([]);
    await checkFold(doc, base, blobs, 'a prompt record', (copy) => {
      apply(
        copy,
        'olive:3',
        { kind: 'delete', ...DIRECT, from: at(copy, 1, 6), to: at(copy, 1, 7) },
        'Olive'
      );

      // The prompt shows again with Ana's format
      expect(orderedCards(copy).map((card) => card.suggestions)).toEqual([anas(copy)]);
    });
  });
});
