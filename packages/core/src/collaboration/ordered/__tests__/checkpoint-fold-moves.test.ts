// A move range that holds another move's range, after a fold (design D14). A range end names its
// range start by `w:id`, so a load gives each end to its own move, and a decision of that move
// takes both its markers, live and after a fold. Before, a load gave an end the move of the last
// wrapper before it: the inner move's, so an accept of the outer move left its end (the fold
// shapes session, seed 78).

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, EDIT, markOf } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

describe('a move range inside another, after a fold', () => {
  test('each range end goes with its own move', async () => {
    // Ana's move to the end of the second paragraph, kept by a fold; Ben's move to that
    // paragraph's end goes inside her range
    const opened = openOrderedWordFile(docx(p(t('Cloud and serverless.')) + p(t('Read online.'))), {
      file: 1,
    });
    const blobs = mediaStore(opened.base);
    const live = opened.document;
    const first = { from: at(live, 0, 0), to: at(live, 0, 6), place: at(live, 1, 12) };
    apply(live, 'ana:1', { kind: 'move', ...EDIT, ...first });
    const { document: doc, base } = await folded(live, opened.base, 2, blobs);
    const second = { from: at(doc, 0, 10), to: at(doc, 0, 21), place: markOf(doc, 1) };
    apply(doc, 'ben:1', { kind: 'move', ...EDIT, ...second }, 'Ben');
    await checkFold(doc, base, blobs, 'a move range inside another');

    // Ana's move accepted: none of her markers is left, Ben's range stays whole
    const suggestions = orderedCardSeen(doc, 'S:ana:1.1') ?? [];
    apply(
      doc,
      'olive:1',
      { kind: 'decide', cards: [{ suggestions }], status: 'accepted' },
      'Olive'
    );
    const xml = strFromU8(
      unzipSync(orderedPackageBytes(doc, base, { blobs }))['word/document.xml']!
    );
    expect(xml.match(/<w:moveToRange(Start|End) /g)).toEqual([
      '<w:moveToRangeStart ',
      '<w:moveToRangeEnd ',
    ]);
  });
});

const STAMP = 'w:author="Ana" w:date="2026-11-25T09:00:00Z"';

/** Ana's move of "Cloud " to the second paragraph's end, as a Word file without range markers. */
const UNRANGED = docx(
  p(`<w:moveFrom w:id="1" ${STAMP}>${t('Cloud ')}</w:moveFrom>` + t('and serverless.')) +
    p(t('Read online.') + `<w:moveTo w:id="2" ${STAMP}>${t('Cloud ')}</w:moveTo>`)
);

/** Ana's move of "Cloud " to the end of the second paragraph, made in the room. */
const moveCloud = (doc: OrderedDocument) =>
  apply(doc, 'ana:1', {
    kind: 'move',
    ...EDIT,
    from: at(doc, 0, 0),
    to: at(doc, 0, 6),
    place: at(doc, 1, 12),
  });

/** Ben's edits at the end of the second paragraph, where Ana's moved words end. */
const atTheEnd = {
  move: (doc: OrderedDocument) =>
    apply(
      doc,
      'ben:1',
      { kind: 'move', ...EDIT, from: at(doc, 0, 10), to: at(doc, 0, 21), place: markOf(doc, 1) },
      'Ben'
    ),
  typing: (doc: OrderedDocument) =>
    apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 18), text: 'X' }, 'Ben'),
};

/** The second paragraph's move markers, moved words and insertions, in order. */
async function secondParagraph(
  file: Uint8Array,
  made: boolean,
  folds: number,
  edit: keyof typeof atTheEnd
) {
  const opened = openOrderedWordFile(file, { file: 1 });
  const blobs = mediaStore(opened.base);
  let { document: doc, base } = opened;
  if (made) {
    moveCloud(doc);
  }

  for (let k = 0; k < folds; k += 1) {
    ({ document: doc, base } = await folded(doc, base, doc.file + 1, blobs));
  }

  atTheEnd[edit](doc);
  const xml = strFromU8(unzipSync(orderedPackageBytes(doc, base, { blobs }))['word/document.xml']!);
  const second = [...xml.matchAll(/<w:p[ >](?:(?!<\/w:p>).)*<\/w:p>/gs)][1]?.[0] ?? '';
  return [
    ...second.matchAll(/<w:(moveTo(?:Range(?:Start|End))?|ins)[ >]|<w:t(?: [^>]*)?>([^<]*)/g),
  ].map((m) => m[1] ?? m[2]);
}

// Seed 78: a move made live had no range marker items, so a later move or typing at the end of
// its words went after its range; after a fold, and in the editor alone, inside it. A file's move
// without range markers read the same way
describe("an edit at the end of another person's moved words", () => {
  const inside = {
    move: ['moveToRangeStart', 'moveTo', 'Cloud ', 'moveToRangeStart', 'moveTo', 'serverless.'],
    typing: ['moveToRangeStart', 'moveTo', 'Cloud ', 'ins', 'X', 'moveToRangeEnd'],
  };
  for (const edit of ['move', 'typing'] as const) {
    for (const [what, file, made] of [
      ['made in the room', docx(p(t('Cloud and serverless.')) + p(t('Read online.'))), true],
      ['in a file without range markers', UNRANGED, false],
    ] as const) {
      test(`a ${edit} goes inside a move's range ${what}, live and after a fold`, async () => {
        const runs = await Promise.all(
          [0, 1].map((folds) => secondParagraph(file, made, folds, edit))
        );
        expect(runs[0]).toEqual(runs[1]);
        expect(runs[0]?.slice(1, 1 + inside[edit].length)).toEqual(inside[edit]);
      });
    }
  }
});
