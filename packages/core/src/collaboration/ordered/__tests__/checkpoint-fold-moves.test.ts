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
