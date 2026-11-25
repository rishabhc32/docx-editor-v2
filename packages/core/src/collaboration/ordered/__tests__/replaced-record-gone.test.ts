// A record that a suggestion took over (policy 2) on a paragraph a direct deletion took shows
// nowhere: the file holds no change record of it, so the suggestion's card names no change it
// replaced there (`orderedReplacedSuggestions`), live and after a fold. A pending deletion of the
// paragraph still shows the record struck, and the card names it.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedReplacedSuggestions } from '../compose/replaced.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, DIRECT, EDIT, markOf, opens } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

const FILE = docx(
  p(t('Before.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell('Left.')}${cell('Right.')}</w:tr></w:tbl>` +
    p(t('After.'))
);

/** The suggestions each of Ana's second suggestion's records was taken from, by their ids. */
const replacedBy = (doc: OrderedDocument): string[] =>
  orderedReplacedSuggestions(doc, 'S:ana:2.1').map((r) => r.suggestion);

/**
 * Ana's alignment on the right cell's paragraph, then her indent on it and on "After.", which
 * takes the alignment's record there; then the right column deleted in `mode`.
 */
function session(mode: typeof DIRECT | typeof EDIT): OrderedDocument {
  const doc = openOrderedWordFile(FILE, { file: 1 }).document;
  const right = markOf(doc, 2);
  const align = { localName: 'jc', attributes: { val: 'right' } };
  apply(doc, 'ana:1', { kind: 'paragraph-property', ...EDIT, marks: [right], properties: [align] });
  const indent = { localName: 'ind', attributes: { left: '720' } };
  const marks = [right, markOf(doc, 3)];
  apply(doc, 'ana:2', { kind: 'paragraph-property', ...EDIT, marks, properties: [indent] });
  expect(replacedBy(doc)).toEqual(['S:ana:1.1']);
  const column = { kind: 'delete-column', ...mode, cell: addressOf(opens(doc, 'tc')[1]!) } as const;
  apply(doc, 'olive:1', column, 'Olive');
  return doc;
}

describe('a record taken over on a paragraph a deletion took', () => {
  test('a direct deletion leaves no change it replaced on the card, live and folded', async () => {
    const doc = session(DIRECT);
    const base = openOrderedWordFile(FILE, { file: 1 }).base;
    const blobs = mediaStore(base);
    expect(replacedBy(doc)).toEqual([]);
    expect(replacedBy((await folded(doc, base, 2, blobs)).document)).toEqual([]);
    await checkFold(doc, base, blobs, 'a direct deletion');
  });

  test('a pending deletion still shows the record struck, and the card names it', async () => {
    const doc = session(EDIT);
    const base = openOrderedWordFile(FILE, { file: 1 }).base;
    expect(replacedBy(doc)).toEqual(['S:ana:1.1']);
    await checkFold(doc, base, mediaStore(base), 'a pending deletion');
  });
});
