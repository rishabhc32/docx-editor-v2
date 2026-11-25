// A record that a suggestion took over (policy 2) on a paragraph a direct deletion took, or on a
// prompt that content took the place of, shows nowhere: the file holds no change record of it in
// the story, so the suggestion's card names no change it replaced there
// (`orderedReplacedSuggestions`), live and after a fold. A pending deletion of the paragraph still
// shows the record struck, and the card names it; so does the prompt once the content goes.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedReplacedSuggestions } from '../compose/replaced.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, at, DIRECT, EDIT, markOf, opens } from './change-apply-fixture.ts';
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

  // Seed 43 of the fold shapes: Ben's underline over part of his italic on a date's prompt, then
  // his typing over the prompt
  test('a prompt content took the place of names no change, until the content goes', async () => {
    const date =
      '<w:sdt><w:sdtPr><w:id w:val="1"/><w:showingPlcHdr/><w:date><w:dateFormat w:val="M/d/yyyy"/>' +
      '</w:date></w:sdtPr><w:sdtContent>' +
      `${t('Click to enter a date.', '<w:rStyle w:val="PlaceholderText"/>')}</w:sdtContent></w:sdt>`;
    const file = docx(p(t('Due on ') + date + t(' at the latest.')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const blobs = mediaStore(base);
    const format = { kind: 'format', ...EDIT } as const;
    const italic = [{ localName: 'i' }];
    apply(
      doc,
      'ben:1',
      { ...format, from: at(doc, 0, 7), to: at(doc, 0, 28), properties: italic },
      'Ben'
    );
    const underline = [{ localName: 'u', attributes: { val: 'single' } }];
    apply(
      doc,
      'ben:2',
      { ...format, from: at(doc, 0, 6), to: at(doc, 0, 14), properties: underline },
      'Ben'
    );
    const byBen = (room: OrderedDocument) =>
      orderedReplacedSuggestions(room, 'S:ben:2.1').map((r) => r.suggestion);
    expect(byBen(doc)).toEqual(['S:ben:1.1']);
    apply(doc, 'ben:3', { kind: 'insert', ...EDIT, place: at(doc, 0, 26), text: '€' }, 'Ben');
    expect(byBen(doc)).toEqual([]);
    expect(byBen((await folded(doc, base, 2, blobs)).document)).toEqual([]);
    await checkFold(doc, base, blobs, 'a prompt content took the place of');
    apply(doc, 'ben:4', { kind: 'retract', suggestions: ['S:ben:3.1'] }, 'Ben');
    expect(byBen(doc)).toEqual(['S:ben:1.1']);
  });
});
