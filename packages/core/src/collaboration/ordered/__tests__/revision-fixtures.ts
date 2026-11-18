// Story bodies with Word revisions already in them, one kind each, as Word writes them.

import { p, t } from './word-fixtures.ts';

const mark = (id: number, author = 'Cara', utc = false) =>
  `w:id="${id}" w:author="${author}" w:date="2026-01-01T00:00:00Z"` +
  (utc ? ' w16du:dateUtc="2026-01-01T00:00:00Z"' : '');
const W16DU = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
const cell = (content: string, tcPr = '<w:tcW w:w="3000" w:type="dxa"/>') =>
  `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;
const table = (rows: string) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  `<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${rows}</w:tbl>`;

/** Bodies by name. Each holds pending revisions; a test decides them. */
export const REVISION_BODIES: Readonly<Record<string, string>> = {
  insertion: p(t('Alpha ') + `<w:ins ${mark(901)}>${t('earlier')}</w:ins>` + t(' Omega.')),
  deletion: p(
    t('Alpha ') +
      `<w:del ${mark(902)}><w:r><w:delText>former</w:delText></w:r></w:del>` +
      t(' Omega.')
  ),
  format: p(
    t('Alpha ') +
      `<w:r><w:rPr><w:b/><w:rPrChange ${mark(903)}><w:rPr/></w:rPrChange></w:rPr><w:t>heavy</w:t></w:r>` +
      t(' Omega.')
  ),
  'struck paragraph mark': p(t('First'), `<w:rPr><w:del ${mark(904)}/></w:rPr>`) + p(t('Second')),
  'inserted paragraph':
    p(`<w:ins ${mark(905)}>${t('New')}</w:ins>`, `<w:rPr><w:ins ${mark(905)}/></w:rPr>`) +
    p(t('Old')),
  'paragraph properties': p(
    t('Centred'),
    `<w:jc w:val="center"/><w:pPrChange ${mark(906)}><w:pPr/></w:pPrChange>`
  ),
  'inserted row':
    table(
      `<w:tr>${cell(p(t('Kept')))}</w:tr>` +
        `<w:tr><w:trPr><w:ins ${mark(907)}/></w:trPr>${cell(p(`<w:ins ${mark(907)}>${t('Added')}</w:ins>`))}</w:tr>`
    ) + p(t('After')),
  'cell properties':
    table(
      `<w:tr>${cell(p(t('Wide')), `<w:tcW w:w="3000" w:type="dxa"/><w:tcPrChange ${mark(914)}><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr></w:tcPrChange>`)}</w:tr>`
    ) + p(t('After')),
  // A rejected row record resets the row's cells with no record of their own (the engine's
  // `restoreUnrecordedCells`); the cell with its own record keeps it
  'row properties':
    table(
      `<w:tr><w:trPr><w:jc w:val="center"/><w:trPrChange ${mark(920)}><w:trPr/></w:trPrChange></w:trPr>${cell(p(t('Row')))}</w:tr>`
    ) + p(t('After')),
  // Word for Mac shows, accepts and rejects a row's `w:tblPrExChange` (tblprex2.docx, 2026-10-29)
  'row exceptions':
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(p(t('Single')))}${cell(p(t('Single')))}</w:tr>` +
    `<w:tr><w:tblPrEx><w:tblBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>` +
    `<w:tblPrExChange ${mark(919)}><w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPrEx></w:tblPrExChange></w:tblPrEx>` +
    cell(
      p(t('Shaded')),
      `<w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tcPrChange ${mark(921)}><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr></w:tcPrChange>`
    ) +
    `${cell(p(t('Double')))}</w:tr></w:tbl>` +
    p(t('After')),
  // Word for Mac counts a first-edition `w:numberingChange` as one revision; both decisions
  // remove it and keep the list (numchange.docx, 2026-10-29)
  'numbering record': p(
    t('Item'),
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:numberingChange ${mark(922)} w:original="%1)"/></w:numPr>`
  ),
  // A tracked list membership (`w:numPr/w:ins`): both decisions remove the record and keep the
  // list, as Word retains `w:numPr` either way (the engine's `validNumberingRevision`)
  'numbering insert': p(
    t('Listed'),
    `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/><w:ins ${mark(924)}/></w:numPr>`
  ),
  'numbering records together': p(
    t('Both'),
    '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>' +
      `<w:numberingChange ${mark(925)} w:original="%1."/><w:ins ${mark(926, 'Ben')}/></w:numPr>`
  ),
  'deletion inside an insertion': p(
    t('Alpha ') +
      `<w:ins ${mark(908)}><w:del ${mark(909, 'Ben')}><w:r><w:delText>gone</w:delText></w:r></w:del>${t('kept')}</w:ins>` +
      t(' Omega.')
  ),
  move: p(
    `<w:moveFromRangeStart ${mark(910)} w:name="mv1"/><w:moveFrom ${mark(911)}>${t('Moved ')}</w:moveFrom><w:moveFromRangeEnd w:id="910"/>` +
      t('stays. ') +
      `<w:moveToRangeStart ${mark(912)} w:name="mv1"/><w:moveTo ${mark(913)}>${t('Moved ')}</w:moveTo><w:moveToRangeEnd w:id="912"/>`
  ),
  'revision with a UTC date': `<w:p ${W16DU}>${t('Alpha ')}<w:ins ${mark(915, 'Cara', true)}>${t('stamped')}</w:ins></w:p>`,
  'deleted field': p(
    t('Page ') +
      `<w:del ${mark(916)}><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>` +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>1</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>'
  ),
  'two authors': p(
    `<w:ins ${mark(917, 'Ana')}>${t('one ')}</w:ins>` +
      t('two ') +
      `<w:del ${mark(918, 'Ben')}><w:r><w:delText>three</w:delText></w:r></w:del>`
  ),
};
