// A Word file with tracked changes to style definitions, as Word writes them in
// the styles part: a paragraph style's `w:rPrChange` and `w:pPrChange`, a table style's
// `w:tblPrChange` and a conditional format's `w:tcPrChange`, and the document defaults'
// `w:rPrChange`. The default paragraph style's record has the `w:id`, author and date of the
// body's insertion, as in the s2 contract, where Word numbered the styles' records from 0 too.

import { docx, p, t, W } from './word-fixtures.ts';

const STYLE_DATE = '2026-10-01T00:00:00Z';

const record = (name: string, id: number, author: string, old: string) =>
  `<w:${name} w:id="${id}" w:author="${author}" w:date="${STYLE_DATE}">${old}</w:${name}>`;

/** The styles part: five records outside the stories. */
const STYLE_REVISIONS =
  `<w:styles xmlns:w="${W}">` +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/>' +
  record('rPrChange', 10, 'Bo', '<w:rPr><w:sz w:val="24"/></w:rPr>') +
  '</w:rPr></w:rPrDefault></w:docDefaults>' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:rPr><w:b/>' +
  record('rPrChange', 1, 'Ana', '<w:rPr><w:i/></w:rPr>') +
  '</w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="300"/>' +
  record('pPrChange', 7, 'Ana', '<w:pPr><w:spacing w:after="100"/></w:pPr>') +
  '</w:pPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
  '<w:tblPr><w:tblInd w:w="200" w:type="dxa"/>' +
  record('tblPrChange', 8, 'Bo', '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/></w:tblPr>') +
  '</w:tblPr><w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:fill="DDDDDD"/>' +
  record('tcPrChange', 9, 'Bo', '<w:tcPr/>') +
  '</w:tcPr></w:tblStylePr></w:style></w:styles>';

/** Body cards: Ana's insertion (`w:id` 1, as the default style's record) and her deletion. */
const STYLED_BODY =
  p(
    t('Clause 0 ') +
      `<w:ins w:id="1" w:author="Ana" w:date="${STYLE_DATE}">${t('new ')}</w:ins>` +
      t('reads.')
  ) +
  p(
    t('Quoted ') +
      `<w:del w:id="2" w:author="Ana" w:date="${STYLE_DATE}"><w:r><w:delText xml:space="preserve">old </w:delText></w:r></w:del>` +
      t('text.'),
    '<w:pStyle w:val="Quote"/>'
  );

export const STYLED = docx(STYLED_BODY, { parts: { 'word/styles.xml': STYLE_REVISIONS } });

/** The body's text after an accept and after a reject of both body cards. */
export const ACCEPTED_TEXTS = ['Clause 0 new reads.', 'Quoted text.'];
export const REJECTED_TEXTS = ['Clause 0 reads.', 'Quoted old text.'];
