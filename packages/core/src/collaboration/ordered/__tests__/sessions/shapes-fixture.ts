// The document the fold shape sessions edit: a date picker, a combo box and a drop-down list,
// each showing its prompt beside a value it stores, a plain text control's prompt, words that
// hold a footnote and an endnote reference, a link, a section break, and a table whose cells
// hold footnote references, and the body's own section.

import { docx, MATRIX_EXTRAS, p, R, sdt, separators, t, W } from '../word-fixtures.ts';

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/><w:listItem w:displayText="Net 60" w:value="60"/>';
const CURRENCIES =
  '<w:listItem w:displayText="EUR" w:value="EUR"/><w:listItem w:displayText="USD" w:value="USD"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';

/** A control with the id and properties, showing its prompt. */
const prompt = (id: number, props: string, text: string): string =>
  sdt(
    `<w:id w:val="${id}"/><w:showingPlcHdr/>${props}`,
    t(text, '<w:rStyle w:val="PlaceholderText"/>')
  );

const footnote = (id: number) =>
  `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const cell = (content: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(content)}</w:tc>`;
const note = (id: number, text: string) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r>${t(text)}</w:p></w:footnote>`;

const DATE = `<w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`;
const SECTION =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" ' +
  'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';
const LINK = `<w:hyperlink r:id="rIdLink">${t('terms', '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/>' +
  `<w:gridCol w:w="3000"/></w:tblGrid><w:tr>${cell(t('Fee') + footnote(2) + t(' due'))}` +
  `${cell(t('Net 30 days'))}</w:tr><w:tr>${cell(t('Late fee'))}` +
  `${cell(t('Two per cent') + footnote(3))}</w:tr></w:tbl>`;

const BODY = [
  p(t('Due on ') + prompt(1, DATE, 'Click to enter a date.') + t(' at the latest.')),
  p(
    t('Paid in ') +
      prompt(2, `<w:comboBox w:lastValue="EUR">${CURRENCIES}</w:comboBox>`, 'Choose an item.') +
      t(' only.')
  ),
  p(
    t('Terms are ') +
      prompt(3, `<w:dropDownList w:lastValue="30">${ITEMS}</w:dropDownList>`, 'Choose an item.') +
      t(' from today.')
  ),
  p(t('Signed by ') + prompt(4, '<w:text/>', 'Click here to enter text.') + t(' for the buyer.')),
  p(
    t('See the fees') +
      footnote(1) +
      t(' and the annex') +
      '<w:r><w:endnoteReference w:id="1"/></w:r>' +
      t(' before you sign.')
  ),
  p(t('Read the ') + LINK + t(' online.')),
  p(t('End of part one.'), SECTION),
  p(t('Part two starts here.')),
  TABLE,
  p(t('Signed and dated.')),
  SECTION,
].join('');

/** The fold shape sessions' file. */
export const SHAPES_FILE = docx(BODY, {
  ...MATRIX_EXTRAS,
  parts: {
    ...MATRIX_EXTRAS.parts,
    'word/footnotes.xml':
      `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${separators('footnote')}` +
      `${note(1, ' Fees are due net of tax.')}${note(2, ' Paid monthly.')}` +
      `${note(3, ' Of the fee.')}</w:footnotes>`,
  },
});
