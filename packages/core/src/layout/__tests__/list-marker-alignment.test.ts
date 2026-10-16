// A centred or right-aligned list paragraph moves its marker with its first line, as Word does:
// the marker, its suffix and the text are one line. The text keeps its place and the marker
// keeps its gap to the text.

import { expect, test } from 'bun:test';
import { readOoxmlPart } from '../../store/index.ts';
import { layoutSemanticDocument, createFixedMeasurer } from '../semantic-layout.ts';
import { buildNumberingIndex } from '../numbering-index.ts';
import { paragraphFragmentsOf } from '../semantic-records.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function part(xml: string) {
  const parsed = readOoxmlPart(xml, { name: '/word/document.xml', contentType: 'app/xml' });
  if (!parsed.ok) throw Error(parsed.reason);
  return parsed.part;
}
const measurer = createFixedMeasurer(6, 14);
// The marker hangs 18pt before text that starts at 36pt, on the leading side.
const numberingIndex = buildNumberingIndex(
  part(
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:suff w:val="tab"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
  ).root
);

function layout(
  jc: string,
  options: { rtl?: boolean; table?: boolean; width?: number; pageBreak?: boolean } = {}
) {
  const width = options.width ?? 200;
  const paragraph = `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${options.rtl ? '<w:bidi/>' : ''}<w:jc w:val="${jc}"/></w:pPr>${options.pageBreak ? '<w:r><w:br w:type="page"/></w:r>' : ''}<w:r><w:t>Term</w:t></w:r></w:p>`;
  const body = options.table
    ? `<w:tbl><w:tblPr><w:tblW w:w="${width * 20}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="${width * 20}"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="${width * 20}" w:type="dxa"/></w:tcPr>${paragraph}</w:tc></w:tr></w:tbl>`
    : paragraph;
  const doc = layoutSemanticDocument(
    part(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`),
    0,
    {
      measurer,
      numberingIndex,
      geometry: { width, height: 300, margin: { top: 0, bottom: 0, left: 0, right: 0 } },
    }
  );
  const fragment = doc.pages
    .flatMap((page) => [...paragraphFragmentsOf(page, true)])
    .find((candidate) => candidate.marker)!;
  const line = fragment.lines[0]!;
  return { marker: fragment.marker!.box, text: line.spans[0]!.box, available: line.box.width };
}

for (const [jc, share] of [
  ['center', 1 / 2],
  ['right', 1],
] as const)
  test(`a list paragraph with w:jc="${jc}" moves its marker as far as its text`, () => {
    const left = layout('left');
    const aligned = layout(jc);
    const shift = (aligned.available - aligned.text.width) * share;
    expect(shift).toBeGreaterThan(0);
    // The text sits where alignment put it before; the marker follows it.
    expect(aligned.text.x).toBeCloseTo(left.text.x + shift, 5);
    expect(aligned.marker.x).toBeCloseTo(left.marker.x + shift, 5);
    expect(aligned.text.x - aligned.marker.x).toBeCloseTo(left.text.x - left.marker.x, 5);
  });

for (const jc of ['left', 'both'])
  test(`a list paragraph with w:jc="${jc}" keeps its marker in its slot`, () => {
    const { marker, text } = layout(jc);
    expect(marker.x).toBe(18);
    expect(text.x).toBe(36);
  });

// In a bidi paragraph `left` is the leading (right) edge and `right` the trailing (left) one.
for (const table of [false, true])
  for (const [jc, share] of [
    ['center', 1 / 2],
    ['right', 1],
  ] as const)
    test(`a right-to-left list paragraph with w:jc="${jc}" moves its marker as far as its text (table=${table})`, () => {
      const leading = layout('left', { rtl: true, table });
      const aligned = layout(jc, { rtl: true, table });
      const shift = (aligned.available - aligned.text.width) * share;
      expect(shift).toBeGreaterThan(0);
      expect(aligned.text.x).toBeCloseTo(leading.text.x - shift, 5);
      expect(aligned.marker.x).toBeCloseTo(leading.marker.x - shift, 5);
      expect(leading.marker.x).toBeGreaterThan(leading.text.x + leading.text.width);
    });

test('a centred list paragraph in a table cell moves its marker as far as its text', () => {
  const left = layout('left', { table: true });
  const centred = layout('center', { table: true });
  const shift = (centred.available - centred.text.width) / 2;
  expect(shift).toBeGreaterThan(0);
  expect(centred.text.x).toBeCloseTo(left.text.x + shift, 5);
  expect(centred.marker.x).toBeCloseTo(left.marker.x + shift, 5);
});

test('a centred list paragraph opened by a page break moves the marker on its text line', () => {
  const left = layout('left', { pageBreak: true });
  const centred = layout('center', { pageBreak: true });
  const shift = (centred.available - centred.text.width) / 2;
  expect(shift).toBeGreaterThan(0);
  expect(centred.text.x).toBeCloseTo(left.text.x + shift, 5);
  expect(centred.marker.x).toBeCloseTo(left.marker.x + shift, 5);
});

for (const options of [{}, { rtl: true }, { table: true }])
  test(`a first line wider than its measure keeps its marker in its slot ${JSON.stringify(options)}`, () => {
    // Indented past a 20pt region, the measure clamps to one point and every glyph overflows.
    const leading = layout('left', { ...options, width: 20 });
    const centred = layout('center', { ...options, width: 20 });
    expect(centred.text.width).toBeGreaterThan(centred.available);
    expect(centred.marker.x).toBe(leading.marker.x);
    expect(centred.text.x).toBe(leading.text.x);
  });
