// An Office 2016 chart (`cx:chart`: a box and whisker, a waterfall, a funnel) on copy and paste.
// Word copies it with its own parts, as it copies a chart: the chartEx part, its style and colors,
// and its embedded workbook. So does the editor: the chartEx part is XML read by the bounded
// reader with its root checked, the workbook bytes are a zip that nothing opens, and the paste
// writes each under a new name with new relationships. The editor paints the drawing Word writes
// in the fallback.

import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { relationshipsOf, resolveContentTypeOf } from '../package/package-edit.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import {
  bodyOf,
  buildPackage,
  CT,
  fullBodyCoverage,
  loadPackage,
  openStore,
  paragraphIdsUnder,
  R,
  REL,
  W,
} from './clipboard-fragment-fixtures.ts';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';
const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const MS = 'http://schemas.microsoft.com/office/2011/relationships';
const CHART_EX = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const WORKBOOK = zipSync({ '[Content_Types].xml': strToU8(`<Types xmlns="${CT}"/>`) });

const CHART_EX_RUN =
  `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:cx1="${CX1}" Requires="cx1">` +
  `<w:drawing><wp:inline xmlns:wp="${WP}"><wp:extent cx="914400" cy="914400"/>` +
  `<wp:docPr id="7" name="Chart 7"/><a:graphic xmlns:a="${A}"><a:graphicData uri="${CX}">` +
  `<cx:chart xmlns:cx="${CX}" r:id="rId5"/></a:graphicData></a:graphic></wp:inline>` +
  '</w:drawing></mc:Choice><mc:Fallback><w:t>chart</w:t></mc:Fallback></mc:AlternateContent></w:r>';

const SOURCE = loadPackage(
  zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Default Extension="xlsx" ContentType="${XLSX}"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/charts/chartEx1.xml" ContentType="application/vnd.ms-office.chartex+xml"/>' +
        '<Override PartName="/word/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${CHART_EX}" Target="charts/chartEx1.xml"/></Relationships>`
    ),
    'word/charts/_rels/chartEx1.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${MS}/chartStyle" Target="style1.xml"/>` +
        `<Relationship Id="rId2" Type="${R}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>`
    ),
    'word/charts/chartEx1.xml': strToU8(
      `<cx:chartSpace xmlns:cx="${CX}" xmlns:r="${R}"><cx:chartData><cx:externalData r:id="rId2"/>` +
        '<cx:data id="0"><cx:numDim type="val"><cx:f>Sheet1!$A$2:$A$4</cx:f></cx:numDim></cx:data>' +
        '</cx:chartData><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="boxWhisker">' +
        '<cx:tx><cx:txData><cx:v>Fees</cx:v></cx:txData></cx:tx></cx:series></cx:plotAreaRegion>' +
        '</cx:plotArea></cx:chart></cx:chartSpace>'
    ),
    'word/charts/style1.xml': strToU8(`<cs:chartStyle xmlns:cs="${CS}" id="410"/>`),
    'word/embeddings/Microsoft_Excel_Worksheet.xlsx': WORKBOOK,
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Chart</w:t></w:r>` +
        `${CHART_EX_RUN}</w:p></w:body></w:document>`
    ),
  })
);

test('a copied Office 2016 chart travels with its parts; the paste names new ones', () => {
  const extracted = extractFragmentPackage(SOURCE, fullBodyCoverage(SOURCE));
  if (!extracted.ok) throw new Error(extracted.reason);
  // A target whose own rId5 is a link: a pasted rId5 would name it
  const target = buildPackage('<w:p/>', {
    'word/_rels/document.xml.rels':
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${R}/hyperlink" ` +
      'Target="https://example.com/" TargetMode="External"/></Relationships>',
  });
  const store = openStore(target);
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    { paragraphId: host, offset: 0, fragmentBytes: extracted.bytes, lastMarkCovered: true }
  );
  expect(result.ok).toBe(true);
  const pkg = store.currentPackage();
  const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
  expect(xml).toContain('Requires="cx1"');
  const id = /<cx:chart [^>]*r:id="([^"]+)"/.exec(xml)?.[1];
  expect(id).toBeDefined();
  expect(id).not.toBe('rId5');
  const record = relationshipsOf(pkg, '/word/document.xml').find((r) => r.id === id);
  expect(record?.type).toBe(CHART_EX);
  const chart = resolveInternalTarget('/word/document.xml', record!.rawTarget);
  if (!chart.ok) throw new Error('no chart part');
  expect(chart.partName).toMatch(/^\/word\/charts\/chartEx\d+\.xml$/);
  expect(resolveContentTypeOf(pkg, chart.partName)).toBe('application/vnd.ms-office.chartex+xml');
  expect(serializeOoxmlPart(pkg.parts.get(chart.partName)!)).toContain('<cx:v>Fees</cx:v>');
  const own = relationshipsOf(pkg, chart.partName);
  expect(own.map((r) => r.type).sort()).toEqual([`${MS}/chartStyle`, `${R}/package`]);
  const workbook = own.find((r) => r.type === `${R}/package`)!;
  const bytes = resolveInternalTarget(chart.partName, workbook.rawTarget);
  expect(bytes.ok ? pkg.partBytes.get(bytes.partName) : undefined).toEqual(WORKBOOK);
});
