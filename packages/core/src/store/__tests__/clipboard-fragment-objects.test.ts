// A chart and a diagram (SmartArt) copy and paste with their own parts, as Word copies them:
// the chart part with its embedded workbook and colors, the diagram's data, layout, style,
// colors and drawing parts. The paste writes each under a new name of Word's form, with new
// relationships, and an undo takes them all back.

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { writeOoxmlPackage, type OoxmlPackage } from '../package/ooxml-package.ts';
import {
  relationshipsOf,
  resolveContentTypeOf,
  validatePackageInvariants,
} from '../package/package-edit.ts';
import { resolveInternalTarget } from '../package/opc-names.ts';
import { serializeOoxmlPart, type OoxmlNode } from '../package/ooxml-tree.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import {
  bodyOf,
  fullBodyCoverage,
  loadPackage,
  openStore,
  paragraphIdsUnder,
} from './clipboard-fragment-fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const DML = 'application/vnd.openxmlformats-officedocument.drawingml';
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const WORKBOOK = zipSync({ '[Content_Types].xml': strToU8(`<Types xmlns="${CT}"/>`) });

const inline = (id: number, uri: string, data: string): string =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>` +
  `<wp:docPr id="${id}" name="Object ${id}"/><a:graphic xmlns:a="${A}">` +
  `<a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const CHART_RUN = inline(31, C, `<c:chart xmlns:c="${C}" xmlns:r="${R}" r:id="rIdChart"/>`);
const DIAGRAM_RUN = inline(
  32,
  DGM,
  `<dgm:relIds xmlns:dgm="${DGM}" xmlns:r="${R}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>`
);

interface SourceOptions {
  /** The chart's workbook is a file outside the package (a linked chart). */
  readonly linked?: boolean;
  /** The chart part's root is not `c:chartSpace`. */
  readonly badRoot?: boolean;
  /** The diagram's drawing part's root is not `dsp:drawing`. */
  readonly badDrawing?: boolean;
}

function source(options: SourceOptions = {}): OoxmlPackage {
  const chartRels = options.linked
    ? `<Relationship Id="rIdXlsx" Type="${R}/oleObject" Target="file:///C:/Book1.xlsx" TargetMode="External"/>`
    : `<Relationship Id="rIdXlsx" Type="${R}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>` +
      `<Relationship Id="rIdColors" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>`;
  const chartRoot = options.badRoot ? 'c:userShapes' : 'c:chartSpace';
  const drawingRoot = options.badDrawing ? 'dsp:spTree' : 'dsp:drawing';
  return loadPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Default Extension="xlsx" ContentType="${XLSX}"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
          `<Override PartName="/word/charts/chart1.xml" ContentType="${DML}.chart+xml"/>` +
          `<Override PartName="/word/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>` +
          `<Override PartName="/word/diagrams/data1.xml" ContentType="${DML}.diagramData+xml"/>` +
          `<Override PartName="/word/diagrams/layout1.xml" ContentType="${DML}.diagramLayout+xml"/>` +
          `<Override PartName="/word/diagrams/quickStyle1.xml" ContentType="${DML}.diagramStyle+xml"/>` +
          `<Override PartName="/word/diagrams/colors1.xml" ContentType="${DML}.diagramColors+xml"/>` +
          `<Override PartName="/word/diagrams/drawing1.xml" ContentType="application/vnd.ms-office.drawingml.diagramDrawing+xml"/>` +
          `</Types>`
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}">` +
          `<Relationship Id="rIdChart" Type="${R}/chart" Target="charts/chart1.xml"/>` +
          `<Relationship Id="rIdDm" Type="${R}/diagramData" Target="diagrams/data1.xml"/>` +
          `<Relationship Id="rIdLo" Type="${R}/diagramLayout" Target="diagrams/layout1.xml"/>` +
          `<Relationship Id="rIdQs" Type="${R}/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>` +
          `<Relationship Id="rIdCs" Type="${R}/diagramColors" Target="diagrams/colors1.xml"/>` +
          `<Relationship Id="rIdDsp" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>` +
          `</Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}"><w:body>` +
          `<w:p><w:r><w:t>Chart: </w:t></w:r>${CHART_RUN}</w:p>` +
          `<w:p><w:r><w:t>Diagram: </w:t></w:r>${DIAGRAM_RUN}</w:p>` +
          `</w:body></w:document>`
      ),
      'word/charts/chart1.xml': strToU8(
        `<${chartRoot} xmlns:c="${C}" xmlns:r="${R}"><c:chart><c:plotArea><c:layout/></c:plotArea></c:chart>` +
          `<c:externalData r:id="rIdXlsx"><c:autoUpdate val="0"/></c:externalData></${chartRoot}>`
      ),
      'word/charts/_rels/chart1.xml.rels': strToU8(
        `<Relationships xmlns="${REL}">${chartRels}</Relationships>`
      ),
      'word/charts/colors1.xml': strToU8(`<cs:colorStyle xmlns:cs="${CS}" meth="cycle" id="10"/>`),
      'word/embeddings/Microsoft_Excel_Worksheet.xlsx': WORKBOOK,
      'word/diagrams/data1.xml': strToU8(
        `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/>` +
          `<dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">` +
          `<dsp:dataModelExt xmlns:dsp="${DSP}" relId="rIdDsp" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>` +
          `</a:ext></dgm:extLst></dgm:dataModel>`
      ),
      'word/diagrams/layout1.xml': strToU8(
        `<dgm:layoutDef xmlns:dgm="${DGM}"><dgm:layoutNode/></dgm:layoutDef>`
      ),
      'word/diagrams/quickStyle1.xml': strToU8(`<dgm:styleDef xmlns:dgm="${DGM}"/>`),
      'word/diagrams/colors1.xml': strToU8(`<dgm:colorsDef xmlns:dgm="${DGM}"/>`),
      'word/diagrams/drawing1.xml': strToU8(
        `<${drawingRoot} xmlns:dsp="${DSP}"><dsp:spTree/></${drawingRoot}>`
      ),
    })
  );
}

/** A target that already holds a chart part named as the source's, so the paste renames. */
function target(): OoxmlPackage {
  return loadPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
          `<Override PartName="/word/charts/chart1.xml" ContentType="${DML}.chart+xml"/>` +
          `</Types>`
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId7" Type="${R}/chart" Target="charts/chart1.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Host</w:t></w:r></w:p></w:body></w:document>`
      ),
      'word/charts/chart1.xml': strToU8(`<c:chartSpace xmlns:c="${C}"><c:chart/></c:chartSpace>`),
    })
  );
}

/** Every attribute value in the relationships namespace under `node`, by local name. */
function relationshipAttributes(
  node: OoxmlNode,
  out: Array<[string, string]> = []
): Array<[string, string]> {
  if (node.kind === 'textValue') return out;
  for (const attribute of node.attributes)
    if (attribute.namespaceUri === R) out.push([attribute.localName, attribute.value]);
  for (const child of node.children) relationshipAttributes(child, out);
  return out;
}

function pasteInto(pkg: OoxmlPackage, from: OoxmlPackage) {
  const extracted = extractFragmentPackage(from, fullBodyCoverage(from));
  if (!extracted.ok) throw new Error(`extract: ${extracted.reason}`);
  const store = openStore(pkg);
  const host = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get(pkg.mainDocumentPart)!)
  )[0]!;
  const pasted = store.applyFragmentPaste(
    { kind: 'body' },
    { paragraphId: host, offset: 4, fragmentBytes: extracted.bytes, lastMarkCovered: true }
  );
  return { extracted, store, pasted };
}

/** The part a story relationship id names in `pkg`. */
function partOf(pkg: OoxmlPackage, owner: string, id: string): string {
  const record = relationshipsOf(pkg, owner).find((entry) => entry.id === id);
  if (record === undefined) throw new Error(`no relationship ${id}`);
  const resolved = resolveInternalTarget(owner, record.rawTarget);
  if (!resolved.ok) throw new Error('unresolved');
  return resolved.partName;
}

describe('a chart and a diagram on copy and paste', () => {
  test('the copy carries every part of the chart and the diagram under its own name', () => {
    const { extracted } = pasteInto(target(), source());
    const names = [...extracted.package.partBytes.keys()].sort();
    expect(names).toContain('/word/charts/chart1.xml');
    expect(names).toContain('/word/charts/colors1.xml');
    expect(names).toContain('/word/embeddings/Microsoft_Excel_Worksheet.xlsx');
    for (const part of ['data1', 'layout1', 'quickStyle1', 'colors1', 'drawing1'])
      expect(names).toContain(`/word/diagrams/${part}.xml`);
    // The diagram's drawing travels by the story relationship its data part names
    const docRels = relationshipsOf(extracted.package, '/word/document.xml').map((r) => r.id);
    expect(docRels).toContain('rIdDsp');
    expect(extracted.mediaBytes).toBe(WORKBOOK.byteLength);
  });

  test('the paste writes each part under a new name, with relationships that resolve', () => {
    const { store, pasted } = pasteInto(target(), source());
    expect(pasted.ok).toBe(true);
    const pkg = store.currentPackage();
    const body = bodyOf(pkg.parts.get('/word/document.xml')!);
    const refs = new Map(relationshipAttributes(body));
    const chart = partOf(pkg, '/word/document.xml', refs.get('id')!);
    // The target's own chart1 stays; the pasted chart takes the next free number
    expect(chart).toBe('/word/charts/chart2.xml');
    expect(serializeOoxmlPart(pkg.parts.get('/word/charts/chart1.xml')!)).not.toContain(
      'externalData'
    );
    expect(resolveContentTypeOf(pkg, chart)).toBe(`${DML}.chart+xml`);
    const chartRels = relationshipsOf(pkg, chart);
    const workbook = chartRels.find((r) => r.type === `${R}/package`)!;
    const workbookPart = resolveInternalTarget(chart, workbook.rawTarget);
    expect(workbook.rawTarget).toBe('../embeddings/Microsoft_Excel_Worksheet1.xlsx');
    expect(
      workbookPart.ok &&
        (pkg.partBytes.has(workbookPart.partName) ||
          pkg.partBytes.has(workbookPart.partName.slice(1)))
    ).toBe(true);
    expect(resolveContentTypeOf(pkg, '/word/embeddings/Microsoft_Excel_Worksheet1.xlsx')).toBe(
      XLSX
    );
    // The chart's own reference names its new relationship
    expect(serializeOoxmlPart(pkg.parts.get(chart)!)).toContain(`r:id="${workbook.id}"`);
    // The diagram's four parts, and its drawing through the data part's own id
    const data = partOf(pkg, '/word/document.xml', refs.get('dm')!);
    expect(data).toBe('/word/diagrams/data1.xml');
    expect(partOf(pkg, '/word/document.xml', refs.get('lo')!)).toBe('/word/diagrams/layout1.xml');
    const dataXml = serializeOoxmlPart(pkg.parts.get(data)!);
    const drawingId = /relId="([^"]+)"/.exec(dataXml)![1]!;
    expect(partOf(pkg, '/word/document.xml', drawingId)).toBe('/word/diagrams/drawing1.xml');
    expect(validatePackageInvariants(pkg).ok).toBe(true);
    // Save and reopen keeps every part and relationship
    const reopened = loadPackage(writeOoxmlPackage(pkg));
    expect(partOf(reopened, '/word/document.xml', refs.get('id')!)).toBe(chart);
    expect(
      relationshipsOf(reopened, chart)
        .map((r) => r.type)
        .sort()
    ).toEqual(chartRels.map((r) => r.type).sort());
  });

  test('a second paste of the same copy takes further names', () => {
    const from = source();
    const { store } = pasteInto(target(), from);
    const extracted = extractFragmentPackage(from, fullBodyCoverage(from));
    if (!extracted.ok) throw new Error('extract');
    const host = paragraphIdsUnder(
      bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
    )[0]!;
    const again = store.applyFragmentPaste(
      { kind: 'body' },
      { paragraphId: host, offset: 0, fragmentBytes: extracted.bytes, lastMarkCovered: true }
    );
    expect(again.ok).toBe(true);
    const pkg = store.currentPackage();
    expect(pkg.parts.has('/word/charts/chart3.xml')).toBe(true);
    // Word gives the second paste its own five diagram parts and its own workbook, and points
    // the second data part at its own drawing (paste evidence 4, `~$paste-out-e4-*-twice.docx`)
    for (const name of ['data2', 'layout2', 'quickStyle2', 'colors2', 'drawing2'])
      expect(pkg.parts.has(`/word/diagrams/${name}.xml`)).toBe(true);
    const dataXml = serializeOoxmlPart(pkg.parts.get('/word/diagrams/data2.xml')!);
    expect(partOf(pkg, '/word/document.xml', /relId="([^"]+)"/.exec(dataXml)![1]!)).toBe(
      '/word/diagrams/drawing2.xml'
    );
    const workbooks = ['/word/charts/chart2.xml', '/word/charts/chart3.xml'].map((chart) => {
      const record = relationshipsOf(pkg, chart).find((r) => r.type === `${R}/package`)!;
      return record.rawTarget;
    });
    expect(new Set(workbooks).size).toBe(2);
    // A pasted object whose drawing id is taken gets a new one; every id stays unique
    const ids = [
      ...serializeOoxmlPart(pkg.parts.get('/word/document.xml')!).matchAll(
        /<wp:docPr [^>]*id="(\d+)"/g
      ),
    ].map((match) => match[1]);
    expect(ids.length).toBeGreaterThanOrEqual(4);
    expect(new Set(ids).size).toBe(ids.length);
    expect(validatePackageInvariants(pkg).ok).toBe(true);
  });

  test('an undo takes every pasted part back', () => {
    const { store } = pasteInto(target(), source());
    const before = target();
    expect(store.undo()).not.toBeNull();
    const pkg = store.currentPackage();
    expect([...pkg.parts.keys()].filter((name) => !name.includes('_rels')).sort()).toEqual(
      [...before.parts.keys()].filter((name) => !name.includes('_rels')).sort()
    );
    expect([...pkg.partBytes.keys()].some((name) => name.includes('embeddings'))).toBe(false);
  });

  test('a chart linked to a workbook file, or a part of the wrong kind, stays out as before', () => {
    for (const options of [{ linked: true }, { badRoot: true }]) {
      const { store, pasted } = pasteInto(target(), source(options));
      expect(pasted.ok).toBe(true);
      const pkg = store.currentPackage();
      const xml = serializeOoxmlPart(pkg.parts.get('/word/document.xml')!);
      expect(xml).not.toContain('c:chart ');
      expect(xml).toContain('dgm:relIds');
      expect(pkg.parts.has('/word/charts/chart2.xml')).toBe(false);
      expect(validatePackageInvariants(pkg).ok).toBe(true);
    }
  });

  test('a diagram whose drawing part fails its check keeps no id of the source story', () => {
    // The drawing is a cache Word draws again; its id would name nothing, or another
    // relationship of the target story
    const { store, pasted } = pasteInto(target(), source({ badDrawing: true }));
    expect(pasted.ok).toBe(true);
    const pkg = store.currentPackage();
    expect(pkg.parts.has('/word/diagrams/drawing1.xml')).toBe(false);
    expect(serializeOoxmlPart(pkg.parts.get('/word/document.xml')!)).toContain('dgm:relIds');
    const dataXml = serializeOoxmlPart(pkg.parts.get('/word/diagrams/data1.xml')!);
    expect(dataXml).toContain('<dgm:whole/>');
    expect(dataXml).not.toContain('dataModelExt');
    expect(validatePackageInvariants(pkg).ok).toBe(true);
  });
});
