// The sources of the chart and diagram paste tests of group 9.10: a line with a chart (its
// chart style, colors and embedded workbook), a line with a SmartArt diagram (its data,
// layout, style, colors and drawing parts), and a line with an Office 2016 chart (its chartEx
// part, style, colors and workbook), with the parts and relationships Word 2016 writes.
// Word's own copy of these is "waiting for Word" (`paste-evidence-report.md`, section 4); the
// parts follow ECMA-376 Part 1 (§14.2.1, §14.2.4 to §14.2.6, §15.2.11) and [MS-ODRAWXML].

import { strToU8, zipSync } from 'fflate';
import { docx, MATRIX_EXTRAS, p, pictureRun, R, t } from '../word-fixtures.ts';
import { copiedBetween, type Copied } from './paste-harness.ts';

const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
const CS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const DML = 'application/vnd.openxmlformats-officedocument.drawingml';
const CX = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
const CX1 = 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex';

/**
 * A chart's embedded workbook: a zip the room keeps as it is and never opens. Its time is fixed,
 * so its bytes and their hash are the same in each run.
 */
export const WORKBOOK: Uint8Array = zipSync(
  {
    '[Content_Types].xml': strToU8(
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'
    ),
    'xl/workbook.xml': strToU8('<workbook/>'),
  },
  { mtime: Date.UTC(2026, 0, 1) }
);

const frame = (id: number, name: string, uri: string, data: string): string =>
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  `<wp:docPr id="${id}" name="${name}"/><wp:cNvGraphicFramePr/>` +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="${uri}">${data}</a:graphicData></a:graphic>` +
  '</wp:inline></w:drawing></w:r>';

const CHART_RUN = frame(7, 'Chart 1', C, `<c:chart xmlns:c="${C}" r:id="rIdChart"/>`);
// An Office 2016 chart (a box and whisker) as Word writes it: a choice that requires `cx1`, and
// a picture of the chart as the fallback ([MS-ODRAWXML] chartex)
const CHART_EX_RUN =
  `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:cx1="${CX1}" Requires="cx1">` +
  frame(9, 'Chart 2', CX, `<cx:chart xmlns:cx="${CX}" r:id="rIdChartEx"/>`).slice(5, -6) +
  `</mc:Choice><mc:Fallback>${pictureRun(10).slice(5, -6)}</mc:Fallback></mc:AlternateContent></w:r>`;

const CHART_EX =
  `<cx:chartSpace xmlns:cx="${CX}" xmlns:a="${A}" xmlns:r="${R}"><cx:chartData>` +
  '<cx:externalData r:id="rId3" cx:autoUpdate="0"/><cx:data id="0"><cx:numDim type="val">' +
  '<cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">100</cx:pt>' +
  '<cx:pt idx="1">250</cx:pt><cx:pt idx="2">175</cx:pt></cx:lvl></cx:numDim></cx:data>' +
  '</cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData>' +
  '<cx:v>Fees spread</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion>' +
  '<cx:series layoutId="boxWhisker" uniqueId="{44444444-0000-0000-0000-000000000000}">' +
  '<cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling/>' +
  '</cx:axis><cx:axis id="1"><cx:valScaling/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>';

const DIAGRAM_RUN = frame(
  8,
  'Diagram 1',
  DGM,
  `<dgm:relIds xmlns:dgm="${DGM}" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/>`
);

const pt = (k: number, value: string) => `<c:pt idx="${k}"><c:v>${value}</c:v></c:pt>`;
const axis = (kind: 'catAx' | 'valAx', id: number, cross: number, position: string) =>
  `<c:${kind}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling>` +
  `<c:delete val="0"/><c:axPos val="${position}"/><c:numFmt formatCode="General" sourceLinked="1"/>` +
  '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
  `<c:crossAx val="${cross}"/><c:crosses val="autoZero"/></c:${kind}>`;

/** A clustered column chart of one series, with its title, as Word 2016 writes it. */
const CHART =
  `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}" ` +
  'xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart">' +
  '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>' +
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice Requires="c14" ` +
  'xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart">' +
  '<c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback>' +
  '</mc:AlternateContent><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>' +
  '<a:r><a:rPr lang="en-US"/><a:t>Fees by quarter</a:t></a:r></a:p></c:rich></c:tx>' +
  '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>' +
  '<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$B$1</c:f>' +
  `<c:strCache><c:ptCount val="1"/>${pt(0, 'Fees')}</c:strCache></c:strRef></c:tx>` +
  '<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr>' +
  '<c:invertIfNegative val="0"/><c:cat><c:strRef><c:f>Sheet1!$A$2:$A$3</c:f><c:strCache>' +
  `<c:ptCount val="2"/>${pt(0, 'Q1')}${pt(1, 'Q2')}</c:strCache></c:strRef></c:cat>` +
  '<c:val><c:numRef><c:f>Sheet1!$B$2:$B$3</c:f><c:numCache><c:formatCode>General</c:formatCode>' +
  `<c:ptCount val="2"/>${pt(0, '100')}${pt(1, '250')}</c:numCache></c:numRef></c:val>` +
  '<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" ' +
  'xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">' +
  '<c16:uniqueId val="{00000000-0001-0000-0000-000000000000}"/></c:ext></c:extLst></c:ser>' +
  '<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="11"/><c:axId val="12"/>' +
  `</c:barChart>${axis('catAx', 11, 12, 'b')}${axis('valAx', 12, 11, 'l')}</c:plotArea>` +
  '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/>' +
  '<c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>' +
  '<a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>' +
  '<c:externalData r:id="rId3"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>';

const CHART_STYLE =
  `<cs:chartStyle xmlns:cs="${CS}" xmlns:a="${A}" id="201"><cs:axisTitle><cs:lnRef idx="0"/>` +
  '<cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1">' +
  '<a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></cs:fontRef>' +
  '<cs:defRPr sz="1000" kern="1200"/></cs:axisTitle></cs:chartStyle>';

const CHART_COLORS =
  `<cs:colorStyle xmlns:cs="${CS}" xmlns:a="${A}" meth="cycle" id="10">` +
  '<a:schemeClr val="accent1"/><cs:variation/><cs:variation><a:lumMod val="60000"/>' +
  '</cs:variation></cs:colorStyle>';

const point = (id: string, type: string, text: string | null) =>
  `<dgm:pt modelId="${id}"${type === 'node' ? '' : ` type="${type}"`}><dgm:prSet/><dgm:spPr/>` +
  '<dgm:t><a:bodyPr/><a:lstStyle/><a:p>' +
  (text === null
    ? '<a:endParaRPr lang="en-US"/>'
    : `<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r>`) +
  '</a:p></dgm:t></dgm:pt>';

/** A Basic Block List of two items, its drawing named by the story's relationship `rIdDsp`. */
const DIAGRAM_DATA =
  `<dgm:dataModel xmlns:dgm="${DGM}" xmlns:a="${A}"><dgm:ptLst>` +
  point('{11111111-0000-0000-0000-000000000000}', 'doc', null) +
  point('{22222222-0000-0000-0000-000000000000}', 'node', 'Draft') +
  point('{33333333-0000-0000-0000-000000000000}', 'node', 'Sign') +
  '</dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{44444444-0000-0000-0000-000000000000}" ' +
  'srcId="{11111111-0000-0000-0000-000000000000}" destId="{22222222-0000-0000-0000-000000000000}" ' +
  'srcOrd="0" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst>' +
  '<a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram">' +
  `<dsp:dataModelExt xmlns:dsp="${DSP}" relId="rIdDsp" ` +
  'minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst>' +
  '</dgm:dataModel>';

const DIAGRAM_LAYOUT =
  `<dgm:layoutDef xmlns:dgm="${DGM}" xmlns:a="${A}" xmlns:r="${R}" ` +
  'uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/default"><dgm:title val=""/>' +
  '<dgm:desc val=""/><dgm:catLst><dgm:cat type="list" pri="400"/></dgm:catLst>' +
  '<dgm:layoutNode name="diagram"><dgm:alg type="snake"><dgm:param type="grDir" val="tL"/></dgm:alg>' +
  '<dgm:shape r:blip=""><dgm:adjLst/></dgm:shape><dgm:presOf/><dgm:constrLst>' +
  '<dgm:constr type="w" for="ch" forName="node" refType="w"/></dgm:constrLst><dgm:ruleLst/>' +
  '<dgm:forEach name="Name0" axis="ch" ptType="node"><dgm:layoutNode name="node">' +
  '<dgm:alg type="tx"/><dgm:shape type="rect" r:blip=""><dgm:adjLst/></dgm:shape>' +
  '<dgm:presOf axis="desc" ptType="node"/></dgm:layoutNode></dgm:forEach></dgm:layoutNode>' +
  '</dgm:layoutDef>';

const DIAGRAM_STYLE =
  `<dgm:styleDef xmlns:dgm="${DGM}" xmlns:a="${A}" ` +
  'uniqueId="urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1"><dgm:title val=""/>' +
  '<dgm:desc val=""/><dgm:catLst><dgm:cat type="simple" pri="10100"/></dgm:catLst>' +
  '<dgm:styleLbl name="node1"><dgm:scene3d><a:camera prst="orthographicFront"/>' +
  '<a:lightRig rig="threePt" dir="t"/></dgm:scene3d><dgm:sp3d/><dgm:txPr/>' +
  '<dgm:style><a:lnRef idx="2"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dgm:style></dgm:styleLbl>' +
  '</dgm:styleDef>';

const DIAGRAM_COLORS =
  `<dgm:colorsDef xmlns:dgm="${DGM}" xmlns:a="${A}" ` +
  'uniqueId="urn:microsoft.com/office/officeart/2005/8/colors/accent1_2"><dgm:title val=""/>' +
  '<dgm:desc val=""/><dgm:catLst><dgm:cat type="accent1" pri="11200"/></dgm:catLst>' +
  '<dgm:styleLbl name="node1"><dgm:fillClrLst meth="repeat"><a:schemeClr val="accent1"/>' +
  '</dgm:fillClrLst><dgm:linClrLst meth="repeat"><a:schemeClr val="lt1"/></dgm:linClrLst>' +
  '<dgm:effectClrLst/><dgm:txLinClrLst/><dgm:txFillClrLst meth="repeat">' +
  '<a:schemeClr val="lt1"/></dgm:txFillClrLst><dgm:txEffectClrLst/></dgm:styleLbl>' +
  '</dgm:colorsDef>';

const shape = (id: string, x: number, text: string) =>
  `<dsp:sp modelId="${id}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>` +
  `<dsp:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="1600000" cy="960000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/>' +
  '</a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" ' +
  'anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" kern="1200"/>' +
  `<a:t>${text}</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x="${x}" y="0"/>` +
  '<a:ext cx="1600000" cy="960000"/></dsp:txXfrm></dsp:sp>';

const DIAGRAM_DRAWING =
  `<dsp:drawing xmlns:dgm="${DGM}" xmlns:dsp="${DSP}" xmlns:a="${A}"><dsp:spTree>` +
  '<dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>' +
  shape('{22222222-0000-0000-0000-000000000000}', 0, 'Draft') +
  shape('{33333333-0000-0000-0000-000000000000}', 1800000, 'Sign') +
  '</dsp:spTree></dsp:drawing>';

const override = (name: string, type: string) =>
  `<Override PartName="/word/${name}" ContentType="${type}"/>`;

/** A line with a chart, a line with a SmartArt diagram, and a line with an Office 2016 chart. */
export const OBJECT_SOURCE: Uint8Array = docx(
  p(t('Chart ') + CHART_RUN + t(' here.')) +
    p(t('Steps ') + DIAGRAM_RUN + t(' here.')) +
    p(t('Box ') + CHART_EX_RUN + t(' here.')),
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/charts/chart1.xml': CHART,
      'word/charts/style1.xml': CHART_STYLE,
      'word/charts/colors1.xml': CHART_COLORS,
      'word/charts/_rels/chart1.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>' +
        `<Relationship Id="rId3" Type="${R}/package" Target="../embeddings/Microsoft_Excel_Worksheet.xlsx"/>` +
        '</Relationships>',
      'word/embeddings/Microsoft_Excel_Worksheet.xlsx': WORKBOOK,
      'word/charts/chartEx1.xml': CHART_EX,
      'word/charts/style2.xml': CHART_STYLE,
      'word/charts/colors2.xml': CHART_COLORS,
      'word/charts/_rels/chartEx1.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style2.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors2.xml"/>' +
        `<Relationship Id="rId3" Type="${R}/package" Target="../embeddings/Microsoft_Excel_Worksheet1.xlsx"/>` +
        '</Relationships>',
      'word/embeddings/Microsoft_Excel_Worksheet1.xlsx': WORKBOOK,
      'word/diagrams/data1.xml': DIAGRAM_DATA,
      'word/diagrams/layout1.xml': DIAGRAM_LAYOUT,
      'word/diagrams/quickStyle1.xml': DIAGRAM_STYLE,
      'word/diagrams/colors1.xml': DIAGRAM_COLORS,
      'word/diagrams/drawing1.xml': DIAGRAM_DRAWING,
    },
    contentTypes:
      `${MATRIX_EXTRAS.contentTypes ?? ''}` +
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' +
      override('charts/chart1.xml', `${DML}.chart+xml`) +
      override('charts/style1.xml', 'application/vnd.ms-office.chartstyle+xml') +
      override('charts/colors1.xml', 'application/vnd.ms-office.chartcolorstyle+xml') +
      override('charts/chartEx1.xml', 'application/vnd.ms-office.chartex+xml') +
      override('charts/style2.xml', 'application/vnd.ms-office.chartstyle+xml') +
      override('charts/colors2.xml', 'application/vnd.ms-office.chartcolorstyle+xml') +
      override('diagrams/data1.xml', `${DML}.diagramData+xml`) +
      override('diagrams/layout1.xml', `${DML}.diagramLayout+xml`) +
      override('diagrams/quickStyle1.xml', `${DML}.diagramStyle+xml`) +
      override('diagrams/colors1.xml', `${DML}.diagramColors+xml`) +
      override('diagrams/drawing1.xml', 'application/vnd.ms-office.drawingml.diagramDrawing+xml'),
    relationships:
      `${MATRIX_EXTRAS.relationships ?? ''}` +
      `<Relationship Id="rIdChart" Type="${R}/chart" Target="charts/chart1.xml"/>` +
      '<Relationship Id="rIdChartEx" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"/>' +
      `<Relationship Id="rIdDm" Type="${R}/diagramData" Target="diagrams/data1.xml"/>` +
      `<Relationship Id="rIdLo" Type="${R}/diagramLayout" Target="diagrams/layout1.xml"/>` +
      `<Relationship Id="rIdQs" Type="${R}/diagramQuickStyle" Target="diagrams/quickStyle1.xml"/>` +
      `<Relationship Id="rIdCs" Type="${R}/diagramColors" Target="diagrams/colors1.xml"/>` +
      '<Relationship Id="rIdDsp" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/>',
  }
);

/** The objects of the paste tests: the source line of each, and what its parts hold. */
export const OBJECTS = {
  chart: { line: 0, text: 'Chart ￼ here', parts: /^\/word\/(charts|embeddings)\// },
  diagram: { line: 1, text: 'Steps ￼ here', parts: /^\/word\/diagrams\// },
  chartEx: { line: 2, text: 'Box ￼ here', parts: /^\/word\/(charts|embeddings)\// },
} as const;

/** A line with a chart or a diagram, without its paragraph mark, as the editor copies it. */
export const copiedObject = (kind: keyof typeof OBJECTS): Copied =>
  copiedBetween(OBJECTS[kind].line, OBJECTS[kind].line, OBJECTS[kind].text.length, OBJECT_SOURCE);
