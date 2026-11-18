// The element names of a pasted chart's and diagram's (SmartArt) own parts. The lists were
// generated from the schemas by a script outside the repository.

/** A set that refuses every change: `add`, `delete` and `clear` throw. */
export function frozenSet(names: Iterable<string>): ReadonlySet<string> {
  const set = new Set(names);
  const refuse = (): never => {
    throw new TypeError('a fragment name list does not change');
  };
  for (const method of ['add', 'delete', 'clear']) {
    Object.defineProperty(set, method, { value: refuse });
  }

  return Object.freeze(set);
}

/** A frozen set of names written with one space between two. */
export function frozenNames(names: string): ReadonlySet<string> {
  return frozenSet(names.split(' '));
}

/** The prefix of a qualified name (`w` of `w:p`). */
export function prefixOf(qualifiedName: string): string {
  return qualifiedName.slice(0, qualifiedName.indexOf(':'));
}

/** `c` element names: `dml-chart.xsd` (220 names). Their attributes have no prefix. */
export const CHART_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'applyToEnd applyToFront applyToSides area3DChart areaChart auto autoTitleDeleted ' +
    'autoUpdate axId axPos backWall backward bandFmt bandFmts bar3DChart barChart barDir ' +
    'baseTimeUnit bubble3D bubbleChart bubbleScale bubbleSize builtInUnit cat catAx chart ' +
    'chartObject chartSpace clrMapOvr crossAx crossBetween crosses crossesAt custSplit ' +
    'custUnit dLbl dLblPos dLbls dPt dTable data date1904 dateAx delete depthPercent ' +
    'dispBlanksAs dispEq dispRSqr dispUnits dispUnitsLbl doughnutChart downBars dropLines ' +
    'errBarType errBars errDir errValType evenFooter evenHeader explosion ext extLst ' +
    'externalData f firstFooter firstHeader firstSliceAng floor fmtId formatCode formatting ' +
    'forward gapDepth gapWidth grouping h hMode hPercent headerFooter hiLowLines holeSize idx ' +
    'intercept invertIfNegative lang layout layoutTarget lblAlgn lblOffset leaderLines legend ' +
    'legendEntry legendPos line3DChart lineChart logBase lvl majorGridlines majorTickMark ' +
    'majorTimeUnit majorUnit manualLayout marker max min minorGridlines minorTickMark ' +
    'minorTimeUnit minorUnit minus multiLvlStrCache multiLvlStrRef name noEndCap ' +
    'noMultiLvlLbl numCache numFmt numLit numRef oddFooter oddHeader ofPieChart ofPieType ' +
    'order orientation overlap overlay pageMargins pageSetup period perspective pictureFormat ' +
    'pictureOptions pictureStackUnit pie3DChart pieChart pivotFmt pivotFmts pivotSource ' +
    'plotArea plotVisOnly plus printSettings protection pt ptCount rAngAx radarChart ' +
    'radarStyle rich rotX rotY roundedCorners scaling scatterChart scatterStyle secondPiePt ' +
    'secondPieSize selection separator ser serAx serLines shape showBubbleSize showCatName ' +
    'showDLblsOverMax showHorzBorder showKeys showLeaderLines showLegendKey showNegBubbles ' +
    'showOutline showPercent showSerName showVal showVertBorder sideWall size sizeRepresents ' +
    'smooth spPr splitPos splitType stockChart strCache strLit strRef style surface3DChart ' +
    'surfaceChart symbol thickness tickLblPos tickLblSkip tickMarkSkip title trendline ' +
    'trendlineLbl trendlineType tx txPr upBars upDownBars userInterface userShapes v val ' +
    'valAx varyColors view3D w wMode wireframe x xMode xVal y yMode yVal'
);

/** `cdr` element names: `dml-chartDrawing.xsd` (29 names). */
export const CHART_DRAWING_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'absSizeAnchor blipFill cNvCxnSpPr cNvGraphicFramePr cNvGrpSpPr cNvPicPr cNvPr cNvSpPr ' +
    'cxnSp ext from graphicFrame grpSp grpSpPr nvCxnSpPr nvGraphicFramePr nvGrpSpPr nvPicPr ' +
    'nvSpPr pic relSizeAnchor sp spPr style to txBody x xfrm y'
);

/** `dgm` element names: `dml-diagram.xsd` (65 names). */
export const DIAGRAM_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'adj adjLst alg animLvl animOne bg bulletEnabled cat catLst chMax chPref choose clrData ' +
    'colorsDef colorsDefHdr colorsDefHdrLst constr constrLst cxn cxnLst dataModel desc dir ' +
    'effectClrLst else extLst fillClrLst forEach hierBranch if layoutDef layoutDefHdr ' +
    'layoutDefHdrLst layoutNode linClrLst orgChart param prSet presLayoutVars presOf pt ptLst ' +
    'relIds resizeHandles rule ruleLst sampData scene3d shape sp3d spPr style styleData ' +
    'styleDef styleDefHdr styleDefHdrLst styleLbl t title txEffectClrLst txFillClrLst ' +
    'txLinClrLst txPr varLst whole'
);

/** The `c` namespace (a chart). */
export const CHART_NAMESPACE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** The `cdr` namespace (a chart's user shapes). */
export const CHART_DRAWING_NAMESPACE_URI =
  'http://schemas.openxmlformats.org/drawingml/2006/chartDrawing';

/** The `dgm` namespace (a diagram). */
export const DIAGRAM_NAMESPACE_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** An Office 2016 chart's namespace: its elements, and its attributes (`cx:autoUpdate`). */
export const CHART_EX_NAMESPACE_URI = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

/**
 * Office's chart and diagram extension namespaces, by the prefix Word writes ([MS-ODRAWXML]; the
 * repository holds no schema for them). Any bounded name is allowed there.
 */
export const OBJECT_EXTENSION_NAMESPACES: ReadonlyMap<string, string> = new Map([
  ['c14', 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart'],
  ['c15', 'http://schemas.microsoft.com/office/drawing/2012/chart'],
  ['c16', 'http://schemas.microsoft.com/office/drawing/2014/chart'],
  ['c16r2', 'http://schemas.microsoft.com/office/drawing/2015/06/chart'],
  ['c16r3', 'http://schemas.microsoft.com/office/drawing/2017/03/chart'],
  ['cs', 'http://schemas.microsoft.com/office/drawing/2012/chartStyle'],
  ['dsp', 'http://schemas.microsoft.com/office/drawing/2008/diagram'],
  ['dgm14', 'http://schemas.microsoft.com/office/drawing/2010/diagram'],
  // An Office 2016 chart ([MS-ODRAWXML] chartex), and the namespaces its choice requires
  ['cx', CHART_EX_NAMESPACE_URI],
  ['cx1', 'http://schemas.microsoft.com/office/drawing/2015/9/8/chartex'],
  ['cx2', 'http://schemas.microsoft.com/office/drawing/2015/10/21/chartex'],
  ['cx3', 'http://schemas.microsoft.com/office/drawing/2016/5/9/chartex'],
  ['cx4', 'http://schemas.microsoft.com/office/drawing/2016/5/10/chartex'],
  ['cx5', 'http://schemas.microsoft.com/office/drawing/2016/5/11/chartex'],
  ['cx6', 'http://schemas.microsoft.com/office/drawing/2016/5/12/chartex'],
  ['cx7', 'http://schemas.microsoft.com/office/drawing/2016/5/13/chartex'],
  ['cx8', 'http://schemas.microsoft.com/office/drawing/2016/5/14/chartex'],
]);

/**
 * The elements of a chart's or a diagram's parts whose children are text, by qualified name: the
 * string-typed elements of `dml-chart.xsd` (a cell's cached value, a formula, a number format, a
 * header), a user shape's anchor (`cdr:x`, `cdr:y`), and DrawingML's text (`a:t`). Text in a
 * chart extension element (`c15:f`, a data label range) is allowed too
 * (`isObjectExtensionText`).
 */
export const OBJECT_TEXT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'c:evenFooter c:evenHeader c:f c:firstFooter c:firstHeader c:formatCode c:name ' +
    'c:oddFooter c:oddHeader c:separator c:v cdr:x cdr:y a:t'
);

/** The chart extension prefixes whose elements may hold text: a range, a formula, a cache. */
const TEXT_EXTENSIONS: ReadonlySet<string> = frozenSet([
  'c14',
  'c15',
  'c16',
  'c16r2',
  'c16r3',
  // An Office 2016 chart's data and text: a formula, a value, a point
  'cx',
]);

/** Whether an element of a chart extension may hold text. */
export function isObjectExtensionText(owner: string): boolean {
  return TEXT_EXTENSIONS.has(prefixOf(owner));
}

/**
 * What a part of a pasted chart or diagram is: named after the relationship type that points
 * at it.
 *
 * @public
 */
export type OrderedFragmentPartType =
  | 'chart'
  | 'chartEx'
  | 'chartStyle'
  | 'chartColorStyle'
  | 'chartUserShapes'
  | 'themeOverride'
  | 'diagramData'
  | 'diagramLayout'
  | 'diagramQuickStyle'
  | 'diagramColors'
  | 'diagramDrawing';

/** Each part type's root element, by qualified name. */
export const FRAGMENT_PART_ROOTS: Readonly<Record<OrderedFragmentPartType, string>> = Object.freeze(
  {
    chart: 'c:chartSpace',
    chartEx: 'cx:chartSpace',
    chartStyle: 'cs:chartStyle',
    chartColorStyle: 'cs:colorStyle',
    chartUserShapes: 'c:userShapes',
    themeOverride: 'a:themeOverride',
    diagramData: 'dgm:dataModel',
    diagramLayout: 'dgm:layoutDef',
    diagramQuickStyle: 'dgm:styleDef',
    diagramColors: 'dgm:colorsDef',
    diagramDrawing: 'dsp:drawing',
  }
);

/** Every part type, once. */
export const FRAGMENT_PART_TYPES: readonly OrderedFragmentPartType[] = Object.freeze(
  Object.keys(FRAGMENT_PART_ROOTS) as OrderedFragmentPartType[]
);

/**
 * The part types a part's own relationships may name, beside media: a chart's style, colors,
 * user shapes and theme override. A chart's media may be its embedded workbook or a picture; the
 * other parts' media are pictures.
 */
export const FRAGMENT_PART_CHILDREN: Readonly<
  Record<OrderedFragmentPartType, readonly OrderedFragmentPartType[]>
> = Object.freeze({
  chart: ['chartStyle', 'chartColorStyle', 'chartUserShapes', 'themeOverride'],
  chartEx: ['chartStyle', 'chartColorStyle'],
  chartStyle: [],
  chartColorStyle: [],
  chartUserShapes: [],
  themeOverride: [],
  diagramData: [],
  diagramLayout: [],
  diagramQuickStyle: [],
  diagramColors: [],
  diagramDrawing: [],
});

/** The part a block's attribute names, as `<owner>@<key>`: a chart, a diagram's four parts. */
const PART_REFERENCES: ReadonlyMap<string, OrderedFragmentPartType> = new Map([
  ['c:chart@r:id', 'chart'],
  ['cx:chart@r:id', 'chartEx'],
  ['dgm:relIds@r:dm', 'diagramData'],
  ['dgm:relIds@r:lo', 'diagramLayout'],
  ['dgm:relIds@r:qs', 'diagramQuickStyle'],
  ['dgm:relIds@r:cs', 'diagramColors'],
]);

/** The part type the attribute `key` of a block's element `owner` names, if it names one. */
export function fragmentPartReference(
  owner: string,
  key: string
): OrderedFragmentPartType | undefined {
  return PART_REFERENCES.get(`${owner}@${key}`);
}

/**
 * What an embedded object's or ink's reference names: a part kept as bytes, held by hash as the
 * change's media and never opened. `object`: an OLE part or an embedded Office file
 * (`ORDERED_EMBEDDED_OBJECT_TYPES`); `ink`: an InkML part.
 */
export type OrderedFragmentObjectUse = 'object' | 'ink';

/**
 * The attributes that name an embedded object's part or an ink part, as `<owner>@<key>`: an
 * embedded object (Transitional `o:OLEObject`, Strict `w:objectEmbed`), and ink (Word 2010's
 * `w14:contentPart`, and the content parts of `wml.xsd` and `dml-wordprocessingDrawing.xsd`).
 */
const OBJECT_REFERENCES: ReadonlyMap<string, OrderedFragmentObjectUse> = new Map([
  ['o:OLEObject@r:id', 'object'],
  ['w:objectEmbed@r:id', 'object'],
  ['w14:contentPart@r:id', 'ink'],
  ['w:contentPart@r:id', 'ink'],
  ['wp:contentPart@r:id', 'ink'],
]);

/**
 * The relationship types by which a story names an object's part: an OLE part
 * (ECMA-376 Part 1, §15.2.10) or an embedded Office file (`package`, §15.2.11); an ink part.
 */
export const FRAGMENT_OBJECT_RELATIONSHIP_TYPES: Readonly<
  Record<'oleObject' | 'package' | 'ink', string>
> = Object.freeze({
  oleObject: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
  package: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
  ink: 'http://schemas.microsoft.com/office/2011/relationships/ink',
});

/** What the attribute `key` of a block's element `owner` names, if it names an object's part. */
export function fragmentObjectReference(
  owner: string,
  key: string
): OrderedFragmentObjectUse | undefined {
  return OBJECT_REFERENCES.get(`${owner}@${key}`);
}

/** A diagram data part's name for its drawing: `#<k>` into the change's parts. */
export const DIAGRAM_DRAWING_REFERENCE = 'dsp:dataModelExt@relId';

const OFFICE_RELATIONSHIPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DML_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml';

/** Each part type's content type, as Word writes it. */
export const FRAGMENT_PART_CONTENT_TYPES: Readonly<Record<OrderedFragmentPartType, string>> =
  Object.freeze({
    chart: `${DML_TYPE}.chart+xml`,
    chartEx: 'application/vnd.ms-office.chartex+xml',
    chartStyle: 'application/vnd.ms-office.chartstyle+xml',
    chartColorStyle: 'application/vnd.ms-office.chartcolorstyle+xml',
    chartUserShapes: `${DML_TYPE}.chartshapes+xml`,
    themeOverride: 'application/vnd.openxmlformats-officedocument.themeOverride+xml',
    diagramData: `${DML_TYPE}.diagramData+xml`,
    diagramLayout: `${DML_TYPE}.diagramLayout+xml`,
    diagramQuickStyle: `${DML_TYPE}.diagramStyle+xml`,
    diagramColors: `${DML_TYPE}.diagramColors+xml`,
    diagramDrawing: 'application/vnd.ms-office.drawingml.diagramDrawing+xml',
  });

/** The type of the relationship to each part type, and to an embedded workbook (`package`). */
export const FRAGMENT_PART_RELATIONSHIP_TYPES: Readonly<
  Record<OrderedFragmentPartType | 'package', string>
> = Object.freeze({
  chart: `${OFFICE_RELATIONSHIPS}/chart`,
  chartEx: 'http://schemas.microsoft.com/office/2014/relationships/chartEx',
  chartStyle: 'http://schemas.microsoft.com/office/2011/relationships/chartStyle',
  chartColorStyle: 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle',
  chartUserShapes: `${OFFICE_RELATIONSHIPS}/chartUserShapes`,
  themeOverride: `${OFFICE_RELATIONSHIPS}/themeOverride`,
  diagramData: `${OFFICE_RELATIONSHIPS}/diagramData`,
  diagramLayout: `${OFFICE_RELATIONSHIPS}/diagramLayout`,
  diagramQuickStyle: `${OFFICE_RELATIONSHIPS}/diagramQuickStyle`,
  diagramColors: `${OFFICE_RELATIONSHIPS}/diagramColors`,
  diagramDrawing: 'http://schemas.microsoft.com/office/2007/relationships/diagramDrawing',
  package: FRAGMENT_OBJECT_RELATIONSHIP_TYPES.package,
});
