// The element names a pasted fragment may hold, by prefix. The lists were generated from the
// schemas in `reference/ecma-376/part1/schemas/` by a script outside the repository.

import { OFFICE, VML, WORD_VML } from '../../../store/package/legacy-vml-values.ts';
import {
  DRAWINGML_MAIN_NAMESPACE_URI,
  MC_NAMESPACE_URI,
  PIC_NAMESPACE_URI,
  RELATIONSHIPS_NAMESPACE_URI,
  W14_NAMESPACE_URI,
  WML_NAMESPACE_URI,
  WP_NAMESPACE_URI,
  XML_NAMESPACE_URI,
} from '../../../store/package/ooxml-shared.ts';
import { OFFICE_MATH_NAMESPACE_URI } from '../../../store/package/omml-equation.ts';
import {
  CHART_DRAWING_FRAGMENT_ELEMENTS,
  CHART_EX_NAMESPACE_URI,
  CHART_DRAWING_NAMESPACE_URI,
  CHART_FRAGMENT_ELEMENTS,
  CHART_NAMESPACE_URI,
  DIAGRAM_FRAGMENT_ELEMENTS,
  DIAGRAM_NAMESPACE_URI,
  frozenNames,
  frozenSet,
  isObjectExtensionText,
  OBJECT_EXTENSION_NAMESPACES,
  OBJECT_TEXT_ELEMENTS,
  prefixOf,
} from './fragment-names-objects.ts';

/** The `w15` namespace (Word 2013). */
export const W15_NAMESPACE_URI = 'http://schemas.microsoft.com/office/word/2012/wordml';

/**
 * `w` element names: `wml.xsd` (the Strict schema, 546 names), and five Transitional names
 * Word writes: `pict` (a VML picture or shape), `txbxContent` (a text box's paragraphs),
 * `fldData`, `numberingChange` and `legacy`. The engine keeps each as a node.
 */
export const WML_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'abstractNum abstractNumId active activeRecord activeWritingStyle addressFieldName ' +
    'adjustLineHeightInTable adjustRightInd alias aliases alignBordersAndEdges allowPNG altChunk ' +
    'altChunkPr altName alwaysMergeEmptyNamespace alwaysShowPlaceholderText annotationRef ' +
    'applyBreakingRules attachedSchema attachedTemplate attr autoCaption autoCaptions ' +
    'autoFormatOverride autoHyphenation autoRedefine autoSpaceDE autoSpaceDN b bCs background ' +
    'balanceSingleByteDoubleByteWidth bar basedOn bdo bdr behavior behaviors between ' +
    'bibliography bidi bidiVisual blockQuote body bodyDiv bookFoldPrinting ' +
    'bookFoldPrintingSheets bookFoldRevPrinting bookmarkEnd bookmarkStart ' +
    'bordersDoNotSurroundFooter bordersDoNotSurroundHeader bottom br calcOnExit calendar ' +
    'cantSplit caps caption captions category cellDel cellIns cellMerge characterSpacingControl ' +
    'charset checkBox checkErrors checked citation clickAndTypeStyle clrSchemeMapping cnfStyle ' +
    'col colDelim color cols column comboBox comment commentRangeEnd commentRangeStart ' +
    'commentReference comments compat compatSetting connectString consecutiveHyphenLimit ' +
    'contentPart contextualSpacing continuationSeparator control cr cs customXml ' +
    'customXmlDelRangeEnd customXmlDelRangeStart customXmlInsRangeEnd customXmlInsRangeStart ' +
    'customXmlMoveFromRangeEnd customXmlMoveFromRangeStart customXmlMoveToRangeEnd ' +
    'customXmlMoveToRangeStart customXmlPr dataBinding dataSource dataType date dateFormat ' +
    'dayLong dayShort ddList decimalSymbol default defaultTabMark defaultTableStyle del ' +
    'delInstrText delText description destination dir dirty displayBackgroundShape ' +
    'displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery div divBdr divId divs ' +
    'divsChild doNotAutoCompressPictures doNotDemarcateInvalidXml doNotDisplayPageBoundaries ' +
    'doNotEmbedSmartTags doNotExpandShiftReturn doNotHyphenateCaps doNotIncludeSubdocsInStats ' +
    'doNotLeaveBackslashAlone doNotOrganizeInFolder doNotRelyOnCSS doNotSaveAsSingleFile ' +
    'doNotShadeFormData doNotSuppressBlankLines doNotTrackFormatting doNotTrackMoves ' +
    'doNotUseLongFileNames doNotUseMarginsForDrawingGridOrigin doNotValidateAgainstSchema ' +
    'docDefaults docGrid docPart docPartBody docPartCategory docPartGallery docPartList ' +
    'docPartObj docPartPr docPartUnique docParts docVar docVars document documentProtection ' +
    'documentType drawing drawingGridHorizontalOrigin drawingGridHorizontalSpacing ' +
    'drawingGridVerticalOrigin drawingGridVerticalSpacing dropDownList dstrike dynamicAddress ' +
    'eastAsianLayout effect em embedBold embedBoldItalic embedItalic embedRegular ' +
    'embedSystemFonts embedTrueTypeFonts emboss enabled encoding end endnote endnotePr ' +
    'endnoteRef endnoteReference endnotes entryMacro equation evenAndOddHeaders exitMacro fHdr ' +
    'family ffData fieldMapData fitText flatBorders fldChar fldSimple font fonts footerReference ' +
    'footnote footnotePr footnoteRef footnoteReference footnotes forceUpgrade formProt format ' +
    'formsDesign frame frameLayout framePr frameset framesetSplitbar ftr gallery ' +
    'glossaryDocument gridAfter gridBefore gridCol gridSpan group guid gutterAtTop hdr header ' +
    'headerReference headerSource headers helpText hidden hideGrammaticalErrors hideMark ' +
    'hideSpellingErrors highlight hps hpsBaseText hpsRaise hyperlink hyphenationZone i iCs id ' +
    'ignoreMixedContent ilvl imprint ind ins insideH insideV instrText isLgl jc keepLines ' +
    'keepNext kern kinsoku label lang lastRenderedPageBreak latentStyles left lid link ' +
    'linkStyles linkToQuery linkedToFile listEntry listItem listSeparator lnNumType lock locked ' +
    'longDesc lsdException lvl lvlJc lvlOverride lvlPicBulletId lvlRestart lvlText ' +
    'mailAsAttachment mailMerge mailSubject mainDocumentType mappedName marBottom marH marLeft ' +
    'marRight marTop marW matchSrc maxLength mirrorIndents mirrorMargins monthLong monthShort ' +
    'moveFrom moveFromRangeEnd moveFromRangeStart moveTo moveToRangeEnd moveToRangeStart movie ' +
    'multiLevelType name next noBorder noBreakHyphen noEndnote noLineBreaksAfter ' +
    'noLineBreaksBefore noProof noPunctuationKerning noResizeAllowed noWrap notTrueType nsid num ' +
    'numFmt numId numIdMacAtCleanup numPicBullet numPr numRestart numStart numStyleLink ' +
    'numbering oMath object objectEmbed objectLink odso optimizeForBrowser outline outlineLvl ' +
    'overflowPunct p pBdr pPr pPrChange pPrDefault pStyle pageBreakBefore panose1 paperSrc ' +
    'permEnd permStart personal personalCompose personalReply pgBorders pgMar pgNum pgNumType ' +
    'pgSz picture pitch pixelsPerInch placeholder pos position printFormsData ' +
    'printFractionalCharacterWidth printPostScriptOverText printTwoOnOne printerSettings ' +
    'proofErr proofState ptab qFormat query r rFonts rPr rPrChange rPrDefault rStyle ' +
    'readModeInkLockDown recipientData recipients removeDateAndTime removePersonalInformation ' +
    'result revisionView richText right rsid rsidRoot rsids rt rtl rtlGutter ruby rubyAlign ' +
    'rubyBase rubyPr saveFormsData saveInvalidXml savePreviewPicture saveSmartTagsAsXml ' +
    'saveSubsetFonts saveThroughXslt saveXmlDataOnly scrollbar sdt sdtContent sdtEndPr sdtPr ' +
    'sectPr sectPrChange semiHidden separator settings shadow shd showEnvelope showXMLTags ' +
    'showingPlcHdr sig size sizeAuto smallCaps smartTag smartTagPr smartTagType snapToGrid ' +
    'softHyphen sourceFileName spaceForUL spacing specVanish src start startOverride statusText ' +
    'storeMappedDataAs strictFirstAndLastChars strike style styleLink styleLockQFSet ' +
    'styleLockTheme stylePaneFormatFilter stylePaneSortMethod styles subDoc suff summaryLength ' +
    'suppressAutoHyphens suppressLineNumbers suppressOverlap sym sz szCs t tab tabIndex table ' +
    'tabs tag targetScreenSz tbl tblBorders tblCaption tblCellMar tblCellSpacing tblDescription ' +
    'tblGrid tblGridChange tblHeader tblInd tblLayout tblLook tblOverlap tblPr tblPrChange ' +
    'tblPrEx tblPrExChange tblStyle tblStyleColBandSize tblStylePr tblStyleRowBandSize tblW ' +
    'tblpPr tc tcBorders tcFitText tcMar tcPr tcPrChange tcW temporary text textAlignment ' +
    'textDirection textInput textboxTightWrap themeFontLang title titlePg tl2br tmpl top ' +
    'topLinePunct tr tr2bl trHeight trPr trPrChange trackRevisions type types u udl uiPriority ' +
    'ulTrailSpace unhideWhenUsed uniqueTag updateFields useXSLTWhenSaving vAlign vMerge vanish ' +
    'vertAlign view viewMergedData w wAfter wBefore webHidden webSettings widowControl wordWrap ' +
    'writeProtection yearLong yearShort zoom ' +
    'fldData legacy numberingChange pict txbxContent'
);

/** `m` element names: `shared-math.xsd` (124 names). */
export const MATH_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'acc accPr aln alnScr argPr argSz bar barPr baseJc begChr borderBox borderBoxPr box boxPr ' +
    'brk brkBin brkBinSub cGp cGpRule cSp chr count ctrlPr d dPr defJc deg degHide den diff ' +
    'dispDef e endChr eqArr eqArrPr f fName fPr func funcPr groupChr groupChrPr grow hideBot ' +
    'hideLeft hideRight hideTop intLim interSp intraSp jc lMargin lim limLoc limLow limLowPr ' +
    'limUpp limUppPr lit m mPr mathFont mathPr maxDist mc mcJc mcPr mcs mr nary naryLim naryPr ' +
    'noBreak nor num oMath oMathPara oMathParaPr objDist opEmu phant phantPr plcHide pos postSp ' +
    'preSp r rMargin rPr rSp rSpRule rad radPr sPre sPrePr sSub sSubPr sSubSup sSubSupPr sSup ' +
    'sSupPr scr sepChr show shp smallFrac strikeBLTR strikeH strikeTLBR strikeV sty sub subHide ' +
    'sup supHide t transp type vertJc wrapIndent wrapRight zeroAsc zeroDesc zeroWid'
);

/**
 * `w14` element names: the check box control and its states, ligatures, the text outline the
 * layout paints (`textOutline` with `solidFill`, `srgbClr`, `prstDash` and its joins), and the
 * OpenType settings Word writes beside ligatures.
 */
export const W14_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'checkbox checked checkedState uncheckedState ligatures numForm numSpacing stylisticSets ' +
    'styleSet cntxtAlts textOutline solidFill srgbClr prstDash round bevel miter'
);

/**
 * `w15` element names: a control's color, appearance and data binding, the repeating section
 * and its item with their settings, the web extension marks, and a collapsed heading.
 */
export const W15_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'color appearance dataBinding repeatingSection repeatingSectionItem sectionTitle ' +
    'doNotAllowInsertDeleteSection webExtensionLinked webExtensionCreated collapsed'
);

/**
 * `wp` element names: `dml-wordprocessingDrawing.xsd` (43 names), a drawing's frame. A
 * drawing's attributes have no prefix; its image is `r:embed` on `a:blip`, an index into the
 * change's media.
 */
export const WP_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'align anchor bg bodyPr cNvCnPr cNvContentPartPr cNvFrPr cNvGraphicFramePr cNvGrpSpPr cNvPr ' +
    'cNvSpPr contentPart docPr effectExtent extLst extent graphicFrame grpSp grpSpPr inline ' +
    'lineTo linkedTxbx nvContentPartPr posOffset positionH positionV simplePos spPr start style ' +
    'txbx txbxContent wgp whole wpc wrapNone wrapPolygon wrapSquare wrapThrough wrapTight ' +
    'wrapTopAndBottom wsp xfrm'
);

/** `a` element names: `dml-main.xsd` (321 names), a drawing's graphic, image and geometry. */
export const DRAWINGML_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'accent1 accent2 accent3 accent4 accent5 accent6 ahLst ahPolar ahXY alpha alphaBiLevel ' +
    'alphaCeiling alphaFloor alphaInv alphaMod alphaModFix alphaOff alphaOutset alphaRepl anchor ' +
    'arcTo audioCd audioFile avLst backdrop band1H band1V band2H band2V bevel bevelB bevelT ' +
    'bgClr bgFillStyleLst biLevel bldChart bldDgm blend blip blipFill blue blueMod blueOff blur ' +
    'bodyPr bottom br buAutoNum buBlip buChar buClr buClrTx buFont buFontTx buNone buSzPct ' +
    'buSzPts buSzTx cNvCxnSpPr cNvGraphicFramePr cNvGrpSpPr cNvPicPr cNvPr cNvSpPr camera cell3D ' +
    'chExt chOff chart close clrChange clrFrom clrMap clrRepl clrScheme clrTo comp cont ' +
    'contourClr cpLocks cs cubicBezTo custClr custClrLst custDash custGeom cxn cxnLst cxnSp ' +
    'cxnSpLocks defPPr defRPr dgm dk1 dk2 ds duotone ea effect effectDag effectLst effectRef ' +
    'effectStyle effectStyleLst end endCxn endParaRPr ext extLst extraClrScheme ' +
    'extraClrSchemeLst extrusionClr fgClr fill fillOverlay fillRect fillRef fillStyleLst ' +
    'fillToRect firstCol firstRow flatTx fld fmtScheme folHlink font fontRef fontScheme gamma gd ' +
    'gdLst glow gradFill graphic graphicData graphicFrame graphicFrameLocks gray grayscl green ' +
    'greenMod greenOff gridCol grpFill grpSp grpSpLocks grpSpPr gs gsLst headEnd header headers ' +
    'highlight hlink hlinkClick hlinkHover hlinkMouseOver hsl hslClr hue hueMod hueOff innerShdw ' +
    'insideH insideV inv invGamma lastCol lastRow latin left lightRig lin ln lnB lnBlToTr lnDef ' +
    'lnL lnR lnRef lnSpc lnStyleLst lnT lnTlToBr lnTo lstStyle lt1 lt2 lum lumMod lumOff lvl1pPr ' +
    'lvl2pPr lvl3pPr lvl4pPr lvl5pPr lvl6pPr lvl7pPr lvl8pPr lvl9pPr majorFont masterClrMapping ' +
    'minorFont miter moveTo neCell noAutofit noFill norm normAutofit nvCxnSpPr nvGraphicFramePr ' +
    'nvGrpSpPr nvPicPr nvSpPr nwCell objectDefaults off outerShdw overrideClrMapping p pPr path ' +
    'pathLst pattFill pic picLocks pos prstClr prstDash prstGeom prstShdw prstTxWarp pt ' +
    'quadBezTo quickTimeFile r rPr rect red redMod redOff reflection relOff right rot round rtl ' +
    'sat satMod satOff scene3d schemeClr scrgbClr seCell shade snd softEdge solidFill sp sp3d ' +
    'spAutoFit spDef spLocks spPr spcAft spcBef spcPct spcPts srcRect srgbClr st stCxn stretch ' +
    'style swCell sx sy sym sysClr t tab tabLst tableStyle tableStyleId tailEnd tbl tblBg ' +
    'tblGrid tblPr tblStyle tblStyleLst tc tcBdr tcPr tcStyle tcTxStyle theme themeElements ' +
    'themeManager themeOverride tile tileRect tint tl2br top tr tr2bl txBody txDef txSp uFill ' +
    'uFillTx uLn uLnTx up useSpRect videoFile wavAudioFile wholeTbl xfrm'
);

/** `pic` element names: `dml-picture.xsd` (6 names), a picture. */
export const PICTURE_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'blipFill cNvPicPr cNvPr nvPicPr pic spPr'
);

/** `wps` element names ([MS-ODRAWXML]): a shape, its text box and its text body. */
export const WPS_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'bodyPr cNvCnPr cNvPr cNvSpPr extLst linkedTxbx spPr style txbx wsp'
);

/** `wpg` element names ([MS-ODRAWXML]): a group, a nested group and a graphic frame. */
export const WPG_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'cNvFrPr cNvGrpSpPr cNvPr extLst graphicFrame grpSp grpSpPr wgp xfrm'
);

/** `wpc` element names ([MS-ODRAWXML]): a drawing canvas, its background, a graphic frame. */
export const WPC_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'bg extLst graphicFrame whole wpc'
);

/**
 * `mc` element names: `mc.xsd` (ECMA-376 Part 3), the `mc:AlternateContent` Word writes around
 * a shape, with its `mc:Choice` and `mc:Fallback`.
 */
export const MC_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'AlternateContent Choice Fallback'
);

/**
 * `v` element names: `vml-main.xsd` (23 names; ECMA-376 Part 4, which the repository does not
 * hold). A VML picture's image is `r:id` or `o:relid` on `v:imagedata`, or its `r:pict`.
 */
export const VML_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'arc background curve f fill formulas group h handles image imagedata line oval path ' +
    'polyline rect roundrect shadow shape shapetype stroke textbox textpath'
);

/** `o` element names: `vml-officeDrawing.xsd` (32 names). */
export const OFFICE_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'FieldCodes LinkType LockedField OLEObject bottom callout clippath colormenu colormru column ' +
    'complex diagram entry equationxml extrusion fill idmap ink left lock proxy r regrouptable ' +
    'rel relationtable right rules shapedefaults shapelayout signatureline skew top'
);

/** `w10` element names: `vml-wordprocessingDrawing.xsd` (6 names). */
export const W10_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'anchorlock borderbottom borderleft borderright bordertop wrap'
);

/** `asvg` element names: a picture's SVG version, whose image is `r:embed` into the media. */
export const SVG_BLIP_FRAGMENT_ELEMENTS: ReadonlySet<string> = frozenNames('svgBlip');

/** The `wps` namespace (Word 2010 shapes). */
export const WPS_NAMESPACE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';

/** The `wpg` namespace (Word 2010 groups). */
export const WPG_NAMESPACE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

/** The `wpc` namespace (Word 2010 drawing canvases). */
export const WPC_NAMESPACE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';

/** The `asvg` namespace (Office 2016 SVG pictures). */
export const SVG_BLIP_NAMESPACE_URI = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** The `wp14` namespace (Word 2010 drawing frames). */
export const WP14_NAMESPACE_URI =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';

/** The `a14` namespace (Office 2010 DrawingML). */
export const A14_NAMESPACE_URI = 'http://schemas.microsoft.com/office/drawing/2010/main';

/** The `a16` namespace (Office 2016 DrawingML). */
export const A16_NAMESPACE_URI = 'http://schemas.microsoft.com/office/drawing/2014/main';

/** The `wpi` namespace: Word 2010's ink, which a choice around ink requires. */
export const WPI_NAMESPACE_URI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';

/** The `adec` namespace (Word 365's "decorative" flag of a drawing's alt text). */
export const ADEC_NAMESPACE_URI = 'http://schemas.microsoft.com/office/drawing/2017/decorative';

/**
 * An extension namespace's list: any bounded name is allowed there (`isFragmentElement`). The
 * drawing extensions (`wp14`, `a14`, `a16`, `adec`) load as inert nodes.
 */
const NO_NAMES: ReadonlySet<string> = frozenSet([]);

/** The prefixes an element name may have, with their namespaces and names. */
const ELEMENT_PREFIXES: ReadonlyMap<string, { uri: string; names: ReadonlySet<string> }> = new Map([
  ['w', { uri: WML_NAMESPACE_URI, names: WML_FRAGMENT_ELEMENTS }],
  ['m', { uri: OFFICE_MATH_NAMESPACE_URI, names: MATH_FRAGMENT_ELEMENTS }],
  ['w14', { uri: W14_NAMESPACE_URI, names: W14_FRAGMENT_ELEMENTS }],
  ['w15', { uri: W15_NAMESPACE_URI, names: W15_FRAGMENT_ELEMENTS }],
  ['wp', { uri: WP_NAMESPACE_URI, names: WP_FRAGMENT_ELEMENTS }],
  ['a', { uri: DRAWINGML_MAIN_NAMESPACE_URI, names: DRAWINGML_FRAGMENT_ELEMENTS }],
  ['pic', { uri: PIC_NAMESPACE_URI, names: PICTURE_FRAGMENT_ELEMENTS }],
  ['wps', { uri: WPS_NAMESPACE_URI, names: WPS_FRAGMENT_ELEMENTS }],
  ['wpg', { uri: WPG_NAMESPACE_URI, names: WPG_FRAGMENT_ELEMENTS }],
  ['wpc', { uri: WPC_NAMESPACE_URI, names: WPC_FRAGMENT_ELEMENTS }],
  ['mc', { uri: MC_NAMESPACE_URI, names: MC_FRAGMENT_ELEMENTS }],
  ['v', { uri: VML, names: VML_FRAGMENT_ELEMENTS }],
  ['o', { uri: OFFICE, names: OFFICE_FRAGMENT_ELEMENTS }],
  ['w10', { uri: WORD_VML, names: W10_FRAGMENT_ELEMENTS }],
  ['asvg', { uri: SVG_BLIP_NAMESPACE_URI, names: SVG_BLIP_FRAGMENT_ELEMENTS }],
  ['wp14', { uri: WP14_NAMESPACE_URI, names: NO_NAMES }],
  ['a14', { uri: A14_NAMESPACE_URI, names: NO_NAMES }],
  ['a16', { uri: A16_NAMESPACE_URI, names: NO_NAMES }],
  ['adec', { uri: ADEC_NAMESPACE_URI, names: NO_NAMES }],
  // Word 2010's ink: a prefix that a choice requires around ink, with no element of its own
  ['wpi', { uri: WPI_NAMESPACE_URI, names: NO_NAMES }],
  ['c', { uri: CHART_NAMESPACE_URI, names: CHART_FRAGMENT_ELEMENTS }],
  ['cdr', { uri: CHART_DRAWING_NAMESPACE_URI, names: CHART_DRAWING_FRAGMENT_ELEMENTS }],
  ['dgm', { uri: DIAGRAM_NAMESPACE_URI, names: DIAGRAM_FRAGMENT_ELEMENTS }],
  ...[...OBJECT_EXTENSION_NAMESPACES].map(
    ([prefix, uri]) => [prefix, { uri, names: NO_NAMES }] as const
  ),
]);

/**
 * The ink content part's elements in `w14` ([MS-DOCX] `CT_WordContentPart`): DrawingML types,
 * whose attributes have no prefix (`bwMode`, a drawing's `id` and `name`, a transform's `rot`).
 */
const INK_PART_ELEMENTS: ReadonlySet<string> = frozenNames(
  'w14:contentPart w14:nvContentPartPr w14:cNvPr w14:cNvContentPartPr w14:xfrm'
);

/** The element prefixes of a drawing, a shape and VML: their attributes have no prefix. */
const DRAWING_PREFIXES: ReadonlySet<string> = frozenSet([
  'wp',
  'a',
  'pic',
  'wps',
  'wpg',
  'wpc',
  'mc',
  'v',
  'o',
  'w10',
  'asvg',
  'wp14',
  'a14',
  'a16',
  'adec',
  'c',
  'cdr',
  'dgm',
  ...OBJECT_EXTENSION_NAMESPACES.keys(),
]);

/** Whether an element, by qualified name, takes attributes without a prefix (a drawing's). */
export function takesPlainAttributes(owner: string): boolean {
  return DRAWING_PREFIXES.has(prefixOf(owner)) || INK_PART_ELEMENTS.has(owner);
}

/**
 * Whether an element of this prefix is VML (`v`, `o`). VML keeps drawing data in attribute
 * values (a path, a polyline's points, `o:gfxdata`, Word's DrawingML copy of the shape), so such
 * a value is bounded by the change's text budget, not by the cap on one value.
 */
export function isVmlPrefix(prefix: string): boolean {
  return prefix === 'v' || prefix === 'o';
}

/** The prefixes an attribute name may have, with their namespaces. */
const ATTRIBUTE_PREFIXES: ReadonlyMap<string, string> = new Map([
  ['w', WML_NAMESPACE_URI],
  ['m', OFFICE_MATH_NAMESPACE_URI],
  ['w14', W14_NAMESPACE_URI],
  ['w15', W15_NAMESPACE_URI],
  ['xml', XML_NAMESPACE_URI],
  ['r', RELATIONSHIPS_NAMESPACE_URI],
  ['v', VML],
  ['o', OFFICE],
  // An Office 2016 chart's own attributes (`cx:externalData/@cx:autoUpdate`)
  ['cx', CHART_EX_NAMESPACE_URI],
]);

/** The namespace of an element prefix, or `undefined` for another prefix. */
export function fragmentElementNamespace(prefix: string): string | undefined {
  return ELEMENT_PREFIXES.get(prefix)?.uri;
}

/**
 * Word's own extension namespaces: the repository holds no schema for them, so any bounded name
 * is allowed there. The `w14` and `w15` lists name what the engine reads; they decide nothing.
 */
const EXTENSIONS: ReadonlySet<string> = frozenSet([
  'w14',
  'w15',
  'wp14',
  'a14',
  'a16',
  'adec',
  ...OBJECT_EXTENSION_NAMESPACES.keys(),
]);
const EXTENSION_NAME = /^[A-Za-z][A-Za-z0-9]{0,63}$/;

/**
 * Whether a local name is allowed in a prefix's namespace: on the list of its schema, or any
 * bounded name in Word's extension namespaces (`EXTENSIONS`: Word's, a drawing's and the chart
 * and diagram extensions). An extension element the engine does not read loads as an inert node.
 */
export function isFragmentElement(prefix: string, localName: string): boolean {
  if (ELEMENT_PREFIXES.get(prefix)?.names.has(localName) === true) return true;
  return EXTENSIONS.has(prefix) && EXTENSION_NAME.test(localName);
}

/** The namespace of an attribute prefix, or `undefined` for another prefix. */
export function fragmentAttributeNamespace(prefix: string): string | undefined {
  return ATTRIBUTE_PREFIXES.get(prefix);
}

/** The element prefix of a namespace, or `undefined` for another namespace. */
export function fragmentElementPrefix(namespaceUri: string): string | undefined {
  for (const [prefix, entry] of ELEMENT_PREFIXES) if (entry.uri === namespaceUri) return prefix;
  return undefined;
}

/** The attribute prefix of a namespace, or `undefined` for another namespace. */
export function fragmentAttributePrefix(namespaceUri: string): string | undefined {
  for (const [prefix, uri] of ATTRIBUTE_PREFIXES) if (uri === namespaceUri) return prefix;
  return undefined;
}

/** Every namespace declaration a written fragment needs, once, on its root element. */
export const FRAGMENT_NAMESPACE_DECLARATIONS: string = [
  ...new Map([
    ...[...ELEMENT_PREFIXES].map(([prefix, entry]) => [prefix, entry.uri] as const),
    ...ATTRIBUTE_PREFIXES,
  ]),
]
  .filter(([prefix]) => prefix !== 'xml')
  .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
  .join(' ');

/**
 * Elements, by qualified name, that hold an object the format does not carry: an ActiveX
 * control (`w:control`, ECMA-376 Part 1, §17.3.3.3), a program, which the editor's copy refuses
 * too. A DrawingML picture or shape (`w:drawing`), a VML one (`w:pict`), an embedded object
 * (`w:object`) and ink (`w14:contentPart`) travel, their parts as media by hash
 * (`fragmentObjectReference`), and so does a list's picture bullet (`w:lvlPicBulletId`, `#<j>`
 * into the change's `bullets`).
 */
export const FRAGMENT_OBJECT_ELEMENTS: ReadonlySet<string> = frozenNames('w:control');

/**
 * Attributes, by key, that name an image or an object by address (VML's `src`, `o:href` and
 * `o:althref`): refused as an object, never fetched, except a picture's address
 * (`isFragmentPictureAddress`).
 */
export const FRAGMENT_OBJECT_ATTRIBUTES: ReadonlySet<string> = frozenNames('src o:href o:althref');

/** The elements whose `r:id` names a link: a run's link, a picture's click or hover link. */
export const FRAGMENT_LINK_OWNERS: ReadonlySet<string> = frozenNames(
  'w:hyperlink a:hlinkClick a:hlinkHover'
);

/**
 * The attributes that name a picture's image, as `<owner>@<key>`: a DrawingML blip, its SVG
 * version, and VML's picture, picture fill and picture stroke, by `r:id` or by `o:relid`, the
 * Office relationship id older Word writes (ECMA-376 Part 4), and a VML picture's `r:pict`.
 */
const MEDIA_REFERENCES: ReadonlySet<string> = frozenNames(
  'a:blip@r:embed asvg:svgBlip@r:embed v:imagedata@r:id v:fill@r:id v:stroke@r:id ' +
    'v:imagedata@o:relid v:fill@o:relid v:stroke@o:relid v:imagedata@r:pict'
);

/** Whether the attribute `key` of an element `owner` names a picture's image. */
export function isFragmentMediaReference(owner: string, key: string): boolean {
  return MEDIA_REFERENCES.has(`${owner}@${key}`);
}

/**
 * The attributes that name a picture's image by address, as `<owner>@<key>`: a DrawingML
 * picture's link, a VML picture's link (`r:href`, the relationship to the linked file), and
 * VML's `src`, `o:href` and `o:althref` on a picture, a picture fill and a picture stroke. Word
 * loads the address when it opens the file, with no click (a tracking request; a `file:` or UNC
 * address can also send the reader's Windows sign-in hash), so the change format holds none
 * (design D18): a pasted picture keeps only its embedded copy.
 */
const PICTURE_ADDRESSES: ReadonlySet<string> = frozenNames(
  'a:blip@r:link v:imagedata@r:href v:imagedata@src v:imagedata@o:href v:imagedata@o:althref ' +
    'v:fill@src v:fill@o:href v:fill@o:althref v:stroke@src v:stroke@o:href v:stroke@o:althref'
);

/** Whether the attribute `key` of an element `owner` names a picture's image by address. */
export function isFragmentPictureAddress(owner: string, key: string): boolean {
  return PICTURE_ADDRESSES.has(`${owner}@${key}`);
}

/**
 * The elements whose children are text, by qualified name: `CT_Text` and its kin, and a
 * drawing's position (`wp:posOffset`, `wp:align`, and `wp14:pctPosHOffset`,
 * `wp14:pctPosVOffset`) and relative size (`wp14:pctWidth`, `wp14:pctHeight`).
 */
export const FRAGMENT_TEXT_ELEMENTS: ReadonlySet<string> = frozenNames(
  'w:t w:delText w:instrText w:delInstrText w:fldData m:t wp:posOffset wp:align ' +
    'wp14:pctWidth wp14:pctHeight wp14:pctPosHOffset wp14:pctPosVOffset ' +
    [...OBJECT_TEXT_ELEMENTS].join(' ')
);

/** Whether an element's children may be text: on the list, or a chart extension's element. */
export function isFragmentTextElement(owner: string): boolean {
  return FRAGMENT_TEXT_ELEMENTS.has(owner) || isObjectExtensionText(owner);
}

/** `Requires` of `mc:Choice`: one to eight fragment prefixes, one space between two. */
const REQUIRES = /^[a-z][a-z0-9]{0,4}(?: [a-z][a-z0-9]{0,4}){0,7}$/;

/** Whether a `mc:Choice/@Requires` value names fragment element prefixes only. */
export function isFragmentRequires(value: string): boolean {
  return (
    REQUIRES.test(value) &&
    value.split(' ').every((prefix) => fragmentElementNamespace(prefix) !== undefined)
  );
}
