// A pasted fragment's round trip (group 9.10): engine nodes → element trees (`fragmentWire`)
// → JSON → the change reader (`fragmentElements`) → engine nodes again (`fragmentNodes`,
// `fragmentResourceNodes`). The nodes come back with the same fingerprint and the same kinds,
// apart from the paragraph ids the apply mints, and a value that holds markup stays text.

import { describe, expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import {
  fragmentBudget,
  fragmentElements,
  type OrderedFragmentElement,
  type OrderedFragmentRoot,
} from '../change/fragment.ts';
import { fragmentWire, type OrderedFragmentWireReferences } from '../client/fragment-wire.ts';
import {
  fragmentNodes,
  fragmentResourceNodes,
  type OrderedFragmentResolver,
} from '../edits/fragment-nodes.ts';
import { M, MC, p, R, t, V, W, WPS } from './word-fixtures.ts';

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml';
const DML = 'http://schemas.openxmlformats.org/drawingml/2006';
const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const NS =
  `xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w15="${W15}" xmlns:r="${R}" xmlns:m="${M}" ` +
  `xmlns:wp="${DML}/wordprocessingDrawing" xmlns:a="${DML}/main" xmlns:pic="${DML}/picture" ` +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" ' +
  'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" ' +
  `xmlns:mc="${MC}" xmlns:wps="${WPS}" xmlns:v="${V}" ` +
  'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"';
const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** The element children of a part root read from literal XML. */
function read(root: 'document' | 'numbering' | 'styles', inner: string): OoxmlElement[] {
  const xml =
    root === 'document'
      ? `<w:document ${NS}><w:body>${inner}</w:body></w:document>`
      : `<w:${root} ${NS}>${inner}</w:${root}>`;
  const part = readOoxmlPart(xml, {
    name: `/word/${root}.xml`,
    contentType: `${WML_CT}.${root === 'document' ? 'document.main' : root}+xml`,
  });
  if (!part.ok) throw new Error(`unreadable fixture: ${part.reason}`);
  const holder = root === 'document' ? part.part.root.children[0] : part.part.root;
  if (holder === undefined || holder.kind === 'textValue') throw new Error('no holder');
  return (holder.children as readonly OoxmlNode[]).filter(
    (child): child is OoxmlElement => child.kind !== 'textValue'
  );
}

/**
 * The wire's references and the apply's resolver: list 7, link `rIdLink`, notes 1 and 2, images
 * `rIdImage` and `rIdSvg`.
 */
const IMAGES = ['rIdImage', 'rIdSvg'];
const WIRE: OrderedFragmentWireReferences = {
  numId: (value) => (value === '7' ? '#0' : value),
  relationship: (id) => (id === 'rIdLink' ? '#0' : `unknown ${id}`),
  media: (id) => (IMAGES.includes(id) ? `#${IMAGES.indexOf(id)}` : `unknown ${id}`),
  note: (kind, id) => (kind === 'footnote' && id === '1' ? '#0' : kind === 'endnote' ? '#1' : id),
};
const RESOLVE: OrderedFragmentResolver = {
  numId: (value) => (value === '#0' ? '7' : value),
  relationship: (value) => (value === '#0' ? 'rIdLink' : `unknown ${value}`),
  media: (value) => IMAGES[Number(value.slice(1))] ?? `unknown ${value}`,
  drawingId: () => '5',
  note: (value) => (value === '#0' ? '1' : value === '#1' ? '2' : `unknown ${value}`),
  part: (value) => `unknown ${value}`,
  object: (value) => `unknown ${value}`,
};
const REFERENCES = { lists: 1, links: 1, notes: 2, media: 2 };

/** The trees as a change carries them: wire, JSON, and the change reader. */
function carried(nodes: readonly OoxmlNode[], root: OrderedFragmentRoot): OrderedFragmentElement[] {
  const json: unknown = JSON.parse(JSON.stringify(fragmentWire(nodes, WIRE)));
  return fragmentElements(json, 'blocks', root, REFERENCES, fragmentBudget());
}

/** Every element's kind, depth first. */
function kinds(nodes: readonly OoxmlNode[]): string[] {
  const out: string[] = [];
  const visit = (node: OoxmlNode): void => {
    out.push(node.kind === 'textValue' ? 'text' : `${node.kind}:${node.localName}`);
    if (node.kind !== 'textValue') for (const child of node.children) visit(child);
  };
  for (const node of nodes) visit(node);
  return out;
}

/** The fixture without the ids the wire leaves out: paragraphs' and a drawing's Word ids. */
const withoutParagraphIds = (xml: string): string =>
  xml.replace(/ (w14:paraId|w14:textId|wp14:anchorId|wp14:editId|w14:anchorId)="[0-9A-F]{8}"/g, '');

const field = (instruction: string, result: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t(result)}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

const HEADING =
  '<w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00A1B2C3" w:rsidRDefault="00A1B2C3">' +
  '<w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>' +
  '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>' +
  '<w:bottom w:val="double" w:sz="6" w:space="1" w:color="FF0000"/></w:pBdr>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="D9E2F3" w:themeFill="accent1" w:themeFillTint="33"/>' +
  '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>' +
  '<w:spacing w:before="240" w:after="120" w:line="276" w:lineRule="auto"/>' +
  '<w:ind w:left="720" w:hanging="360"/><w:jc w:val="both"/>' +
  '<w:rPr><w:b/><w:color w:val="2F5496"/></w:rPr></w:pPr>' +
  '<w:r><w:rPr><w:rStyle w:val="Strong"/>' +
  '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsia="MS Mincho" w:cs="Arial"/>' +
  '<w:b/><w:i/><w:color w:val="1F3864" w:themeColor="accent1" w:themeShade="80"/><w:sz w:val="28"/>' +
  '<w:highlight w:val="yellow"/><w:lang w:val="en-GB" w:eastAsia="ja-JP"/>' +
  '<w14:ligatures w14:val="standard"/></w:rPr><w:t xml:space="preserve">  Lead  </w:t></w:r>' +
  '<w:r><w:tab/><w:t>After a tab</w:t><w:br w:type="page"/><w:br w:type="textWrapping" w:clear="all"/>' +
  '<w:sym w:font="Wingdings" w:char="F0E0"/><w:softHyphen/><w:noBreakHyphen/><w:t>end</w:t></w:r></w:p>';

const LINKS =
  '<w:p w14:paraId="2B3C4D5E" w14:textId="2B3C4D5F"><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr>' +
  '<w:bookmarkStart w:id="0" w:name="Intro"/>' +
  '<w:hyperlink r:id="rIdLink" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
  '<w:t>external</w:t></w:r></w:hyperlink>' +
  t(' and ') +
  '<w:hyperlink w:anchor="Intro" w:tooltip="Go &quot;up&quot;"><w:r><w:t>anchor</w:t></w:r></w:hyperlink>' +
  '<w:bookmarkEnd w:id="0"/>' +
  '<w:fldSimple w:instr=" PAGE \\* MERGEFORMAT "><w:r><w:t>3</w:t></w:r></w:fldSimple>' +
  field('DATE \\@ "d MMMM yyyy"', '1 May 2026') +
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
  '<w:r><w:endnoteReference w:id="2"/></w:r>' +
  '<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:tag w:val="name"/><w:id w:val="12345"/>' +
  '<w15:color w:val="FF0000"/><w15:appearance w15:val="tags"/><w:text/></w:sdtPr>' +
  '<w:sdtContent><w:r><w:t>Ann</w:t></w:r></w:sdtContent></w:sdt>' +
  '<w:sdt><w:sdtPr><w:id w:val="-7"/><w14:checkbox><w14:checked w14:val="1"/>' +
  '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox></w:sdtPr>' +
  '<w:sdtContent><w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>' +
  '<w:t>☒</w:t></w:r></w:sdtContent></w:sdt></w:p>';

const EQUATION = p(
  '<m:oMathPara><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>' +
    '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>' +
    '<m:r><m:rPr><m:sty m:val="p"/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>' +
    '<m:t>+1</m:t></m:r></m:oMath></m:oMathPara>'
);

const cell = (width: number, props: string, content: string): string =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${props}</w:tcPr>${content}</w:tc>`;
const NESTED =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
  `<w:tr>${cell(2000, '', p(t('Inner')))}</w:tr></w:tbl>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>' +
  '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>' +
  '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
  '</w:tblPr><w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>' +
  '<w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr>' +
  cell(
    9350,
    '<w:gridSpan w:val="2"/><w:tcBorders><w:bottom w:val="double" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders>' +
      '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>',
    p(t('Head'))
  ) +
  '</w:tr><w:tr>' +
  cell(4675, '<w:vMerge w:val="restart"/>', p(t('Merged'))) +
  cell(4675, '', `${NESTED}<w:p/>`) +
  '</w:tr><w:tr>' +
  cell(4675, '<w:vMerge/>', '<w:p/>') +
  cell(4675, '', p(t('Cell'))) +
  '</w:tr></w:tbl>';

const BLOCK_CONTROL =
  '<w:sdt><w:sdtPr><w:alias w:val="Terms"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>' +
  '<w:p w14:paraId="0ABCDEF1" w14:textId="0ABCDEF2"><w:proofErr w:type="spellStart"/>' +
  `${t('Teh terms')}<w:proofErr w:type="spellEnd"/></w:p></w:sdtContent></w:sdt>`;

/** A picture as Word writes one: anchored, cropped, linked, with Word's extensions. */
const PICTURE = p(
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
    'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" ' +
    'wp14:anchorId="1A2B3C4D" wp14:editId="5E6F7A8B"><wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="5" name="Picture 5" descr="A seal">' +
    '<a:hlinkClick r:id="rIdLink"/><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' +
    '<a16:creationId id="{0B1C2D3E-0000-0000-0000-000000000000}"/></a:ext></a:extLst></wp:docPr>' +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="seal.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
    '<pic:blipFill><a:blip r:embed="rIdImage"><a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">' +
    '<a14:useLocalDpi val="0"/></a:ext><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
    '<asvg:svgBlip r:embed="rIdSvg"/></a:ext></a:extLst></a:blip><a:srcRect l="1000" r="2000"/>' +
    '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>' +
    '<a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
);

/** Word's DrawingML copy of a VML shape: base64 in lines of 76. */
const GFXDATA =
  'UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#xA;' +
  '90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA';

/** A text box as Word writes one: a floating shape, and its VML copy beside it (ids as minted). */
const TEXT_BOX = p(
  `<w:r><mc:AlternateContent><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing>` +
    '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
    'relativeHeight="251660288" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1" ' +
    'wp14:anchorId="3209AFDE" wp14:editId="38ACA0DB"><wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="margin"><wp14:pctPosHOffset>40000</wp14:pctPosHOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-182880</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>' +
    '<wp:docPr id="5" name="Text Box 2"/><wp:cNvGraphicFramePr/>' +
    `<a:graphic><a:graphicData uri="${WPS}"><wps:wsp>` +
    '<wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr bwMode="auto">' +
    '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' +
    '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln>' +
    `</wps:spPr><wps:txbx><w:txbxContent>${p(t('Box text'))}</w:txbxContent></wps:txbx>` +
    '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" ' +
    'anchor="t" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>' +
    '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH></wp:anchor>' +
    '</w:drawing></mc:Choice><mc:Fallback><w:pict>' +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
    '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
    '<v:shape id="Text Box 2" o:spid="_x0000_s1026" type="#_x0000_t202" ' +
    'style="position:absolute;width:72pt;height:36pt;z-index:251660288" w14:anchorId="3CEC9D07" ' +
    `o:gfxdata="${GFXDATA}"><v:textbox><w:txbxContent>${p(t('Box text'))}</w:txbxContent></v:textbox>` +
    '<w10:wrap anchorx="margin"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>'
);

/** A group of a shape and a picture, and its VML copy. */
const GROUP = p(
  `<w:r><mc:AlternateContent><mc:Choice xmlns:wpg="${WPG}" Requires="wpg"><w:drawing>` +
    '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
    `<wp:docPr id="5" name="Group 3"/><a:graphic><a:graphicData uri="${WPG}"><wpg:wgp><wpg:cNvGrpSpPr/>` +
    '<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/><a:chOff x="0" y="0"/>' +
    '<a:chExt cx="914400" cy="457200"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="4" name="Oval 4"/>' +
    '<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>' +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:style><a:lnRef idx="2">' +
    '<a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor">' +
    '<a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:bodyPr/></wps:wsp><pic:pic><pic:nvPicPr>' +
    '<pic:cNvPr id="5" name="seal.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>' +
    '<a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>' +
    '<a:xfrm><a:off x="457200" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></wpg:wgp></a:graphicData>' +
    '</a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict>' +
    '<v:group style="width:72pt;height:36pt" coordsize="914400,457200"><v:oval style="width:36pt;height:36pt"/>' +
    '<v:shape type="#_x0000_t75" style="left:457200;width:36pt;height:36pt">' +
    '<v:imagedata r:id="rIdImage" o:title=""/></v:shape></v:group></w:pict></mc:Fallback>' +
    '</mc:AlternateContent></w:r>'
);

/** A VML picture, and a VML shape with a picture fill. */
const VML_PICTURE = p(
  '<w:r><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:36pt;height:36pt">' +
    '<v:imagedata r:id="rIdImage" o:title="seal"/><o:lock v:ext="edit" aspectratio="t"/></v:shape>' +
    '</w:pict></w:r><w:r><w:pict><v:rect style="width:36pt;height:36pt" href="https://example.com/terms">' +
    '<v:fill r:id="rIdSvg" o:title="" type="frame"/></v:rect></w:pict></w:r>'
);

const FRAGMENT =
  HEADING + LINKS + EQUATION + TABLE + BLOCK_CONTROL + PICTURE + TEXT_BOX + GROUP + VML_PICTURE;

describe('a fragment round trip', () => {
  test('blocks come back with the same fingerprint and kinds, without paragraph ids', () => {
    const nodes = read('document', FRAGMENT);
    const expected = read('document', withoutParagraphIds(FRAGMENT));
    const trees = carried(nodes, 'block');
    const back = fragmentNodes(trees, RESOLVE);

    expect(back.map(canonicalOoxmlFingerprint)).toEqual(expected.map(canonicalOoxmlFingerprint));
    expect(kinds(back)).toEqual(kinds(expected));
    expect(kinds(back)).toContain('table:tbl');
    expect(kinds(back)).toContain('contentControl:sdt');
    expect(kinds(back)).toContain('hyperlink:hyperlink');
    expect(kinds(back)).toContain('noteReference:footnoteReference');
    expect(kinds(back)).toContain('contentControlCheckbox:checkbox');
    expect(kinds(back).filter((kind) => kind.endsWith(':drawing'))).toHaveLength(3);
    expect(kinds(back).filter((kind) => kind.endsWith(':pict'))).toHaveLength(4);
  });

  test('the wire names references by index and drops paragraph ids', () => {
    const wire = JSON.stringify(fragmentWire(read('document', FRAGMENT), WIRE));
    expect(wire).not.toContain('paraId');
    expect(wire).not.toContain('textId');
    expect(wire).not.toContain('rIdLink');
    expect(wire).toContain('{"e":"w:numId","a":{"w:val":"#0"}}');
    expect(wire).toContain('{"e":"w:numId","a":{"w:val":"3"}}');
    expect(wire).toContain('{"e":"w:hyperlink","a":{"r:id":"#0","w:history":"1"}');
    expect(wire).toContain('{"e":"w:footnoteReference","a":{"w:id":"#0"}}');
    expect(wire).toContain('{"e":"w:endnoteReference","a":{"w:id":"#1"}}');
    expect(wire).toContain('{"e":"w:t","a":{"xml:space":"preserve"},"c":["  Lead  "]}');
    expect(wire).toContain('{"e":"a:blip","a":{"r:embed":"#0"}');
    expect(wire).toContain('{"e":"a:hlinkClick","a":{"r:id":"#0"}}');
    expect(wire).not.toContain('anchorId');
    expect(wire).not.toContain('rIdImage');
    expect(wire).toContain('{"e":"asvg:svgBlip","a":{"r:embed":"#1"}}');
    expect(wire).toContain('{"e":"mc:Choice","a":{"Requires":"wps"}');
    expect(wire).toContain('{"e":"v:imagedata","a":{"r:id":"#0","o:title":"seal"}}');
    expect(wire).toContain('{"e":"v:fill","a":{"r:id":"#1","o:title":"","type":"frame"}}');
    expect(wire).toContain('"o:gfxdata":"UEsDBBQ');
  });

  test('the change reader returns new trees equal to the wire, with known fields only', () => {
    const wire = fragmentWire(read('document', FRAGMENT), WIRE);
    const json = JSON.parse(JSON.stringify(wire)) as OrderedFragmentElement[];
    const checked = fragmentElements(json, 'blocks', 'block', REFERENCES, fragmentBudget());
    expect(checked).toEqual(wire);
    expect(checked[0]).not.toBe(json[0]);
    expect(checked[0]?.a).not.toBe(json[0]?.a);
  });

  test('list definitions, overrides and styles come back as resources', () => {
    const numbering = read(
      'numbering',
      '<w:abstractNum w:abstractNumId="4" w15:restartNumberingAfterBreak="0">' +
        '<w:nsid w:val="1C2D3E4F"/><w:multiLevelType w:val="hybridMultilevel"/><w:tmpl w:val="0409000F"/>' +
        '<w:lvl w:ilvl="0" w:tplc="0409000F"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
        '<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>' +
        '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>' +
        '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' +
        '<w:num w:numId="7"><w:abstractNumId w:val="4"/>' +
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>'
    );
    const abstract = numbering.filter((node) => node.localName === 'abstractNum');
    const num = numbering.find((node) => node.localName === 'num');
    const overrides = (num?.children ?? []).filter(
      (node): node is OoxmlElement => node.kind !== 'textValue' && node.localName === 'lvlOverride'
    );
    const styles = read(
      'styles',
      '<w:style w:type="paragraph" w:customStyle="1" w:styleId="ListPara"><w:name w:val="List Para"/>' +
        '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr>' +
        '<w:ind w:left="720"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>'
    );
    for (const [nodes, root] of [
      [abstract, 'abstractNum'],
      [overrides, 'lvlOverride'],
      [styles, 'style'],
    ] as const) {
      const back = fragmentResourceNodes(carried(nodes, root), root, RESOLVE);
      expect(back.map(canonicalOoxmlFingerprint)).toEqual(nodes.map(canonicalOoxmlFingerprint));
      expect(kinds(back)).toEqual(kinds(nodes));
    }
  });
});

describe('a fragment value that holds markup', () => {
  const MARKUP = 'a"<>&]]>\'b</w:t></w:r><w:r><w:t>x';
  const ATTRIBUTE = '"/><w:b/><w:i a="&amp;]]>';

  test('stays text in an attribute and in a text, with no new element', () => {
    const json: unknown = JSON.parse(
      JSON.stringify([
        {
          e: 'w:p',
          c: [
            { e: 'w:pPr', c: [{ e: 'w:pStyle', a: { 'w:val': ATTRIBUTE } }] },
            { e: 'w:r', c: [{ e: 'w:t', c: [MARKUP] }] },
          ],
        },
      ])
    );
    const trees = fragmentElements(json, 'blocks', 'block', REFERENCES, fragmentBudget());
    const [paragraph] = fragmentNodes(trees, RESOLVE);

    expect(kinds(paragraph ? [paragraph] : [])).toEqual([
      'paragraph:p',
      'paragraphProperties:pPr',
      'generic:pStyle',
      'run:r',
      'text:t',
      'text',
    ]);
    const properties = paragraph?.children[0];
    const style = properties?.kind === 'textValue' ? undefined : properties?.children[0];
    expect(style?.kind === 'textValue' ? undefined : style?.attributes[0]?.value).toBe(ATTRIBUTE);
    const run = paragraph?.children[1];
    const text = run?.kind === 'textValue' ? undefined : run?.children[0];
    const value = text?.kind === 'textValue' ? undefined : text?.children[0];
    expect(value?.kind === 'textValue' ? value.value : undefined).toBe(MARKUP);
  });

  test('round trips through the wire unchanged', () => {
    const nodes = read(
      'document',
      p(t('a&quot;&lt;&gt;&amp;]]&gt;b'), '<w:pStyle w:val="&quot;/&gt;&lt;w:b/&gt;&amp;]]&gt;"/>')
    );
    const back = fragmentNodes(carried(nodes, 'block'), RESOLVE);
    expect(back.map(canonicalOoxmlFingerprint)).toEqual(nodes.map(canonicalOoxmlFingerprint));
    expect(JSON.stringify(fragmentWire(nodes, WIRE))).toContain('"c":["a\\"<>&]]>b"]');
  });
});
