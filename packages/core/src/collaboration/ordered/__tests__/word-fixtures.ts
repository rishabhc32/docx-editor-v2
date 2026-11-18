// Word packages for the ordered model's tests: a small `.docx` builder, and the matrix of item
// kinds Word puts in a story (from spike 5's `matrix-items.ts`). An inline item sits in
// paragraph 1 as "Alpha <item> Omega.", between "Before the item." and "After the item."; a
// block item is its own paragraph or table. Every item shares one package: media, numbering,
// both notes parts and comments are always there.

import { strToU8, zipSync } from 'fflate';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';

export const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
export const V = 'urn:schemas-microsoft-com:vml';
export const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const WML_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/** Parts a fixture adds to the package. */
export interface PackageExtras {
  /** Extra parts by path (`word/numbering.xml`): XML as text, media as bytes. */
  readonly parts?: Readonly<Record<string, string | Uint8Array>>;
  /** Extra `Override` and `Default` entries for `[Content_Types].xml`. */
  readonly contentTypes?: string;
  /** Extra `Relationship` entries for the document part. */
  readonly relationships?: string;
  /** Children of `w:document` before `w:body` (a page background). */
  readonly documentPrefix?: string;
}

const NAMESPACES =
  `xmlns:w="${W}" xmlns:w14="${W14}" xmlns:r="${R}" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

const STYLES =
  `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Hyperlink">` +
  '<w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>' +
  '</w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/>' +
  '<w:rPr><w:b/><w:bCs/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>' +
  '<w:pPr><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>' +
  '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders>' +
  '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '</w:tblBorders></w:tblPr></w:style></w:styles>';

/** A `.docx` whose body holds `body`. */
export function docx(body: string, extras: PackageExtras = {}): Uint8Array {
  const parts: Record<string, Uint8Array> = {};
  for (const [path, content] of Object.entries(extras.parts ?? {}))
    parts[path] = typeof content === 'string' ? strToU8(content) : content;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${WML_CT}.document.main+xml"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${WML_CT}.styles+xml"/>` +
        `${extras.contentTypes ?? ''}</Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdLink" Type="${R}/hyperlink" Target="https://example.com/old" TargetMode="External"/>` +
        `${extras.relationships ?? ''}</Relationships>`
    ),
    'word/styles.xml': strToU8(STYLES),
    'word/document.xml': strToU8(
      `<w:document ${NAMESPACES}>${extras.documentPrefix ?? ''}<w:body>${body}</w:body></w:document>`
    ),
    ...parts,
  });
}

/** The package of a `.docx`, or a thrown reason. */
export function packageOf(bytes: Uint8Array): OoxmlPackage {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(`unreadable fixture: ${read.reason}`);
  return read.package;
}

export const t = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
export const p = (runs: string, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs}</w:p>`;

/** A one-pixel red PNG. */
export const RED_PNG = new Uint8Array([
  137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0,
  0, 144, 119, 83, 222, 0, 0, 0, 12, 73, 68, 65, 84, 120, 156, 99, 184, 35, 39, 7, 0, 2, 242, 1, 25,
  115, 68, 11, 203, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
]);

/** A red square as an SVG document. */
export const RED_SVG = new TextEncoder().encode(
  '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">' +
    '<rect width="16" height="16" fill="red"/></svg>'
);

/** An enhanced metafile: its header record ([MS-EMF] 2.3.4.2) and its end record. */
export function emfBytes(): Uint8Array {
  const bytes = new Uint8Array(108);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 1, true);
  view.setUint32(4, 88, true);
  view.setInt32(16, 31, true);
  view.setInt32(20, 31, true);
  view.setInt32(32, 846, true);
  view.setInt32(36, 846, true);
  view.setUint32(40, 0x464d4520, true);
  view.setUint32(44, 0x10000, true);
  view.setUint32(48, 108, true);
  view.setUint32(52, 2, true);
  view.setUint16(56, 1, true);
  view.setInt32(72, 1920, true);
  view.setInt32(76, 1080, true);
  view.setInt32(80, 508, true);
  view.setInt32(84, 286, true);
  view.setUint32(88, 14, true);
  view.setUint32(92, 20, true);
  view.setUint32(100, 16, true);
  view.setUint32(104, 20, true);
  return bytes;
}

/**
 * A Windows metafile with a placeable header ([MS-WMF] 2.3.2.3), its standard header and its end
 * record; without the placeable header when `placeable` is false.
 */
export function wmfBytes(placeable = true): Uint8Array {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, 0x9ac6cdd7, true);
  view.setInt16(10, 1440, true);
  view.setInt16(12, 1440, true);
  view.setUint16(14, 1440, true);
  let sum = 0;
  for (let k = 0; k < 20; k += 2) sum ^= view.getUint16(k, true);
  view.setUint16(20, sum, true);
  view.setUint16(22, 1, true);
  view.setUint16(24, 9, true);
  view.setUint16(26, 0x0300, true);
  view.setUint32(28, 12, true);
  view.setUint32(34, 3, true);
  view.setUint32(40, 3, true);
  return placeable ? bytes : bytes.slice(22);
}

export const pictureRun = (id: number): string =>
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  `<wp:extent cx="457200" cy="457200"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="red.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  '<pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

const separators = (kind: 'footnote' | 'endnote') =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

const level = (ilvl: number, format: string, text: string, left: number) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/>` +
  `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;

/**
 * More list definitions: bullets (2), three levels (3), roman (4), a restart of list 1 at 5
 * (5), a custom number text (6), and a picture bullet (7).
 */
const MORE_LISTS =
  '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>' +
  '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/>' +
  level(0, 'decimal', '%1.', 720) +
  level(1, 'lowerLetter', '%2)', 1440) +
  level(2, 'lowerRoman', '%3.', 2160) +
  '</w:abstractNum>' +
  `<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="singleLevel"/>${level(0, 'upperRoman', '%1.', 720)}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="singleLevel"/>${level(0, 'decimal', 'Article %1.', 1080)}</w:abstractNum>` +
  '<w:abstractNum w:abstractNumId="5"><w:multiLevelType w:val="singleLevel"/>' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>' +
  '<w:lvlPicBulletId w:val="0"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>';
const MORE_NUMS =
  '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>' +
  '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>' +
  '<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>' +
  '<w:num w:numId="5"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>' +
  '<w:num w:numId="6"><w:abstractNumId w:val="4"/></w:num>' +
  '<w:num w:numId="7"><w:abstractNumId w:val="5"/></w:num>';

/** One package for every item: the picture's media, numbering, both notes parts, comments. */
export const MATRIX_EXTRAS: PackageExtras = {
  parts: {
    'word/media/red.png': RED_PNG,
    'word/footnotes.xml':
      `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${separators('footnote')}` +
      `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r>${t(' Fees are due net of tax.')}</w:p></w:footnote></w:footnotes>`,
    'word/endnotes.xml':
      `<w:endnotes xmlns:w="${W}" xmlns:r="${R}">${separators('endnote')}` +
      `<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r>${t(' See the annex.')}</w:p></w:endnote></w:endnotes>`,
    'word/comments.xml':
      `<w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Cara" w:date="2026-01-01T00:00:00Z" w:initials="C">` +
      `${p(t('Check this.'))}</w:comment></w:comments>`,
    'word/numbering.xml':
      `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
      '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
      '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
      MORE_LISTS +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      MORE_NUMS +
      '</w:numbering>',
  },
  contentTypes:
    '<Default Extension="png" ContentType="image/png"/>' +
    `<Override PartName="/word/footnotes.xml" ContentType="${WML_CT}.footnotes+xml"/>` +
    `<Override PartName="/word/endnotes.xml" ContentType="${WML_CT}.endnotes+xml"/>` +
    `<Override PartName="/word/comments.xml" ContentType="${WML_CT}.comments+xml"/>` +
    `<Override PartName="/word/numbering.xml" ContentType="${WML_CT}.numbering+xml"/>`,
  relationships:
    `<Relationship Id="rIdImg1" Type="${R}/image" Target="media/red.png"/>` +
    `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
    `<Relationship Id="rIdEndnotes" Type="${R}/endnotes" Target="endnotes.xml"/>` +
    `<Relationship Id="rIdComments" Type="${R}/comments" Target="comments.xml"/>` +
    `<Relationship Id="rIdNumbering" Type="${R}/numbering" Target="numbering.xml"/>`,
};

/** One kind of item: inline XML between "Alpha " and " Omega.", or blocks of its own. */
export interface MatrixItem {
  readonly key: string;
  readonly name: string;
  readonly xml?: string;
  readonly blocks?: string;
}

const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const field = (instr: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  result +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
const anchoredPicture =
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
  'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="457200" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="7" name="Picture 7"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="7" name="red.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
const textBox =
  `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing>` +
  '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="8" name="Text Box 8"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
  '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${p(t('Box text'))}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>` +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
  `<mc:Fallback><w:pict><v:shape xmlns:v="${V}" style="width:72pt;height:36pt">` +
  `<v:textbox><w:txbxContent>${p(t('Box text'))}</w:txbxContent></v:textbox></v:shape></w:pict>` +
  '</mc:Fallback></mc:AlternateContent></w:r>';
const vmlBox =
  `<w:r><w:pict><v:shape xmlns:v="${V}" style="width:72pt;height:36pt">` +
  `<v:textbox><w:txbxContent>${p(t('Old box'))}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
const cell = (content: string, tcPr = '<w:tcW w:w="3000" w:type="dxa"/>') =>
  `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;
const table = (cols: number, rows: string) =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(cols)}</w:tblGrid>${rows}</w:tbl>`;
export const alphaOmega = (xml: string): string => t('Alpha ') + xml + t(' Omega.');
const AO = p(alphaOmega(t('middle')));
const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';

const inline = (key: string, name: string, xml: string): MatrixItem => ({ key, name, xml });
const block = (key: string, name: string, blocks: string): MatrixItem => ({ key, name, blocks });

/** Every kind of item Word puts in a story: the spike's 65 kinds, its two emoji cases as one. */
export const MATRIX_ITEMS: readonly MatrixItem[] = [
  inline('bold', 'bold text', t('Bold', '<w:b/>')),
  inline(
    'face',
    'color, size and highlight',
    t('Red', '<w:color w:val="FF0000"/><w:sz w:val="28"/><w:highlight w:val="yellow"/>')
  ),
  inline('styled', 'character style', t('Strong', '<w:rStyle w:val="Hyperlink"/>')),
  inline('hidden', 'hidden text', t('Hidden', '<w:vanish/>')),
  inline('emoji', 'emoji', t('Doc \u{1F4C4} here')),
  inline(
    'glow',
    'text effect (w14)',
    t('Glow', '<w14:glow w14:rad="63500"><w14:schemeClr w14:val="accent1"/></w14:glow>')
  ),
  inline('rtl', 'right-to-left run', t('שלום', '<w:rtl/>')),
  inline('tab', 'tab', '<w:r><w:tab/></w:r>'),
  inline('br', 'line break', '<w:r><w:br/></w:r>'),
  inline('page', 'page break', '<w:r><w:br w:type="page"/></w:r>'),
  inline('column', 'column break', '<w:r><w:br w:type="column"/></w:r>'),
  inline('cr', 'carriage return', '<w:r><w:cr/></w:r>'),
  inline('picture', 'inline picture', pictureRun(1)),
  inline('anchor', 'floating picture', anchoredPicture),
  inline('textbox', 'text box (wps)', textBox),
  inline('vml', 'text box (VML)', vmlBox),
  inline('sym', 'symbol', '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>'),
  inline('soft', 'soft hyphen', '<w:r><w:t>co</w:t><w:softHyphen/><w:t>op</w:t></w:r>'),
  inline('nbh', 'non-breaking hyphen', '<w:r><w:t>X</w:t><w:noBreakHyphen/><w:t>ray</w:t></w:r>'),
  inline(
    'ptab',
    'positional tab',
    '<w:r><w:ptab w:alignment="right" w:leader="dot" w:relativeTo="margin"/></w:r>'
  ),
  inline(
    'fnref',
    'footnote reference',
    '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'
  ),
  inline('enref', 'endnote reference', '<w:r><w:endnoteReference w:id="1"/></w:r>'),
  inline('math', 'equation', `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>`),
  inline(
    'ruby',
    'phonetic guide (ruby)',
    '<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="center"/><w:hps w:val="10"/><w:hpsRaise w:val="18"/>' +
      `<w:hpsBaseText w:val="20"/><w:lid w:val="ja-JP"/></w:rubyPr><w:rt>${t('kan')}</w:rt><w:rubyBase>${t('漢')}</w:rubyBase></w:ruby></w:r>`
  ),
  inline(
    'bookmark',
    'bookmark',
    `<w:bookmarkStart w:id="0" w:name="mark"/>${t('marked')}<w:bookmarkEnd w:id="0"/>`
  ),
  inline(
    'emptymark',
    'empty bookmark',
    '<w:bookmarkStart w:id="1" w:name="here"/><w:bookmarkEnd w:id="1"/>'
  ),
  inline(
    'comment',
    'comment',
    `<w:commentRangeStart w:id="0"/>${t('noted')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`
  ),
  inline(
    'perm',
    'editing permission range',
    `<w:permStart w:id="5" w:edGrp="everyone"/>${t('open')}<w:permEnd w:id="5"/>`
  ),
  inline(
    'proof',
    'proofing marks',
    `<w:proofErr w:type="spellStart"/>${t('teh')}<w:proofErr w:type="spellEnd"/>`
  ),
  inline('lrpb', 'rendered page break', '<w:r><w:lastRenderedPageBreak/><w:t>next</w:t></w:r>'),
  inline(
    'link',
    'hyperlink (external)',
    `<w:hyperlink r:id="rIdLink">${t('Docs', '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>`
  ),
  inline(
    'anchorlink',
    'hyperlink (to a bookmark)',
    `<w:hyperlink w:anchor="mark">${t('Jump')}</w:hyperlink>`
  ),
  inline(
    'smart',
    'smart tag',
    `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">${t('Paris')}</w:smartTag>`
  ),
  inline('xml', 'custom XML (inline)', `<w:customXml w:element="party">${t('Acme')}</w:customXml>`),
  inline('dir', 'bidirectional override', `<w:dir w:val="rtl">${t('abc')}</w:dir>`),
  inline(
    'plain',
    'plain text control',
    sdt('<w:alias w:val="Name"/><w:tag w:val="name"/><w:text/>', t('Name'))
  ),
  inline(
    'rich',
    'rich text control',
    sdt('<w:alias w:val="Rich"/><w:tag w:val="rich"/>', t('Rich'))
  ),
  inline(
    'placeholder',
    'control showing its placeholder',
    sdt('<w:showingPlcHdr/><w:text/>', t('Click here', '<w:rStyle w:val="PlaceholderText"/>'))
  ),
  inline(
    'dropdown',
    'drop-down list',
    sdt(
      '<w:dropDownList w:lastValue="1"><w:listItem w:displayText="One" w:value="1"/><w:listItem w:displayText="Two" w:value="2"/></w:dropDownList>',
      t('One')
    )
  ),
  inline(
    'combo',
    'combo box',
    sdt(
      '<w:comboBox><w:listItem w:displayText="Red" w:value="r"/><w:listItem w:displayText="Blue" w:value="b"/></w:comboBox>',
      t('Blue')
    )
  ),
  inline(
    'date',
    'date picker',
    sdt(
      '<w:date w:fullDate="2026-01-02T00:00:00Z"><w:dateFormat w:val="M/d/yyyy"/></w:date>',
      t('1/2/2026')
    )
  ),
  inline(
    'checkbox',
    'check box',
    sdt(
      '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
        '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>',
      t('☐', '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>')
    )
  ),
  inline('pagefield', 'PAGE field', field('PAGE', t('1'))),
  inline(
    'ref',
    'REF field (two runs of result)',
    field('REF mark \\h', t('see ') + t('marked', '<w:b/>'))
  ),
  inline('simple', 'simple field', `<w:fldSimple w:instr=" DATE ">${t('1/2/2026')}</w:fldSimple>`),
  inline('ins', 'insertion from Word', `<w:ins w:id="901" ${REV}>${t('earlier')}</w:ins>`),
  inline(
    'del',
    'deletion from Word',
    `<w:del w:id="902" ${REV}><w:r><w:delText>former</w:delText></w:r></w:del>`
  ),
  inline(
    'fmt',
    'format change from Word',
    `<w:r><w:rPr><w:b/><w:rPrChange w:id="903" ${REV}><w:rPr/></w:rPrChange></w:rPr><w:t>heavy</w:t></w:r>`
  ),
  block(
    'list',
    'list paragraph',
    p(alphaOmega(t('middle')), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
  ),
  block('heading', 'heading style', p(alphaOmega(t('middle')), '<w:pStyle w:val="Heading1"/>')),
  block(
    'formatted',
    'borders, shading, spacing, indents, tab stops',
    p(
      alphaOmega(t('middle')),
      '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr>' +
        '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>' +
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8000"/></w:tabs>' +
        '<w:spacing w:before="120" w:after="120"/><w:ind w:left="720" w:firstLine="360"/><w:jc w:val="both"/>'
    )
  ),
  block(
    'markfmt',
    'formatted paragraph mark',
    p(alphaOmega(t('middle')), '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr>')
  ),
  block(
    'frame',
    'text frame',
    p(
      alphaOmega(t('middle')),
      '<w:framePr w:w="3000" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:x="1000" w:y="1"/>'
    )
  ),
  block(
    'section',
    'section break',
    p(
      alphaOmega(t('middle')),
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr>'
    )
  ),
  block(
    'mathpara',
    'equation paragraph',
    `<w:p>${t('Alpha ')}<m:oMathPara xmlns:m="${M}"><m:oMath><m:r><m:t>y=2</m:t></m:r></m:oMath></m:oMathPara>${t(' Omega.')}</w:p>`
  ),
  block('sdtblock', 'block control', sdt('<w:alias w:val="Clause"/><w:tag w:val="clause"/>', AO)),
  block(
    'blockprompt',
    'block control showing its prompt',
    sdt(
      '<w:alias w:val="Clause"/><w:showingPlcHdr/>',
      p(
        alphaOmega(t('middle')).replace(
          /<w:r>/g,
          '<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>'
        )
      )
    )
  ),
  block('xmlblock', 'custom XML block', `<w:customXml w:element="clause">${AO}</w:customXml>`),
  block('empty', 'empty paragraph before', `<w:p/>${AO}`),
  block(
    'table',
    'table cell',
    table(
      2,
      `<w:tr>${cell(AO)}${cell(p(t('Fee')))}</w:tr><w:tr>${cell(p(t('Due')))}${cell(p(t('30 days')))}</w:tr>`
    )
  ),
  block(
    'span',
    'merged cell (across columns)',
    table(
      2,
      `<w:tr>${cell(AO, '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>')}</w:tr><w:tr>${cell(p(t('Due')))}${cell(p(t('30 days')))}</w:tr>`
    )
  ),
  block(
    'vmerge',
    'merged cell (down rows)',
    table(
      2,
      `<w:tr>${cell(AO, '<w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/>')}${cell(p(t('Fee')))}</w:tr>` +
        `<w:tr>${cell('<w:p/>', '<w:tcW w:w="3000" w:type="dxa"/><w:vMerge/>')}${cell(p(t('30 days')))}</w:tr>`
    )
  ),
  block(
    'nested',
    'nested table',
    table(
      1,
      `<w:tr>${cell(p(t('Outer')) + table(1, `<w:tr>${cell(AO)}</w:tr>`) + p(t('Outer end')))}</w:tr>`
    )
  ),
  block(
    'headrow',
    'repeating header row',
    table(
      1,
      `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(AO)}</w:tr><w:tr>${cell(p(t('Row two')))}</w:tr>`
    )
  ),
];

/** The body of an item's fixture. */
export function bodyOf(item: MatrixItem): string {
  const between = item.blocks ?? p(alphaOmega(item.xml ?? ''));
  return p(t('Before the item.')) + between + p(t('After the item.'));
}
