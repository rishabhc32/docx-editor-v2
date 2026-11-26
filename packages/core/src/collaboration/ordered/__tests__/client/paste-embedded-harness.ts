// The sources of the embedded object and ink paste tests of group 9.10: a line with an embedded
// workbook (`w:object`: its preview picture and its `o:OLEObject` part), and two lines with ink
// (`w14:contentPart` and its InkML part), with the parts and relationships Word writes.

import { strToU8, zipSync } from 'fflate';
import { docx, p, R, RED_PNG, t } from '../word-fixtures.ts';
import { copiedBetween, type Copied } from './paste-harness.ts';

const V = 'urn:schemas-microsoft-com:vml';
const O = 'urn:schemas-microsoft-com:office:office';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WPI = 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/**
 * An embedded workbook: a zip that nothing opens. Its time is fixed, so its bytes and their hash
 * are the same in each run (a version 5 case names the hash).
 */
export const EMBEDDED_WORKBOOK: Uint8Array = zipSync(
  { 'xl/workbook.xml': strToU8('<workbook/>') },
  { mtime: Date.UTC(2026, 0, 1) }
);

/** An InkML part: XML text that nothing reads. */
export const INK: Uint8Array = strToU8(
  '<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"><inkml:trace>0 0, 10 10</inkml:trace></inkml:ink>'
);

const OBJECT_RUN =
  `<w:r><w:object xmlns:v="${V}" xmlns:o="${O}" w:dxaOrig="1440" w:dyaOrig="720">` +
  '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:72pt;height:36pt" o:ole="">' +
  '<v:imagedata r:id="rIdPreview" o:title=""/></v:shape>' +
  '<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Content" ' +
  'ObjectID="_1234567890" r:id="rIdSheet"/></w:object></w:r>';

const INK_RUN =
  `<w:r><mc:AlternateContent xmlns:mc="${MC}" xmlns:w14="${W14}"><mc:Choice Requires="w14">` +
  '<w14:contentPart r:id="rIdInk"/></mc:Choice><mc:Fallback><w:t>x</w:t></mc:Fallback>' +
  '</mc:AlternateContent></w:r>';

// Ink in the fuller form of Word 2010 ([MS-DOCX] `CT_WordContentPart`): a choice that requires
// `wpi`, and the content part's DrawingML-typed elements, whose attributes have no prefix
const INK_WPI_RUN =
  `<w:r><mc:AlternateContent xmlns:mc="${MC}" xmlns:w14="${W14}" xmlns:wpi="${WPI}" ` +
  `xmlns:a="${A}"><mc:Choice Requires="wpi"><w14:contentPart bwMode="auto" r:id="rIdInk2">` +
  '<w14:nvContentPartPr><w14:cNvContentPartPr/></w14:nvContentPartPr><w14:xfrm>' +
  '<a:off x="0" y="0"/><a:ext cx="360" cy="360"/></w14:xfrm></w14:contentPart></mc:Choice>' +
  '<mc:Fallback><w:t>x</w:t></mc:Fallback></mc:AlternateContent></w:r>';

/** A line with an embedded workbook, a line with ink, and a line with ink in the fuller form. */
export const EMBEDDED_SOURCE: Uint8Array = docx(
  p(t('Sheet ') + OBJECT_RUN + t(' here.')) +
    p(t('Ink ') + INK_RUN + t(' here.')) +
    p(t('Pen ') + INK_WPI_RUN + t(' here.')),
  {
    parts: {
      'word/media/preview.png': RED_PNG,
      'word/embeddings/Microsoft_Excel_Worksheet.xlsx': EMBEDDED_WORKBOOK,
      'word/ink/ink1.xml': INK,
      'word/ink/ink2.xml': INK,
    },
    contentTypes:
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>' +
      '<Override PartName="/word/ink/ink1.xml" ContentType="application/inkml+xml"/>' +
      '<Override PartName="/word/ink/ink2.xml" ContentType="application/inkml+xml"/>',
    relationships:
      `<Relationship Id="rIdPreview" Type="${R}/image" Target="media/preview.png"/>` +
      `<Relationship Id="rIdSheet" Type="${R}/package" Target="embeddings/Microsoft_Excel_Worksheet.xlsx"/>` +
      '<Relationship Id="rIdInk" Type="http://schemas.microsoft.com/office/2011/relationships/ink" Target="ink/ink1.xml"/>' +
      '<Relationship Id="rIdInk2" Type="http://schemas.microsoft.com/office/2011/relationships/ink" Target="ink/ink2.xml"/>',
  }
);

/** What each line holds, and its text. */
export const EMBEDDED = {
  object: { line: 0, text: 'Sheet ￼ here.', shows: /<o:OLEObject [^>]*r:id="[^"]+"/ },
  ink: { line: 1, text: 'Ink ￼ here.', shows: /<w14:contentPart r:id="[^"]+"/ },
  inkWpi: {
    line: 2,
    text: 'Pen ￼ here.',
    shows: /<w14:contentPart bwMode="auto" r:id="[^"]+"><w14:nvContentPartPr>/,
  },
} as const;

/** A line, without its paragraph mark, as the editor copies it. */
export const copiedEmbedded = (kind: keyof typeof EMBEDDED): Copied =>
  copiedBetween(EMBEDDED[kind].line, EMBEDDED[kind].line, undefined, EMBEDDED_SOURCE);
