// More item kinds, past spike 5's 65: every kind of Word content the editor lists in
// `docs/site/data/word-features.ts` that the matrix can put in a story, so each one meets every
// edit from several authors. Story kinds sit in a header, a footer or a note, not the body.

import {
  alphaOmega,
  M,
  MATRIX_EXTRAS,
  MC,
  p,
  pictureRun,
  R,
  t,
  V,
  W,
  WML_CT,
  WPS,
  type PackageExtras,
} from '../word-fixtures.ts';
import type { Group, MatrixKind, Selector } from './catalogue.ts';

const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';
const at = (find: string, delta: number): Selector => ({ find, edge: 'start', delta });
const AO = p(alphaOmega(t('middle')));
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const PROMPT = '<w:rStyle w:val="PlaceholderText"/>';
const field = (instr: string, result: string, begin = '<w:fldChar w:fldCharType="begin"/>') =>
  `<w:r>${begin}</w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const cell = (content: string, tcPr = '<w:tcW w:w="3000" w:type="dxa"/>') =>
  `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;
const table = (cols: number, rows: string, tblPr = '<w:tblW w:w="0" w:type="auto"/>') =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>${'<w:gridCol w:w="3000"/>'.repeat(cols)}</w:tblGrid>${rows}</w:tbl>`;
const list = (numId: number, ilvl = 0) =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
const sectPr = (more: string) =>
  `<w:sectPr>${more}<w:pgSz w:w="12240" w:h="15840"/>` +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

/** A drawing in a run: `inner` is the `wp:inline` or `wp:anchor` body after its extent. */
const drawing = (id: number, graphic: string, anchor?: string) => {
  const body =
    `<wp:extent cx="457200" cy="457200"/><wp:docPr id="${id}" name="Object ${id}"/>` +
    `<a:graphic><a:graphicData ${graphic}</a:graphicData></a:graphic>`;
  return anchor === undefined
    ? `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">${body}</wp:inline></w:drawing></w:r>`
    : '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" ' +
        `relativeHeight="251659264" ${anchor} locked="0" layoutInCell="1" allowOverlap="1">` +
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
        `${body.replace('<wp:docPr', '<wp:effectExtent l="0" t="0" r="0" b="0"/>__WRAP__<wp:docPr')}` +
        '</wp:anchor></w:drawing></w:r>';
};
const PIC = 'uri="http://schemas.openxmlformats.org/drawingml/2006/picture">';
const picture = (id: number, blip = '<a:blip r:embed="rIdImg1"/>', more = '') =>
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="red.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill>${blip}${more}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>';
const floating = (id: number, wrap: string, anchor = 'behindDoc="0"') =>
  drawing(id, `${PIC}${picture(id)}`, anchor).replace('__WRAP__', wrap);

const inline = (
  key: string,
  name: string,
  group: Group,
  xml: string,
  text: string,
  more: Partial<MatrixKind> = {}
): MatrixKind => ({ key, name, group, xml, text, ...more });
const block = (
  key: string,
  name: string,
  group: Group,
  blocks: string,
  more: Partial<MatrixKind> = {}
): MatrixKind => ({ key, name, group, blocks, text: 'middle', ...more });
const paragraph = (key: string, name: string, pPr: string, more: Partial<MatrixKind> = {}) =>
  block(key, name, 'block', p(alphaOmega(t('middle')), pPr), more);

// ---- story parts: a header, a footer, a note ----

const NAMESPACES =
  `xmlns:w="${W}" xmlns:r="${R}" xmlns:v="${V}" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
const STORY = (xml: string) =>
  p(t('Before the item.')) + p(alphaOmega(xml)) + p(t('After the item.'));

/** Which story a kind's edits go to, and how the editor opens it. */
export interface StoryTarget {
  readonly part: string;
  /** A text box: the one text box of the main document, a story of its own in the model. */
  readonly scope: 'header' | 'footer' | 'note' | 'textbox';
  /** The header's or footer's relationship id, the note's scope id (`footnote:2`), or ''. */
  readonly id: string;
  /** Paragraphs the editor's scope leaves out at the start (a notes part's two separators). */
  readonly skip?: number;
  /** The page whose header or footer the editor opens (an even-page header shows on page 2). */
  readonly page?: number;
}

export function furniture(
  key: string,
  name: string,
  kind: 'header' | 'footer',
  xml: string,
  text: string,
  type: 'default' | 'first' | 'even' = 'default'
): MatrixKind {
  const root = kind === 'header' ? 'hdr' : 'ftr';
  const file = `${kind}9.xml`;
  const rId = `rId${kind}9`;
  // Even-page headers are used only when the settings turn them on
  const even: PackageExtras =
    type === 'even'
      ? {
          parts: {
            'word/settings.xml': `<w:settings xmlns:w="${W}"><w:evenAndOddHeaders/></w:settings>`,
          },
          contentTypes: `<Override PartName="/word/settings.xml" ContentType="${WML_CT}.settings+xml"/>`,
          relationships: `<Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>`,
        }
      : {};
  const extras: PackageExtras = {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      ...even.parts,
      [`word/${file}`]: `<w:${root} ${NAMESPACES}>${STORY(xml)}</w:${root}>`,
      [`word/_rels/${file}.rels`]:
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rIdImg1" Type="${R}/image" Target="media/red.png"/></Relationships>`,
    },
    contentTypes: `${MATRIX_EXTRAS.contentTypes ?? ''}${even.contentTypes ?? ''}<Override PartName="/word/${file}" ContentType="${WML_CT}.${kind}+xml"/>`,
    relationships: `${MATRIX_EXTRAS.relationships ?? ''}${even.relationships ?? ''}<Relationship Id="${rId}" Type="${R}/${kind}" Target="${file}"/>`,
  };
  const reference = `<w:${kind}Reference w:type="${type}" r:id="${rId}"/>`;
  return {
    key,
    name,
    group: 'story',
    // An even-page header shows on the second page
    body:
      p(t('Body text.')) +
      (type === 'even' ? p('<w:r><w:br w:type="page"/></w:r>' + t('Page two.')) : '') +
      sectPr(`${reference}${type === 'first' ? '<w:titlePg/>' : ''}`),
    text,
    extras,
    story: { part: `/word/${file}`, scope: kind, id: rId, ...(type === 'even' ? { page: 1 } : {}) },
    p: 1,
  };
}

export function note(
  key: string,
  name: string,
  kind: 'footnote' | 'endnote',
  xml: string
): MatrixKind {
  const parts = MATRIX_EXTRAS.parts ?? {};
  const file = `word/${kind}s.xml`;
  const existing = parts[file];
  if (typeof existing !== 'string') throw new Error(`no ${file}`);
  const ref = kind === 'footnote' ? 'footnoteRef' : 'endnoteRef';
  const added =
    `<w:${kind} w:id="2"><w:p><w:r><w:${ref}/></w:r>${t('Before the item.')}</w:p>` +
    `${p(alphaOmega(xml))}${p(t('After the item.'))}</w:${kind}>`;
  return {
    key,
    name,
    group: 'story',
    body: p(t('Body text.') + `<w:r><w:${kind}Reference w:id="2"/></w:r>`),
    text: 'middle',
    extras: {
      ...MATRIX_EXTRAS,
      parts: { ...parts, [file]: existing.replace(`</w:${kind}s>`, `${added}</w:${kind}s>`) },
    },
    story: { part: `/${file}`, scope: 'note', id: `${kind}:2`, skip: 2 },
    // The separators' paragraphs, the first note's, then the item's note
    p: 4,
  };
}

/**
 * A text box in the body whose content is the story (Word's DrawingML text box with its VML
 * fallback, as `word-fixtures.ts` writes one), so every edit goes inside the text box. The
 * `textbox` and `vml` kinds are the object as one unit in the body; these are its text.
 */
export function textBoxText(key: string, name: string, xml: string, text: string): MatrixKind {
  const content = `<w:txbxContent>${STORY(xml)}</w:txbxContent>`;
  const box =
    `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing>` +
    '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="914400"/>' +
    `<wp:docPr id="9" name="Text Box 9"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr txBox="1"/>` +
    '<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
    `<wps:txbx>${content}</wps:txbx><wps:bodyPr/></wps:wsp>` +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>' +
    `<mc:Fallback><w:pict><v:shape xmlns:v="${V}" style="width:216pt;height:72pt">` +
    `<v:textbox>${content}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
  return {
    key,
    name,
    group: 'story',
    body: p(t('Body text.') + box),
    text,
    story: { part: '/word/document.xml', scope: 'textbox', id: '' },
    p: 1,
  };
}

// ---- the kinds ----

const watermark =
  `<w:r><w:pict><v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s1" xmlns:o="urn:schemas-microsoft-com:office:office" type="#_x0000_t136" ` +
  'style="position:absolute;width:400pt;height:100pt;rotation:315" fillcolor="silver" stroked="f">' +
  '<v:textpath style="font-family:&quot;Calibri&quot;" string="DRAFT"/></v:shape></w:pict></w:r>';

const toc = sdt(
  '<w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj>',
  `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${alphaOmega(t('middle'))}</w:p>` +
    `<w:p>${t('Second entry')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`
);

/** Kinds past spike 5's 65, grouped as the feature list groups them. */
export { cell, list, sdt, table };

export const MORE_KINDS: readonly MatrixKind[] = [
  // Run formatting
  inline('italic', 'italic', 'text', t('Slanted', '<w:i/>'), 'Slanted', {
    inside: at('Slanted', 3),
  }),
  inline('udouble', 'double underline', 'text', t('Twice', '<w:u w:val="double"/>'), 'Twice', {
    inside: at('Twice', 2),
  }),
  inline(
    'uwave',
    'wavy colored underline',
    'text',
    t('Wavy', '<w:u w:val="wave" w:color="FF0000"/>'),
    'Wavy',
    {
      inside: at('Wavy', 2),
    }
  ),
  inline('strike', 'strikethrough', 'text', t('Lined', '<w:strike/>'), 'Lined', {
    inside: at('Lined', 2),
  }),
  inline('dstrike', 'double strikethrough', 'text', t('Doubled', '<w:dstrike/>'), 'Doubled', {
    inside: at('Doubled', 3),
  }),
  inline(
    'font',
    'font family',
    'text',
    t('Serif', '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/>'),
    'Serif',
    {
      inside: at('Serif', 2),
    }
  ),
  inline(
    'size',
    'font size',
    'text',
    t('Large', '<w:sz w:val="40"/><w:szCs w:val="40"/>'),
    'Large',
    {
      inside: at('Large', 2),
    }
  ),
  inline(
    'theme',
    'theme color',
    'text',
    t('Accent', '<w:color w:val="4472C4" w:themeColor="accent1"/>'),
    'Accent',
    {
      inside: at('Accent', 3),
    }
  ),
  inline(
    'runshade',
    'run shading',
    'text',
    t('Shaded', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>'),
    'Shaded',
    {
      inside: at('Shaded', 3),
    }
  ),
  inline('caps', 'all caps', 'text', t('Loud', '<w:caps/>'), 'Loud', { inside: at('Loud', 2) }),
  inline('smallcaps', 'small caps', 'text', t('Quiet', '<w:smallCaps/>'), 'Quiet', {
    inside: at('Quiet', 2),
  }),
  inline('sup', 'superscript', 'text', t('th', '<w:vertAlign w:val="superscript"/>'), 'th', {
    inside: at('th', 1),
  }),
  inline('sub', 'subscript', 'text', t('2', '<w:vertAlign w:val="subscript"/>'), '2'),
  inline(
    'complex',
    'complex script',
    'text',
    t('نص عربي', '<w:rtl/><w:cs/><w:rFonts w:cs="Arial"/><w:bCs/><w:szCs w:val="32"/>'),
    'نص عربي'
  ),
  inline(
    'effects',
    'outline, shadow, emboss and emphasis mark',
    'text',
    t('Raised', '<w:outline/><w:shadow/><w:emboss/><w:em w:val="dot"/>'),
    'Raised',
    {
      inside: at('Raised', 3),
    }
  ),
  inline(
    'eastasia',
    'East Asian text and layout',
    'text',
    t('漢字です', '<w:rFonts w:eastAsia="MS Mincho"/><w:eastAsianLayout w:id="1" w:combine="1"/>'),
    '漢字です',
    {
      inside: at('漢字です', 2),
    }
  ),
  inline(
    'charstyle',
    'character style (Strong)',
    'text',
    t('Strong words', '<w:rStyle w:val="Strong"/>'),
    'Strong words',
    {
      inside: at('Strong words', 6),
    }
  ),
  // Pictures and drawings
  inline(
    'wraptb',
    'floating picture, top and bottom wrap',
    'object',
    floating(11, '<wp:wrapTopAndBottom/>'),
    '￼'
  ),
  inline(
    'behind',
    'floating picture behind text',
    'object',
    floating(12, '<wp:wrapNone/>', 'behindDoc="1"'),
    '￼'
  ),
  inline(
    'tight',
    'floating picture, tight wrap',
    'object',
    floating(
      13,
      '<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon></wp:wrapTight>'
    ),
    '￼'
  ),
  inline(
    'crop',
    'cropped picture',
    'object',
    drawing(
      14,
      `${PIC}${picture(14, '<a:blip r:embed="rIdImg1"/>', '<a:srcRect l="10000" t="10000" r="10000" b="10000"/>')}`
    ),
    '￼'
  ),
  inline(
    'picfx',
    'picture adjustments and effects',
    'object',
    drawing(
      15,
      `${PIC}${picture(15, '<a:blip r:embed="rIdImg1"><a:lum bright="20000" contrast="10000"/></a:blip>')}`
    ).replace(
      '</pic:spPr>',
      '<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000"><a:srgbClr val="000000"/></a:outerShdw></a:effectLst></pic:spPr>'
    ),
    '￼'
  ),
  block(
    'pictable',
    'picture in a table cell',
    'table',
    table(1, `<w:tr>${cell(p(alphaOmega(pictureRun(20))))}</w:tr>`),
    {
      text: '\uFFFC',
    }
  ),
  inline(
    'picins',
    'picture inserted by a tracked change from Word',
    'revision',
    `<w:ins w:id="921" ${REV}>${pictureRun(16)}</w:ins>`,
    '￼'
  ),
  inline(
    'shape',
    'drawing shape',
    'object',
    `<w:r><mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:wps="${WPS}" Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="457200"/><wp:docPr id="17" name="Oval 17"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:oval xmlns:v="${V}" style="width:36pt;height:36pt"/></w:pict></mc:Fallback></mc:AlternateContent></w:r>`,
    '￼'
  ),
  inline(
    'ole',
    'embedded object (OLE)',
    'object',
    `<w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440"><v:shape xmlns:v="${V}" id="_x0000_i1025" style="width:72pt;height:72pt"/><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1"/></w:object></w:r>`,
    ''
  ),
  inline(
    'chart',
    'chart',
    'object',
    drawing(
      21,
      'uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdImg1"/>'
    ),
    '\uFFFC'
  ),
  inline(
    'smartart',
    'SmartArt diagram',
    'object',
    drawing(
      22,
      'uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdImg1" r:lo="rIdImg1" r:qs="rIdImg1" r:cs="rIdImg1"/>'
    ),
    '\uFFFC'
  ),
  inline('ink', 'ink annotation', 'object', '<w:r><w:contentPart r:id="rIdImg1"/></w:r>', ''),
  // Fields
  inline('numpages', 'NUMPAGES field', 'field', field('NUMPAGES', t('3')), '￼'),
  inline('datefield', 'DATE field', 'field', field('DATE \\@ "M/d/yyyy"', t('1/2/2026')), '￼'),
  inline(
    'hyperfield',
    'HYPERLINK field',
    'field',
    field('HYPERLINK "https://example.com"', t('site')),
    '￼'
  ),
  inline('noteref', 'NOTEREF cross-reference', 'field', field('NOTEREF _Ref1 \\h', t('1')), '￼'),
  inline('seq', 'SEQ field', 'field', field('SEQ Figure \\* ARABIC', t('1')), '￼'),
  inline('merge', 'MERGEFIELD', 'field', field('MERGEFIELD Name', t('«Name»')), '￼'),
  inline('autonum', 'AUTONUM field', 'field', field('AUTONUM', t('1.')), '￼'),
  inline('styleref', 'STYLEREF field', 'field', field('STYLEREF "Heading 1"', t('Title')), '￼'),
  inline(
    'citation',
    'CITATION field',
    'field',
    field('CITATION Smi20 \\l 1033', t('(Smith, 2020)')),
    '￼'
  ),
  inline(
    'formtext',
    'legacy text form field',
    'field',
    field(
      'FORMTEXT',
      t('typed'),
      '<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Text1"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar>'
    ),
    '￼'
  ),
  inline(
    'formcheck',
    'legacy check box form field',
    'field',
    field(
      'FORMCHECKBOX',
      '',
      '<w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar>'
    ),
    '￼'
  ),
  // Content controls, every type, with prompts
  inline(
    'dropprompt',
    'drop-down list showing its prompt',
    'control',
    sdt(
      '<w:showingPlcHdr/><w:dropDownList><w:listItem w:displayText="One" w:value="1"/><w:listItem w:displayText="Two" w:value="2"/></w:dropDownList>',
      t('Choose an item.', PROMPT)
    ),
    'Choose an item.',
    {
      inside: at('Choose', 3),
    }
  ),
  inline(
    'dateprompt',
    'date picker showing its prompt',
    'control',
    sdt(
      '<w:showingPlcHdr/><w:date><w:dateFormat w:val="M/d/yyyy"/></w:date>',
      t('Click to enter a date.', PROMPT)
    ),
    'Click to enter a date.',
    {
      inside: at('Click', 3),
    }
  ),
  inline(
    'richprompt',
    'rich text control showing its prompt',
    'control',
    sdt('<w:showingPlcHdr/>', t('Click here to enter text.', PROMPT)),
    'Click here to enter text.',
    {
      inside: at('Click', 3),
    }
  ),
  inline('piccontrol', 'picture control', 'control', sdt('<w:picture/>', pictureRun(18)), '￼'),
  inline(
    'gallery',
    'building block gallery control',
    'control',
    sdt('<w:docPartList><w:docPartGallery w:val="Quick Parts"/></w:docPartList>', t('Part')),
    'Part',
    {
      inside: at('Part', 2),
    }
  ),
  inline('group', 'group control', 'control', sdt('<w:group/>', t('Grouped')), 'Grouped', {
    inside: at('Grouped', 3),
  }),
  inline(
    'bound',
    'control bound to custom XML',
    'control',
    sdt(
      '<w:dataBinding w:xpath="/root/name" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/><w:text/>',
      t('Bound')
    ),
    'Bound',
    {
      inside: at('Bound', 2),
    }
  ),
  // A repeating section control
  block(
    'repeating',
    'repeating section',
    'block',
    sdt(
      '<w15:repeatingSection xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
      sdt(
        '<w15:repeatingSectionItem xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
        AO
      )
    )
  ),
  // Breaks outside runs
  inline('barebr', 'break outside a run', 'object', '<w:br/>', ''),
  // Paragraph formatting
  paragraph('right', 'right aligned', '<w:jc w:val="right"/>'),
  paragraph('left', 'left aligned', '<w:jc w:val="left"/>'),
  paragraph('justify', 'justified', '<w:jc w:val="both"/>'),
  paragraph('centered', 'centered', '<w:jc w:val="center"/>'),
  paragraph('lineauto', 'line spacing, multiple', '<w:spacing w:line="360" w:lineRule="auto"/>'),
  paragraph('exact', 'line spacing, exact', '<w:spacing w:line="320" w:lineRule="exact"/>'),
  paragraph('atleast', 'line spacing, at least', '<w:spacing w:line="300" w:lineRule="atLeast"/>'),
  paragraph('beforeafter', 'space before and after', '<w:spacing w:before="240" w:after="240"/>'),
  paragraph('indents', 'left and right indents', '<w:ind w:left="1440" w:right="1440"/>'),
  paragraph('hanging', 'hanging indent', '<w:ind w:left="720" w:hanging="720"/>'),
  paragraph('firstline', 'first-line indent', '<w:ind w:firstLine="720"/>'),
  paragraph(
    'tabstops',
    'tab stops and leaders',
    '<w:tabs><w:tab w:val="center" w:leader="hyphen" w:pos="4680"/><w:tab w:val="decimal" w:leader="underscore" w:pos="8000"/></w:tabs>'
  ),
  paragraph(
    'borders',
    'borders on every edge and a fill',
    '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/><w:left w:val="double" w:sz="4" w:space="4" w:color="FF0000"/><w:bottom w:val="dotted" w:sz="8" w:space="1" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="4" w:color="auto"/><w:between w:val="single" w:sz="4" w:space="1" w:color="auto"/></w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/>'
  ),
  paragraph(
    'keeps',
    'keep with next, keep lines, page break before',
    '<w:keepNext/><w:keepLines/><w:pageBreakBefore/>'
  ),
  paragraph('widow', 'widow control off', '<w:widowControl w:val="0"/>'),
  paragraph('contextual', 'contextual spacing', '<w:contextualSpacing/><w:spacing w:after="200"/>'),
  paragraph('outline', 'outline level', '<w:outlineLvl w:val="1"/>'),
  paragraph(
    'eastpara',
    'East Asian line breaking',
    '<w:kinsoku w:val="0"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="0"/><w:autoSpaceDE w:val="0"/><w:snapToGrid w:val="0"/>'
  ),
  paragraph('nohyphen', 'automatic hyphenation off', '<w:suppressAutoHyphens/>'),
  paragraph('bidi', 'right-to-left paragraph', '<w:bidi/>'),
  paragraph(
    'dropcap',
    'drop cap frame',
    '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'
  ),
  paragraph('quote', 'paragraph style (Quote)', '<w:pStyle w:val="Quote"/>'),
  paragraph(
    'pprchange',
    'paragraph properties changed by Word',
    `<w:jc w:val="center"/><w:pPrChange w:id="922" ${REV}><w:pPr/></w:pPrChange>`
  ),
  // Lists
  paragraph('bullet', 'bullet list item', list(2)),
  paragraph('multilevel', 'second level of a multi-level list', list(3, 1)),
  paragraph('roman', 'roman numbered item', list(4)),
  paragraph('restart', 'numbering restarted at 5', list(5)),
  paragraph('custom', 'custom number text (Article 1.)', list(6)),
  paragraph('picbullet', 'picture bullet', list(7)),
  block(
    'listrun',
    'middle item of a three-item list',
    'block',
    p(t('Item one'), list(1)) + p(alphaOmega(t('middle')), list(1)) + p(t('Item three'), list(1)),
    {
      p: 2,
      around: ['Item one', 'Item three'],
    }
  ),
  block(
    'listcell',
    'list item inside a table cell',
    'table',
    table(1, `<w:tr>${cell(p(alphaOmega(t('middle')), list(2)))}</w:tr>`)
  ),
  // Tables
  block(
    'tblborders',
    'table and cell borders and shading',
    'table',
    table(
      2,
      `<w:tr>${cell(AO, '<w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>')}${cell(p(t('Fee')))}</w:tr>`,
      '<w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>'
    ),
    {
      around: ['Before the item.', 'Fee'],
    }
  ),
  block(
    'widths',
    'fixed table and column widths',
    'table',
    table(
      2,
      `<w:tr>${cell(AO, '<w:tcW w:w="4000" w:type="dxa"/>')}${cell(p(t('Fee')), '<w:tcW w:w="2000" w:type="dxa"/>')}</w:tr>`,
      '<w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
    ),
    {
      around: ['Before the item.', 'Fee'],
    }
  ),
  block(
    'tblstyle',
    'table style with header row and banding',
    'table',
    table(
      1,
      `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell(AO)}</w:tr><w:tr>${cell(p(t('Row two')))}</w:tr>`,
      '<w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
    ),
    {
      around: ['Before the item.', 'Row two'],
    }
  ),
  block(
    'rtltable',
    'right-to-left table',
    'table',
    table(1, `<w:tr>${cell(AO)}</w:tr>`, '<w:bidiVisual/><w:tblW w:w="0" w:type="auto"/>')
  ),
  block(
    'floattable',
    'floating table',
    'table',
    table(
      1,
      `<w:tr>${cell(AO)}</w:tr>`,
      '<w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:tblpY="1"/><w:tblW w:w="0" w:type="auto"/>'
    )
  ),
  block(
    'textdir',
    'vertical cell text',
    'table',
    table(
      1,
      `<w:tr>${cell(AO, '<w:tcW w:w="3000" w:type="dxa"/><w:textDirection w:val="btLr"/>')}</w:tr>`
    )
  ),
  block(
    'cantsplit',
    "row that can't split across pages",
    'table',
    table(1, `<w:tr><w:trPr><w:cantSplit/></w:trPr>${cell(AO)}</w:tr>`)
  ),
  block(
    'cellrev',
    'cell shading changed by Word',
    'revision',
    table(
      1,
      `<w:tr>${cell(AO, `<w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:tcPrChange w:id="923" ${REV}><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr></w:tcPrChange>`)}</w:tr>`
    )
  ),
  block(
    'rowins',
    'row inserted by Word',
    'revision',
    table(
      1,
      `<w:tr>${cell(p(t('Kept')))}</w:tr><w:tr><w:trPr><w:ins w:id="924" ${REV}/></w:trPr>${cell(p(`<w:ins w:id="924" ${REV}>${alphaOmega(t('middle'))}</w:ins>`))}</w:tr>`
    ),
    {
      around: ['Kept', 'After the item.'],
      p: 2,
    }
  ),
  block(
    'mergeboth',
    'merged across and down in one table',
    'table',
    table(
      3,
      `<w:tr>${cell(AO, '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}${cell(p(t('Fee')))}</w:tr><w:tr>${cell('<w:p/>', '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:vMerge/>')}${cell(p(t('Due')))}</w:tr>`
    ),
    {
      around: ['Before the item.', 'Fee'],
    }
  ),
  // Sections and page layout
  paragraph(
    'nextpage',
    'next-page section break, landscape',
    sectPr('<w:type w:val="nextPage"/>').replace(
      '<w:pgSz w:w="12240" w:h="15840"/>',
      '<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>'
    )
  ),
  paragraph(
    'cols',
    'two text columns',
    sectPr('<w:type w:val="continuous"/>').replace(
      '</w:sectPr>',
      '<w:cols w:num="2" w:space="720"/></w:sectPr>'
    )
  ),
  paragraph(
    'sectfancy',
    'page borders, line numbers, vertical alignment, page number format',
    sectPr(
      '<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders><w:lnNumType w:countBy="1"/><w:pgNumType w:fmt="upperRoman" w:start="1"/>'
    ).replace('</w:sectPr>', '<w:vAlign w:val="center"/></w:sectPr>')
  ),
  {
    ...paragraph('background', 'page background color', ''),
    extras: { ...MATRIX_EXTRAS, documentPrefix: '<w:background w:color="FFF2CC"/>' },
  },
  // Tracked changes from Word
  block(
    'movefile',
    'move from Word, across paragraphs',
    'revision',
    p(
      `<w:moveFromRangeStart w:id="925" ${REV} w:name="mv9"/><w:moveFrom w:id="926" ${REV}>${t('Moved ')}</w:moveFrom><w:moveFromRangeEnd w:id="925"/>` +
        alphaOmega(t('middle'))
    ) +
      p(
        t('Landing ') +
          `<w:moveToRangeStart w:id="927" ${REV} w:name="mv9"/><w:moveTo w:id="928" ${REV}>${t('Moved ')}</w:moveTo><w:moveToRangeEnd w:id="927"/>`
      ),
    {
      around: ['Before the item.', 'Landing Moved '],
    }
  ),
  block(
    'markdel',
    'paragraph mark deleted by Word',
    'revision',
    p(alphaOmega(t('middle')), `<w:rPr><w:del w:id="929" ${REV}/></w:rPr>`) + p(t('Joined next.')),
    {
      around: ['Before the item.', 'Joined next.'],
    }
  ),
  // Equations, symbols and text boxes past the spike's
  inline(
    'mathfrac',
    'equation with a fraction',
    'object',
    `<m:oMath xmlns:m="${M}"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>`,
    '￼'
  ),
  // Stories past the body
  furniture('header', 'header', 'header', t('middle'), 'middle'),
  furniture('footer', 'footer', 'footer', t('middle'), 'middle'),
  furniture('firstheader', 'first-page header', 'header', t('middle'), 'middle', 'first'),
  furniture('evenheader', 'even-page header', 'header', t('middle'), 'middle', 'even'),
  furniture(
    'watermark',
    'header with a watermark',
    'header',
    watermark + t('middle'),
    '\uFFFCmiddle'
  ),
  furniture('picheader', 'picture in a header', 'header', pictureRun(19), '\uFFFC'),
  note('footnote', 'footnote text', 'footnote', t('middle')),
  note('endnote', 'endnote text', 'endnote', t('middle')),
  textBoxText('boxtext', 'text box text', t('middle'), 'middle'),
  // A table of contents result
  block('toc', 'table of contents result', 'field', toc, {
    around: ['Before the item.', 'Second entry'],
  }),
];
