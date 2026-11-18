// The sources of the paste tests of group 9.10: a document with a style, a numbered list, bold
// text, a link and a table, one with a picture (embedded, or linked to an address), one with a
// shape, a VML picture or a picture with an SVG version as Word writes them, and their
// clipboard flavours as the editor copies them.

import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import {
  docx,
  emfBytes,
  MATRIX_EXTRAS,
  MC,
  p,
  pictureRun,
  R,
  RED_SVG,
  t,
  V,
  W,
  WPS,
} from '../word-fixtures.ts';

const SOURCE_STYLES =
  `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
  '<w:name w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Hyperlink">' +
  '<w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ClauseTitle"><w:name w:val="Clause Title"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:caps/></w:rPr></w:style>' +
  '</w:styles>';

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** Paragraphs 0 to 3, a table, and paragraph 4. */
const SOURCE = docx(
  p(t('Payment'), '<w:pStyle w:val="ClauseTitle"/>') +
    p(t('First term'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    p(t('Second term'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    p(
      t('Pay ') +
        t('now', '<w:b/>') +
        '<w:hyperlink r:id="rIdPay"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>' +
        '<w:t>the portal</w:t></w:r></w:hyperlink>' +
        t(' today.')
    ) +
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/>' +
    `<w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell('Fee')}${cell('100')}</w:tr></w:tbl>` +
    p(t('End.')),
  {
    ...MATRIX_EXTRAS,
    parts: { ...MATRIX_EXTRAS.parts, 'word/styles.xml': SOURCE_STYLES },
    relationships:
      `${MATRIX_EXTRAS.relationships ?? ''}` +
      `<Relationship Id="rIdPay" Type="${R}/hyperlink" Target="https://pay.example.com/" TargetMode="External"/>`,
  }
);

export interface Copied {
  readonly text: string;
  readonly html: string | null;
}

/** The address of the linked picture: a test never reaches it. */
export const LINKED_PICTURE = 'https://images.example.com/seal.png';

/**
 * How a pasted picture names its image: the file's bytes, an address, or both; or the file's
 * bytes with Word 365's "decorative" flag on its alt text.
 */
export type PictureSource = 'embedded' | 'linked' | 'both' | 'decorative';

const BLIP_LINKS: Readonly<Record<PictureSource, string>> = {
  embedded: 'r:embed="rIdImg1"',
  linked: 'r:link="rIdSeal"',
  both: 'r:embed="rIdImg1" r:link="rIdSeal"',
  decorative: 'r:embed="rIdImg1"',
};

/** Word 365's "decorative" flag, in a drawing's `wp:docPr` extension list. */
export const DECORATIVE =
  '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative ' +
  'xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/>' +
  '</a:ext></a:extLst>';

/** A line with a picture whose image the file holds (`RED_PNG`), links to an address, or both. */
const pictureSource = (source: PictureSource) => {
  const run = pictureRun(3).replace('r:embed="rIdImg1"', BLIP_LINKS[source]);
  const flagged =
    source === 'decorative'
      ? run.replace('name="Picture 3"/>', `name="Picture 3">${DECORATIVE}</wp:docPr>`)
      : run;
  return docx(p(t('Seal ') + flagged + t(' here.')), {
    ...MATRIX_EXTRAS,
    relationships:
      `${MATRIX_EXTRAS.relationships ?? ''}` +
      `<Relationship Id="rIdSeal" Type="${R}/image" Target="${LINKED_PICTURE}" TargetMode="External"/>`,
  });
};

const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const WPC = 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas';

/** The namespaces Word declares on a document's root: a shape's prefixes resolve through them. */
const WORD_ROOT =
  `xmlns:w="${W}" xmlns:r="${R}" xmlns:mc="${MC}" xmlns:wps="${WPS}" xmlns:wpg="${WPG}" ` +
  `xmlns:wpc="${WPC}" ` +
  `xmlns:v="${V}" ` +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" mc:Ignorable="w14 wp14"';

/** Word's DrawingML copy of a VML shape: base64 in lines of 76. */
const GFXDATA =
  'UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#xA;' +
  '90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA';

/** A text box as Word writes one: a shape in markup compatibility, its VML copy beside it. */
const TEXT_BOX =
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" ' +
  'distB="0" distL="0" distR="0" wp14:anchorId="3209AFDE" wp14:editId="38ACA0DB">' +
  '<wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' +
  `<wp:docPr id="2" name="Text Box 2"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${WPS}">` +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
  '<a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill>' +
  '<a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>' +
  `<wps:txbx><w:txbxContent>${p(t('Box text'))}</w:txbxContent></wps:txbx>` +
  '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" ' +
  'bIns="45720" anchor="t" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr></wps:wsp>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
  '<v:shape id="Text Box 2" o:spid="_x0000_s1026" type="#_x0000_t202" ' +
  'style="width:72pt;height:36pt;visibility:visible" w14:anchorId="3CEC9D07" ' +
  `o:gfxdata="${GFXDATA}"><v:textbox><w:txbxContent>${p(t('Box text'))}</w:txbxContent>` +
  '</v:textbox><w10:wrap type="none"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

/** A group of an oval and a picture, and its VML copy. */
const GROUP =
  '<w:r><mc:AlternateContent><mc:Choice Requires="wpg"><w:drawing><wp:inline distT="0" ' +
  'distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="4" name="Group 4"/><a:graphic><a:graphicData uri="${WPG}"><wpg:wgp>` +
  '<wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm></wpg:grpSpPr><wps:wsp>' +
  '<wps:cNvPr id="5" name="Oval 5"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
  '<a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
  '</wps:spPr><wps:bodyPr/></wps:wsp><pic:pic><pic:nvPicPr><pic:cNvPr id="6" name="red.png"/>' +
  '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch>' +
  '<a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="457200" y="0"/>' +
  '<a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '</pic:spPr></pic:pic></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>' +
  '</mc:Choice><mc:Fallback><w:pict><v:group style="width:72pt;height:36pt" ' +
  'coordsize="914400,457200"><v:oval style="width:36pt;height:36pt"/><v:shape ' +
  'style="left:457200;width:36pt;height:36pt"><v:imagedata r:id="rIdImg1" o:title=""/>' +
  '</v:shape></v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

/** A drawing canvas with its background, a shape and a picture, as Word writes one. */
const CANVAS =
  '<w:r><mc:AlternateContent><mc:Choice Requires="wpc"><w:drawing><wp:inline distT="0" ' +
  'distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/>' +
  `<wp:docPr id="7" name="Canvas 7"/><a:graphic><a:graphicData uri="${WPC}"><wpc:wpc>` +
  '<wpc:bg><a:noFill/></wpc:bg><wpc:whole><a:ln><a:noFill/></a:ln></wpc:whole><wps:wsp>' +
  '<wps:cNvPr id="8" name="Oval 8"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>' +
  '<a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
  '</wps:spPr><wps:bodyPr/></wps:wsp><pic:pic><pic:nvPicPr><pic:cNvPr id="9" name="red.png"/>' +
  '<pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch>' +
  '<a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="457200" y="0"/>' +
  '<a:ext cx="457200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '</pic:spPr></pic:pic></wpc:wpc></a:graphicData></a:graphic></wp:inline></w:drawing>' +
  '</mc:Choice><mc:Fallback><w:pict><v:group editas="canvas" style="width:72pt;height:36pt" ' +
  'coordsize="914400,457200"><v:oval style="width:36pt;height:36pt"/><v:shape ' +
  'style="left:457200;width:36pt;height:36pt"><v:imagedata r:id="rIdImg1" o:title=""/>' +
  '</v:shape></v:group></w:pict></mc:Fallback></mc:AlternateContent></w:r>';

/** A VML picture, as older Word writes one, of an EMF image. */
const VML_PICTURE =
  '<w:r><w:pict><v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" ' +
  'o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' +
  '<v:stroke joinstyle="miter"/><v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/>' +
  '</v:formulas><v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' +
  '<o:lock v:ext="edit" aspectratio="t"/></v:shapetype><v:shape id="_x0000_i1025" ' +
  'type="#_x0000_t75" style="width:36pt;height:36pt"><v:imagedata r:id="rIdEmf" ' +
  'o:title=""/></v:shape></w:pict></w:r>';

/** A picture with an SVG version, its PNG image the fallback. */
const SVG_PICTURE = pictureRun(3).replace(
  '<a:blip r:embed="rIdImg1"/>',
  '<a:blip r:embed="rIdImg1"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
    '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" ' +
    'r:embed="rIdSvg"/></a:ext></a:extLst></a:blip>'
);

/** Word 365's sketched outline (`ask`), an extension the format does not hold, and Office 2016's creation id. */
export const SKETCH_EXTENSION =
  '<a:ext uri="{C807C97D-BFC1-408E-A445-0C87EB9F89A2}"><ask:lineSketchStyleProps ' +
  'xmlns:ask="http://schemas.microsoft.com/office/drawing/2018/sketchyshapes" sd="1219033472"/></a:ext>';
const SKETCHED_PICTURE = pictureRun(3).replace(
  '<a:blip r:embed="rIdImg1"/>',
  `<a:blip r:embed="rIdImg1"><a:extLst>${SKETCH_EXTENSION}<a:ext ` +
    'uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId ' +
    'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{3D1E4C2A}"/>' +
    '</a:ext></a:extLst></a:blip>'
);

/** A 3D model (`am3d`), a namespace the format does not hold, with Word's picture of it as fallback. */
const MODEL_3D =
  '<w:r><mc:AlternateContent><mc:Choice ' +
  'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" Requires="am3d">' +
  pictureRun(3)
    .slice('<w:r>'.length, -'</w:r>'.length)
    .replace(
      /<a:graphicData .*<\/a:graphicData>/,
      '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2017/model3d">' +
        '<am3d:model3d/></a:graphicData>'
    ) +
  `</mc:Choice><mc:Fallback>${pictureRun(3).slice('<w:r>'.length, -'</w:r>'.length)}` +
  '</mc:Fallback></mc:AlternateContent></w:r>';

/** The shapes of the paste tests, by kind. */
export const SHAPES = {
  textbox: TEXT_BOX,
  group: GROUP,
  canvas: CANVAS,
  vml: VML_PICTURE,
  svg: SVG_PICTURE,
  /** A VML picture with its embedded copy and the address it was linked to. */
  vmlBoth: VML_PICTURE.replace('o:title=""', `o:href="${LINKED_PICTURE}" o:title=""`),
  /** A VML picture whose image only an external relationship names. */
  vmlLinked: VML_PICTURE.replace('r:id="rIdEmf"', 'r:id="rIdSeal"'),
  /** A VML picture whose image older Word names by the Office relationship id `o:relid`. */
  vmlRelid: VML_PICTURE.replace('r:id="rIdEmf"', 'o:relid="rIdEmf"'),
  /** A VML picture with a second image by `r:pict`. */
  vmlPict: VML_PICTURE.replace('r:id="rIdEmf"', 'r:id="rIdEmf" r:pict="rIdSvg"'),
  /** A picture with an extension of a namespace the format does not hold. */
  sketched: SKETCHED_PICTURE,
  /** A 3D model whose namespace the format does not hold, and its fallback picture. */
  model3d: MODEL_3D,
  /** A VML picture with its embedded copy and its link to the file (`r:href`). */
  vmlHref: VML_PICTURE.replace('r:id="rIdEmf"', 'r:id="rIdEmf" r:href="rIdSeal"'),
} as const;

/**
 * A line with a shape, a VML picture or a picture with an SVG version, as Word writes it. `svg`
 * is the bytes of its SVG image.
 */
export const shapeSource = (kind: keyof typeof SHAPES, svg: Uint8Array = RED_SVG) =>
  docx('', {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/document.xml':
        `<w:document ${WORD_ROOT}><w:body>` +
        `${p(t('Seal ') + SHAPES[kind] + t(' here.'))}</w:body></w:document>`,
      'word/media/seal.emf': emfBytes(),
      'word/media/seal.svg': svg,
    },
    contentTypes:
      `${MATRIX_EXTRAS.contentTypes ?? ''}<Default Extension="emf" ContentType="image/x-emf"/>` +
      '<Default Extension="svg" ContentType="image/svg+xml"/>',
    relationships:
      `${MATRIX_EXTRAS.relationships ?? ''}` +
      `<Relationship Id="rIdEmf" Type="${R}/image" Target="media/seal.emf"/>` +
      `<Relationship Id="rIdSvg" Type="${R}/image" Target="media/seal.svg"/>` +
      `<Relationship Id="rIdSeal" Type="${R}/image" Target="${LINKED_PICTURE}" TargetMode="External"/>`,
  });

type SourceEditor = ReturnType<typeof createDocxEditor>;

/** An editor of `source` in a container of its own; `close` removes both. */
function sourceEditor(source: Uint8Array): { editor: SourceEditor; close: () => void } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: source, author: 'Source' });
  const close = () => {
    editor.destroy();
    container.remove();
  };
  return { editor, close };
}

/**
 * The clipboard flavours of an editor's text from the start of paragraph `first` to `end` in
 * paragraph `last` (its end when not given).
 */
function flavoursBetween(
  editor: SourceEditor,
  first: number,
  last: number,
  end: number | undefined
): Copied {
  const s = editor.surface;
  if (!s) {
    throw new Error('no surface');
  }

  const ids = s.session.paragraphIds();
  const from = ids[first];
  const to = ids[last];
  if (from === undefined || to === undefined) {
    throw new Error('no paragraphs');
  }

  s.setSelection({
    anchor: { paragraphId: from, offset: 0 },
    head: {
      paragraphId: to,
      offset: end ?? (paragraphTextOf(s.session.part(), to) ?? '').length,
    },
  });
  return s.copyFlavours();
}

/**
 * The clipboard flavours of the source document from the start of paragraph `first` to `end`
 * in paragraph `last` (its end when not given).
 */
export function copiedBetween(
  first: number,
  last: number,
  end?: number,
  source: Uint8Array = SOURCE
): Copied {
  const { editor, close } = sourceEditor(source);
  try {
    return flavoursBetween(editor, first, last, end);
  } finally {
    close();
  }
}

/** The clipboard flavours of the source document's first `count` paragraphs. */
export const copied = (count: number, end?: number): Copied => copiedBetween(0, count - 1, end);

/** Plain, bold and linked text of one line, without its paragraph mark. */
export const copiedLine = (): Copied => copiedBetween(3, 3, 'Pay nowthe portal'.length);

/** A line with a picture in it, without its paragraph mark: embedded, linked, or both. */
export const copiedPicture = (source: PictureSource = 'embedded'): Copied =>
  copiedBetween(0, 0, 'Seal \uFFFC here'.length, pictureSource(source));

/** A line with a shape, a VML picture or a picture with an SVG version, without its mark. */
export const copiedShape = (kind: keyof typeof SHAPES): Copied =>
  copiedBetween(0, 0, 'Seal \uFFFC here'.length, shapeSource(kind));

/**
 * A line with a picture whose SVG version is `svg`, without its mark. A file past the editor's
 * yield threshold opens behind a timer, so this waits for the source editor to open.
 */
export async function copiedSvgPicture(svg: Uint8Array): Promise<Copied> {
  const { editor, close } = sourceEditor(shapeSource('svg', svg));
  try {
    for (let round = 0; editor.surface === null && round < 100; round += 1) {
      await new Promise((resolve) => setTimeout(resolve, 10));
    }

    return flavoursBetween(editor, 0, 0, 'Seal \uFFFC here'.length);
  } finally {
    close();
  }
}
