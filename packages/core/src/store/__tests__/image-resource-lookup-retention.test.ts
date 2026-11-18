// An image lookup must not keep the document trees of the package it was made for.
//
// A layout owner makes one lookup and keeps it for every later revision whose media is the
// same, which in a long session is nearly every revision. The lookup held its package, and
// with it the package's first document tree, so the session kept the first version of every
// paragraph it edited. It reads only the relationships, part bytes and content types.

import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import {
  createOwnedImageResourceLookup,
  imageResourceLookupFor,
  validateRasterHeader,
  type ImageDecodePort,
  type ImageResourceLookup,
} from '../package/image-resources.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../package/ooxml-package.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PNG_1X1 = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII='
  ),
  (c) => c.charCodeAt(0)
);
const MEDIA = '/word/media/image1.png';

const picture =
  `<w:p><w:r><w:drawing><wp:inline xmlns:wp="${WP}">` +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="pic"/>' +
  `<a:graphic xmlns:a="${A}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic xmlns:pic="${PIC}"><pic:nvPicPr><pic:cNvPr id="1" name=""/><pic:cNvPicPr/></pic:nvPicPr>` +
  '<pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill><pic:spPr/></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';

function openPackage(): OoxmlPackage {
  const loaded = readOoxmlPackage(
    zipSync({
      '[Content_Types].xml': strToU8(
        `<Types xmlns="${CT}">` +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="png" ContentType="image/png"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '</Types>'
      ),
      '_rels/.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
      ),
      'word/document.xml': strToU8(
        `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${picture}<w:p/></w:body></w:document>`
      ),
      'word/_rels/document.xml.rels': strToU8(
        `<Relationships xmlns="${REL}"><Relationship Id="rId2" Type="${R}/image" Target="media/image1.png"/></Relationships>`
      ),
      'word/media/image1.png': PNG_1X1,
    })
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.package;
}

const decodePort: ImageDecodePort = {
  decode: async (bytes, mime) => {
    const header = validateRasterHeader(bytes, mime);
    if (!header) throw new Error('structural header required');
    return { ...header, dpiX: 96, dpiY: 96 };
  },
};

for (const [name, make] of [
  ['an owned lookup', createOwnedImageResourceLookup],
  ['a registry lookup', imageResourceLookupFor],
] as const)
  test(`${name} lets go of its package's document tree`, () => {
    const { lookup, tree } = ((): { lookup: ImageResourceLookup; tree: WeakRef<object> } => {
      const pkg = openPackage();
      const made = make(pkg, { decodePort });
      // While the package is held, the lookup still counts its drawings.
      expect(made.liveReferenceCount(MEDIA)).toBe(1);
      return { lookup: made, tree: new WeakRef(pkg.parts.get(pkg.mainDocumentPart)!.root) };
    })();
    Bun.gc(true);
    expect(tree.deref()).toBeUndefined();
    // Nothing holds that package any more, so no drawing of it references the picture.
    expect(lookup.liveReferenceCount(MEDIA)).toBe(0);
    lookup.dispose();
  });

test('a lookup kept across revisions resolves pictures from its own relationships', async () => {
  const pkg = openPackage();
  const lookup = createOwnedImageResourceLookup(pkg, { decodePort });
  const state = await lookup.resolveEmbedded('/word/document.xml', 'rId2');
  expect(state.kind).toBe('ready');
  lookup.dispose();
});
