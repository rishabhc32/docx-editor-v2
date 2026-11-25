// Rooms whose checkpoints hold what changes added outside the story text, in the manifest's
// section `resources` (design D14): a list level, a link's address, an embedded object, a chart's
// part, and a pasted list's picture bullet. Ana makes each one, pending.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OrderedMediaType } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { apply, at, EDIT, markOf } from './change-apply-fixture.ts';
import type { CheckpointRoom } from './checkpoint-fixture.ts';
import { mediaStore } from './fold-check.ts';
import { docx, p, R, RED_PNG, t, W, WML_CT } from './word-fixtures.ts';

/** Bytes kept as they are, never opened: an embedded object's part. */
const OLE_BYTES = new TextEncoder().encode('an embedded object, kept as bytes');

/** The media a change names, by its bytes. */
const mediaOf = (bytes: Uint8Array, contentType: OrderedMediaType) => ({
  sha256: orderedBlobHash(bytes),
  contentType,
  size: bytes.byteLength,
});

const PNG = mediaOf(RED_PNG, 'image/png');
const OLE = mediaOf(OLE_BYTES, 'application/vnd.openxmlformats-officedocument.oleObject');

/** A room's storage: the file's media and the bytes the rooms' changes name. */
export function resourceBlobs(base: OoxmlPackage): (sha256: string) => Uint8Array | undefined {
  const file = mediaStore(base);
  const named = new Map([
    [PNG.sha256, RED_PNG],
    [OLE.sha256, OLE_BYTES],
  ]);
  return (sha256) => named.get(sha256) ?? file(sha256);
}

/** A list whose definition declares only level 0, and a plain paragraph. */
const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
  '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
const LIST_FILE = docx(
  p(t('First item.'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    p(t('Plain words.')),
  {
    parts: { 'word/numbering.xml': NUMBERING },
    contentTypes: `<Override PartName="/word/numbering.xml" ContentType="${WML_CT}.numbering+xml"/>`,
    relationships: `<Relationship Id="rIdNum" Type="${R}/numbering" Target="numbering.xml"/>`,
  }
);

/** Ana's paste of one paragraph of `runs` at the start of the first paragraph. */
const paste = (doc: OrderedDocument, runs: readonly object[], more: object) =>
  ({
    kind: 'insert-fragment',
    ...EDIT,
    place: at(doc, 0, 0),
    blocks: [{ e: 'w:p', c: runs }],
    ...more,
  }) as unknown as Parameters<typeof apply>[2];

/** Increase Indent to a level the list does not declare (`list-level`). */
export const LIST_LEVEL_ROOM: CheckpointRoom = {
  name: 'a list level',
  source: LIST_FILE,
  edit: (doc) => {
    const level = { numId: '1', level: 1, format: 'ordered', marks: [markOf(doc, 0)] } as const;
    apply(doc, 'ana:1', { kind: 'list-level', ...EDIT, ...level });
  },
};

/** A link on Ana's own new words: its address is a relationship to outside. */
export const LINK_ROOM: CheckpointRoom = {
  name: "a link's address",
  source: docx(p(t('Cloud and serverless.'))),
  edit: (doc) => {
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
    const link = { from: at(doc, 0, 6), to: at(doc, 0, 9), href: 'https://a.example/' };
    apply(doc, 'ana:2', { kind: 'insert-link', ...EDIT, ...link });
  },
};

/** A pasted embedded object: its part is media, named by a relationship. */
export const OBJECT_ROOM: CheckpointRoom = {
  name: 'an embedded object',
  source: docx(p(t('Fees.'))),
  edit: (doc) => {
    const object = {
      e: 'w:r',
      c: [{ e: 'w:object', c: [{ e: 'o:OLEObject', a: { Type: 'Embed', 'r:id': '#0' } }] }],
    };
    apply(doc, 'ana:1', paste(doc, [object], { media: [OLE] }));
  },
};

/** A pasted chart: its part, and the relationship to it. */
export const CHART_ROOM: CheckpointRoom = {
  name: "a chart's part",
  source: docx(p(t('Fees.'))),
  edit: (doc) => {
    const data = {
      e: 'a:graphicData',
      a: { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' },
      c: [{ e: 'c:chart', a: { 'r:id': '#0' } }],
    };
    const inline = {
      e: 'wp:inline',
      c: [
        { e: 'wp:extent', a: { cx: '5486400', cy: '3200400' } },
        { e: 'wp:docPr', a: { id: '7', name: 'Chart 1' } },
        { e: 'a:graphic', c: [data] },
      ],
    };
    const chart = { e: 'w:r', c: [{ e: 'w:drawing', c: [inline] }] };
    const part = { type: 'chart', root: { e: 'c:chartSpace', c: [{ e: 'c:chart' }] } };
    apply(doc, 'ana:1', paste(doc, [chart], { parts: [part] }));
  },
};

/** A pasted list with a picture bullet: its image is media of the numbering part. */
export const BULLET_ROOM: CheckpointRoom = {
  name: 'a picture bullet',
  source: docx(p(t('Fees.'))),
  edit: (doc) => {
    const level = {
      e: 'w:lvl',
      a: { 'w:ilvl': '0' },
      c: [
        { e: 'w:numFmt', a: { 'w:val': 'bullet' } },
        { e: 'w:lvlText', a: { 'w:val': '' } },
        { e: 'w:lvlPicBulletId', a: { 'w:val': '#0' } },
      ],
    };
    const image = { e: 'v:shape', c: [{ e: 'v:imagedata', a: { 'r:id': '#0' } }] };
    const bullet = { e: 'w:numPicBullet', c: [{ e: 'w:pict', c: [image] }] };
    const numPr = {
      e: 'w:numPr',
      c: [
        { e: 'w:ilvl', a: { 'w:val': '0' } },
        { e: 'w:numId', a: { 'w:val': '#0' } },
      ],
    };
    const words = { e: 'w:r', c: [{ e: 'w:t', c: ['A.'] }] };
    apply(doc, 'ana:1', {
      kind: 'insert-fragment',
      ...EDIT,
      place: at(doc, 0, 0),
      blocks: [{ e: 'w:p', c: [{ e: 'w:pPr', c: [numPr] }, words] }],
      whole: true,
      lists: [{ abstract: { e: 'w:abstractNum', c: [level] } }],
      bullets: [bullet],
      media: [PNG],
    });
  },
};

export const RESOURCE_ROOMS: readonly CheckpointRoom[] = [
  LIST_LEVEL_ROOM,
  LINK_ROOM,
  OBJECT_ROOM,
  CHART_ROOM,
  BULLET_ROOM,
];
