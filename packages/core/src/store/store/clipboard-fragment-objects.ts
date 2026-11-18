// The parts that travel with a drawing on copy and paste, as Word copies them: a chart's, a
// diagram's (SmartArt), an embedded object's and ink's.

import { readOoxmlPart, serializeOoxmlPart } from '../package/ooxml-tree.ts';
import type { OoxmlAttribute, OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import {
  relationshipsOf,
  resolveContentTypeOf,
  withNewPart,
  withRelationships,
  withRelationshipsPartFor,
} from '../package/package-edit.ts';
import { partNameKey, resolveInternalTarget } from '../package/opc-names.ts';
import type { RelationshipRecord } from '../package/relationships.ts';
import { withBinaryParts, validateEmbeddedImageBytes } from '../package/drawing-package-edit.ts';
import { sniffImageMime, type SupportedImageMime } from '../package/image-resources.ts';
import { partNumberCandidates } from '../package/actor-scoped-ids.ts';
import { DRAWINGML_MAIN_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import { relativeTarget } from '../package/custom-xml-part.ts';

const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MS_2007 = 'http://schemas.microsoft.com/office/2007/relationships';
const MS_2011 = 'http://schemas.microsoft.com/office/2011/relationships';
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const DIAGRAM_NS = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
const CHART_STYLE_NS = 'http://schemas.microsoft.com/office/drawing/2012/chartStyle';

/** An Office 2016 chart's namespace (a box and whisker, a waterfall, a funnel: [MS-ODRAWXML]). */
const CHART_EX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';

/** The namespace of a diagram's drawing, and of `dataModelExt` in its data part. */
const DIAGRAM_DRAWING_NS = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

const CHART_REL = `${R_NS}/chart`;
const CHART_EX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
const OLE_REL = `${R_NS}/oleObject`;
const INK_REL = `${MS_2011}/ink`;
const DIAGRAM_DATA_REL = `${R_NS}/diagramData`;
const DIAGRAM_DRAWING_REL = `${MS_2007}/diagramDrawing`;
const PACKAGE_REL = `${R_NS}/package`;
const IMAGE_REL = `${R_NS}/image`;
const CHART_STYLE_REL = `${MS_2011}/chartStyle`;
const CHART_COLORS_REL = `${MS_2011}/chartColorStyle`;
const CHART_SHAPES_REL = `${R_NS}/chartUserShapes`;
const THEME_OVERRIDE_REL = `${R_NS}/themeOverride`;

/** The workbook content type of a chart's data (ECMA-376 Part 1, §15.2.11). */
export const EMBEDDED_WORKBOOK_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * What a part of a chart, a diagram, an embedded object or ink is, by the type of the
 * relationship to it.
 */
interface ObjectPartType {
  /** Its content type; `image/*` for a picture, whose type its signature decides. */
  readonly contentType: string;
  /** The name a paste gives a new one, as Word names it: stem, a number, extension. */
  readonly stem: string;
  readonly extension: string;
  /** Its root element (namespace, local name); null for a part of bytes. */
  readonly root: readonly [string, string] | null;
  /** The types its own relationships may have. */
  readonly relationships: readonly string[];
  /** For a part of bytes: the signature it must have (a zip when absent). */
  readonly signature?: 'cfb' | 'xml';
  /** Other content types it may have, with the name a paste gives each (an embedded file). */
  readonly variants?: ReadonlyMap<string, readonly [stem: string, extension: string]>;
}

const xmlType = (
  contentType: string,
  stem: string,
  root: readonly [string, string],
  relationships: readonly string[] = [IMAGE_REL]
): ObjectPartType => ({ contentType, stem, extension: 'xml', root, relationships });

const DML = 'application/vnd.openxmlformats-officedocument.drawingml';

/**
 * The parts that travel with a drawing, by relationship type. A chart's own relationships name
 * its embedded workbook, chart style, chart colors and pictures; a diagram's name its pictures.
 */
const PART_TYPES: ReadonlyMap<string, ObjectPartType> = new Map([
  [
    CHART_REL,
    xmlType(
      `${DML}.chart+xml`,
      '/word/charts/chart',
      [CHART_NS, 'chartSpace'],
      [
        PACKAGE_REL,
        CHART_STYLE_REL,
        CHART_COLORS_REL,
        CHART_SHAPES_REL,
        THEME_OVERRIDE_REL,
        IMAGE_REL,
      ]
    ),
  ],
  [
    CHART_EX_REL,
    xmlType(
      'application/vnd.ms-office.chartex+xml',
      '/word/charts/chartEx',
      [CHART_EX_NS, 'chartSpace'],
      [PACKAGE_REL, CHART_STYLE_REL, CHART_COLORS_REL, IMAGE_REL]
    ),
  ],
  [
    DIAGRAM_DATA_REL,
    xmlType(`${DML}.diagramData+xml`, '/word/diagrams/data', [DIAGRAM_NS, 'dataModel']),
  ],
  [
    `${R_NS}/diagramLayout`,
    xmlType(`${DML}.diagramLayout+xml`, '/word/diagrams/layout', [DIAGRAM_NS, 'layoutDef']),
  ],
  [
    `${R_NS}/diagramQuickStyle`,
    xmlType(`${DML}.diagramStyle+xml`, '/word/diagrams/quickStyle', [DIAGRAM_NS, 'styleDef']),
  ],
  [
    `${R_NS}/diagramColors`,
    xmlType(`${DML}.diagramColors+xml`, '/word/diagrams/colors', [DIAGRAM_NS, 'colorsDef']),
  ],
  [
    DIAGRAM_DRAWING_REL,
    xmlType('application/vnd.ms-office.drawingml.diagramDrawing+xml', '/word/diagrams/drawing', [
      DIAGRAM_DRAWING_NS,
      'drawing',
    ]),
  ],
  [
    CHART_STYLE_REL,
    xmlType('application/vnd.ms-office.chartstyle+xml', '/word/charts/style', [
      CHART_STYLE_NS,
      'chartStyle',
    ]),
  ],
  [
    CHART_COLORS_REL,
    xmlType('application/vnd.ms-office.chartcolorstyle+xml', '/word/charts/colors', [
      CHART_STYLE_NS,
      'colorStyle',
    ]),
  ],
  [
    CHART_SHAPES_REL,
    xmlType(`${DML}.chartshapes+xml`, '/word/drawings/drawing', [CHART_NS, 'userShapes']),
  ],
  [
    THEME_OVERRIDE_REL,
    xmlType(
      'application/vnd.openxmlformats-officedocument.themeOverride+xml',
      '/word/theme/themeOverride',
      [DRAWINGML_MAIN_NAMESPACE_URI, 'themeOverride'],
      []
    ),
  ],
  [
    PACKAGE_REL,
    {
      contentType: EMBEDDED_WORKBOOK_TYPE,
      stem: '/word/embeddings/Microsoft_Excel_Worksheet',
      extension: 'xlsx',
      root: null,
      relationships: [],
      // An embedded object's Office file, as Word names it
      variants: new Map([
        [
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          ['/word/embeddings/Microsoft_Word_Document', 'docx'],
        ],
        [
          'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          ['/word/embeddings/Microsoft_PowerPoint_Presentation', 'pptx'],
        ],
      ]),
    },
  ],
  [
    OLE_REL,
    {
      contentType: 'application/vnd.openxmlformats-officedocument.oleObject',
      stem: '/word/embeddings/oleObject',
      extension: 'bin',
      root: null,
      relationships: [],
      signature: 'cfb',
    },
  ],
  [
    INK_REL,
    {
      contentType: 'application/inkml+xml',
      stem: '/word/ink/ink',
      extension: 'xml',
      root: null,
      relationships: [],
      signature: 'xml',
    },
  ],
  [
    IMAGE_REL,
    {
      contentType: 'image/*',
      stem: '/word/media/image',
      extension: '',
      root: null,
      relationships: [],
    },
  ],
]);

/**
 * The relationship types by which a story names a part of a chart, a diagram, an embedded object
 * or ink: `c:chart/@r:id`, `cx:chart/@r:id`, a diagram's `dgm:relIds/@r:dm`, `@r:lo`, `@r:qs` and
 * `@r:cs`, an embedded object's `o:OLEObject/@r:id` and ink's `w14:contentPart/@r:id`. A linked
 * object (an external relationship) is never loaded, and an ActiveX control never travels.
 */
const STORY_TYPES: ReadonlySet<string> = new Set([
  CHART_REL,
  CHART_EX_REL,
  DIAGRAM_DATA_REL,
  `${R_NS}/diagramLayout`,
  `${R_NS}/diagramQuickStyle`,
  `${R_NS}/diagramColors`,
  DIAGRAM_DRAWING_REL,
  OLE_REL,
  PACKAGE_REL,
  INK_REL,
]);

/** At most this many parts travel with the drawings of one copy, and this deep. */
const MAX_OBJECT_PARTS = 256;
const MAX_DEPTH = 3;

/** Whether a story's relationship names a part of a chart, a diagram, an embedded object or ink. */
export function isObjectRelationship(record: RelationshipRecord): boolean {
  return record.targetMode !== 'External' && STORY_TYPES.has(record.type);
}

const SUPPORTED_RASTER_MIMES: ReadonlySet<string> = new Set([
  'image/png',
  'image/jpeg',
  'image/gif',
  'image/bmp',
  'image/webp',
]);

/** Aliases OPC files legitimately declare for the same signature class. */
const DECLARED_MIME_ALIASES: Readonly<Record<string, string>> = {
  'image/jpg': 'image/jpeg',
  'image/x-ms-bmp': 'image/bmp',
  'image/x-bmp': 'image/bmp',
};

/**
 * The content type fragment media is admitted under, or null to drop it.
 *
 * Signature over claim: the sniffed mime must agree with the declared class, and a
 * supported raster must additionally pass the same header + dimension gate the
 * insert-image lane applies (`validateEmbeddedImageBytes`). Vector and preserved formats
 * (SVG, TIFF, EMF, WMF) travel signature-checked; the paint lane re-validates and renders
 * them inert.
 */
export function admittedMediaMime(bytes: Uint8Array, declaredType: string): string | null {
  const sniffed = sniffImageMime(bytes);
  if (sniffed === 'unknown') return null;
  const normalizedDeclared = Object.hasOwn(DECLARED_MIME_ALIASES, declaredType)
    ? DECLARED_MIME_ALIASES[declaredType]!
    : declaredType;
  if (normalizedDeclared !== sniffed) return null;
  if (SUPPORTED_RASTER_MIMES.has(sniffed)) {
    if (!validateEmbeddedImageBytes(bytes, sniffed as SupportedImageMime)) return null;
  }

  return sniffed;
}

/** A zip's local file header: an embedded workbook is checked by it, and never opened. */
const isZip = (bytes: Uint8Array): boolean =>
  bytes.length >= 4 &&
  bytes[0] === 0x50 &&
  bytes[1] === 0x4b &&
  bytes[2] === 0x03 &&
  bytes[3] === 0x04;

/** An OLE compound file's signature (an embedded object's part), checked and never opened. */
export const isCompoundFile = (bytes: Uint8Array): boolean =>
  bytes.length >= 8 &&
  [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((byte, k) => bytes[k] === byte);

/** XML text: its first character, past a byte order mark and white space, is `<`. */
export function isXmlText(bytes: Uint8Array): boolean {
  let index = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  while (
    index < bytes.length &&
    index < 4096 &&
    (bytes[index] === 0x20 ||
      bytes[index] === 0x09 ||
      bytes[index] === 0x0a ||
      bytes[index] === 0x0d)
  ) {
    index += 1;
  }

  return bytes[index] === 0x3c;
}

/** The signature a part of bytes must have, by `ObjectPartType.signature` (a zip when absent). */
const SIGNATURE_CHECKS: Readonly<Record<'cfb' | 'xml' | 'zip', (bytes: Uint8Array) => boolean>> = {
  cfb: isCompoundFile,
  xml: isXmlText,
  zip: isZip,
};

/** The parts that travel with the drawings of one copy or one paste, by source part name. */
export interface ObjectParts {
  readonly xml: Map<string, { readonly part: OoxmlPart; readonly type: ObjectPartType }>;
  readonly bytes: Map<
    string,
    { readonly bytes: Uint8Array; readonly contentType: string; readonly type: ObjectPartType }
  >;
  /** Each part's own relationships, their targets absolute part names. */
  readonly relationships: Map<string, RelationshipRecord[]>;
}

export function emptyObjectParts(): ObjectParts {
  return { xml: new Map(), bytes: new Map(), relationships: new Map() };
}

const partCount = (parts: ObjectParts): number => parts.xml.size + parts.bytes.size;

function bytesOf(pkg: OoxmlPackage, name: string): Uint8Array | undefined {
  return pkg.partBytes.get(name) ?? pkg.partBytes.get(name.replace(/^\//, ''));
}

/**
 * Adds the part of bytes at `name` to `into` when it checks. A picture passes the media gate of
 * a pasted picture. Any other part needs its type's content type and signature, and is never
 * opened, read or run.
 */
function admitBytes(
  pkg: OoxmlPackage,
  name: string,
  type: ObjectPartType,
  declared: string,
  into: ObjectParts
): boolean {
  const bytes = bytesOf(pkg, name);
  if (bytes === undefined) {
    return false;
  }

  if (type.contentType === 'image/*') {
    const mime = admittedMediaMime(bytes, declared);
    if (mime === null) {
      return false;
    }

    into.bytes.set(name, { bytes, contentType: mime, type });
    return true;
  }

  // An embedded Office file takes the name Word gives its own content type
  const variant = type.variants?.get(declared);
  if (!variant && declared !== type.contentType.toLowerCase()) {
    return false;
  }

  if (!SIGNATURE_CHECKS[type.signature ?? 'zip'](bytes)) {
    return false;
  }

  const named = variant ? { ...type, stem: variant[0], extension: variant[1] } : type;
  into.bytes.set(name, { bytes, contentType: variant ? declared : type.contentType, type: named });
  return true;
}

/**
 * Adds the XML part at `name` to `into` when it has its type's content type and root element.
 * The part came through the bounded part reader, as every part of a file does.
 */
function admitXml(
  pkg: OoxmlPackage,
  name: string,
  type: ObjectPartType,
  [namespaceUri, localName]: readonly [string, string],
  declared: string,
  into: ObjectParts
): boolean {
  const part = pkg.parts.get(name);

  // Content types compare without case (RFC 2045), as the package reader resolves them
  if (part === undefined || declared !== type.contentType.toLowerCase()) {
    return false;
  }

  if (part.root.namespaceUri !== namespaceUri || part.root.localName !== localName) {
    return false;
  }

  into.xml.set(name, { part, type });
  return true;
}

/**
 * Add to `into` the part a story's relationship names, and every part its own relationships
 * name, in turn. Every part is untrusted: it may hold only relationships of its type's list, to
 * parts inside the package. False, with `into` unchanged, when one of them is missing, not of its
 * type, linked outside the package (a chart linked to a workbook file), or past the bounds: the
 * drawing does not travel.
 */
export function collectObjectParts(
  pkg: OoxmlPackage,
  record: RelationshipRecord,
  into: ObjectParts
): boolean {
  const found = emptyObjectParts();
  const pending: Array<readonly [RelationshipRecord, number]> = [[record, 0]];
  for (let next = pending.pop(); next !== undefined; next = pending.pop()) {
    const [relationship, depth] = next;
    const type = PART_TYPES.get(relationship.type);
    if (type === undefined || relationship.targetMode === 'External' || depth > MAX_DEPTH) {
      return false;
    }

    const target = resolveInternalTarget(relationship.ownerPart, relationship.rawTarget);
    if (!target.ok) return false;
    const name = target.partName;
    if (found.xml.has(name) || found.bytes.has(name)) continue;
    const declared = (resolveContentTypeOf(pkg, name) ?? '').toLowerCase();
    const admitted =
      type.root === null
        ? admitBytes(pkg, name, type, declared, found)
        : admitXml(pkg, name, type, type.root, declared, found);
    if (!admitted) return false;
    if (partCount(found) + partCount(into) > MAX_OBJECT_PARTS) return false;
    const own = relationshipsOf(pkg, name);
    if (own.length === 0) continue;
    if (type.root === null) return false;
    const records: RelationshipRecord[] = [];
    for (const child of own) {
      if (child.targetMode === 'External' || !type.relationships.includes(child.type)) return false;
      const resolved = resolveInternalTarget(child.ownerPart, child.rawTarget);
      if (!resolved.ok) return false;
      records.push({ ...child, rawTarget: resolved.partName });
      pending.push([child, depth + 1]);
    }

    found.relationships.set(name, records);
  }

  for (const [name, entry] of found.xml) {
    into.xml.set(name, entry);
  }

  for (const [name, entry] of found.bytes) {
    into.bytes.set(name, entry);
  }

  for (const [name, records] of found.relationships) {
    into.relationships.set(name, records);
  }

  return true;
}

/**
 * Whether `node` is a diagram data part's `dsp:dataModelExt`, which names the diagram's drawing
 * by a relationship of the STORY part.
 */
function isDrawingExtension(node: OoxmlElement): boolean {
  return node.namespaceUri === DIAGRAM_DRAWING_NS && node.localName === 'dataModelExt';
}

/** Whether `attribute` is a `dsp:dataModelExt`'s `relId`, the story's id for the drawing. */
function isDrawingRelId(attribute: OoxmlAttribute): boolean {
  return attribute.namespaceUri === '' && attribute.localName === 'relId';
}

/** The story relationship ids a diagram's data part names for its drawing. */
function diagramDrawingIds(node: OoxmlNode, out: Set<string>): void {
  if (node.kind === 'textValue') return;
  if (isDrawingExtension(node)) {
    const id = node.attributes.find(isDrawingRelId);
    if (id !== undefined && id.value.length > 0) out.add(id.value);
  }

  for (const child of node.children) {
    diagramDrawingIds(child, out);
  }
}

/**
 * `used` with the ids of the diagram drawings its diagrams' data parts name: the drawing's
 * relationship belongs to the story, and no element of the story names it.
 */
export function withDiagramDrawingIds(
  pkg: OoxmlPackage,
  ownerPart: string,
  used: ReadonlySet<string>
): Set<string> {
  const out = new Set(used);
  const records = relationshipsOf(pkg, ownerPart);
  const drawings = new Set<string>();
  for (const record of records) {
    if (!used.has(record.id) || record.type !== DIAGRAM_DATA_REL) continue;
    const target = resolveInternalTarget(record.ownerPart, record.rawTarget);
    const part = target.ok ? pkg.parts.get(target.partName) : undefined;
    if (part !== undefined) diagramDrawingIds(part.root, drawings);
  }

  for (const record of records) {
    if (drawings.has(record.id) && record.type === DIAGRAM_DRAWING_REL) out.add(record.id);
  }

  return out;
}

/**
 * Whether `node` is a diagram data part's extension (`a:ext`) naming a drawing whose story
 * relationship did not travel. The drawing is a cache that Word draws again from the data and
 * layout (a Word 2007 file has none), so the extension goes and the diagram stays.
 */
function namesLostDrawing(node: OoxmlNode, story: ReadonlyMap<string, string>): boolean {
  if (
    node.kind === 'textValue' ||
    node.namespaceUri !== DRAWINGML_MAIN_NAMESPACE_URI ||
    node.localName !== 'ext'
  ) {
    return false;
  }

  const drawings = new Set<string>();
  diagramDrawingIds(node, drawings);
  return [...drawings].some((id) => !story.has(id));
}

/**
 * A part's tree with its relationship ids mapped (`local`), and a diagram data part's
 * drawing id mapped through the story's map (`story`).
 */
function withMappedIds(
  node: OoxmlNode,
  local: ReadonlyMap<string, string>,
  story: ReadonlyMap<string, string>
): OoxmlNode {
  if (node.kind === 'textValue') return node;
  const extension = isDrawingExtension(node);
  const mapFor = (attribute: OoxmlAttribute): ReadonlyMap<string, string> | null => {
    if (attribute.namespaceUri === R_NS) {
      return local;
    }

    if (extension && isDrawingRelId(attribute)) {
      return story;
    }

    return null;
  };
  const attributes = node.attributes.map((attribute) => {
    const mapped = mapFor(attribute)?.get(attribute.value);
    return mapped === undefined ? attribute : { ...attribute, value: mapped };
  });
  const children = node.children
    .filter((child) => !namesLostDrawing(child, story))
    .map((child) => withMappedIds(child, local, story));
  return { ...node, attributes, children } as OoxmlElement;
}

/** Hands each part of `parts` to the copy's package writer, under its own name. */
export function writeObjectParts(
  parts: ObjectParts,
  sink: {
    readonly xml: (name: string, contentType: string, root: OoxmlElement) => void;
    readonly bytes: (name: string, bytes: Uint8Array, contentType: string) => void;
    readonly relationships: (owner: string, records: readonly RelationshipRecord[]) => void;
  }
): void {
  for (const [name, { part, type }] of parts.xml) {
    sink.xml(name, type.contentType, part.root);
  }

  for (const [name, { bytes, contentType }] of parts.bytes) {
    sink.bytes(name, bytes, contentType);
  }

  for (const [owner, records] of parts.relationships) {
    sink.relationships(owner, records);
  }
}

/** Total bytes of the byte parts, for the copy's size tiers. */
export function objectBytes(parts: ObjectParts): number {
  let total = 0;
  for (const { bytes } of parts.bytes.values()) {
    total += bytes.byteLength;
  }

  return total;
}

/**
 * The extension a pasted part of bytes takes: its type's, else the source's own when it is a
 * plain one, else its picture type's.
 */
function byteExtension(type: ObjectPartType, source: string, contentType: string): string {
  if (type.extension !== '') {
    return type.extension;
  }

  const dot = source.lastIndexOf('.');
  const sourceExtension = dot === -1 ? '' : source.slice(dot + 1).toLowerCase();
  if (/^[a-z0-9]{1,8}$/.test(sourceExtension)) {
    return sourceExtension;
  }

  return contentType.slice('image/'.length).replace(/[^a-z0-9]/g, '') || 'bin';
}

/**
 * New part names of Word's form (`/word/charts/chart2.xml`), free in `pkg`, for every part of
 * `parts`.
 */
function newNames(pkg: OoxmlPackage, parts: ObjectParts): Map<string, string> {
  const taken = new Set<string>();
  for (const name of pkg.parts.keys()) {
    taken.add(partNameKey(name));
  }

  for (const name of pkg.partBytes.keys()) {
    taken.add(partNameKey(name.startsWith('/') ? name : `/${name}`));
  }

  const numbers = new Map<string, Generator<number>>();
  const out = new Map<string, string>();
  const assignName = (source: string, type: ObjectPartType, extension: string): void => {
    let candidates = numbers.get(type.stem);
    if (candidates === undefined) {
      candidates = partNumberCandidates(Number.MAX_SAFE_INTEGER);
      numbers.set(type.stem, candidates);
    }

    for (;;) {
      const candidate = `${type.stem}${candidates.next().value}.${extension}`;
      if (taken.has(partNameKey(candidate))) continue;
      taken.add(partNameKey(candidate));
      out.set(source, candidate);
      return;
    }
  };
  for (const [source, { type }] of parts.xml) {
    assignName(source, type, type.extension);
  }

  for (const [source, { type, contentType }] of parts.bytes) {
    assignName(source, type, byteExtension(type, source, contentType));
  }

  return out;
}

/**
 * `pkg` with a pasted XML part written under its new `name`: its own relationships point at the
 * new names, and a diagram data part's drawing id maps through `storyIds`. Null when a write
 * fails.
 */
function withPastedXmlPart(
  pkg: OoxmlPackage,
  name: string,
  { part, type }: { readonly part: OoxmlPart; readonly type: ObjectPartType },
  relationships: readonly RelationshipRecord[],
  names: ReadonlyMap<string, string>,
  storyIds: ReadonlyMap<string, string>
): OoxmlPackage | null {
  let out = pkg;
  const local = new Map<string, string>();
  if (relationships.length > 0) {
    const added = withRelationships(
      withRelationshipsPartFor(out, name),
      name,
      relationships.map(
        (record) => [record.type, relativeTarget(name, names.get(record.rawTarget)!)] as const
      )
    );
    if (!added.ok) {
      return null;
    }

    out = added.pkg;
    relationships.forEach((record, index) => local.set(record.id, added.ids[index]!));
  }

  const root = withMappedIds(part.root, local, storyIds) as OoxmlElement;
  const read = readOoxmlPart(serializeOoxmlPart({ ...part, root }), {
    name,
    contentType: type.contentType,
  });
  if (!read.ok) {
    return null;
  }

  const written = withNewPart(out, name, read.part.root, type.contentType);
  return written === out ? null : written;
}

/**
 * Paste the parts of the charts, diagrams, embedded objects and ink that `records` (a fragment
 * story's relationships) name into `pkg`, under new names, with new relationships from
 * `targetOwner` to them. A relationship whose parts do not all check is dropped (its drawing
 * goes); a failed write drops them all and leaves `pkg` as it was.
 */
export function mergeObjectParts(
  target: OoxmlPackage,
  fragment: OoxmlPackage,
  records: readonly RelationshipRecord[],
  targetOwner: string
): {
  readonly pkg: OoxmlPackage;
  readonly relIdMap: Map<string, string>;
  readonly dropRelIds: Set<string>;
} {
  const relIdMap = new Map<string, string>();
  const dropRelIds = new Set<string>();
  if (records.length === 0) return { pkg: target, relIdMap, dropRelIds };
  const parts = emptyObjectParts();
  const kept: RelationshipRecord[] = [];
  for (const record of records) {
    if (collectObjectParts(fragment, record, parts)) {
      kept.push(record);
    } else {
      dropRelIds.add(record.id);
    }
  }

  const failed = () => {
    for (const record of kept) {
      dropRelIds.add(record.id);
    }

    return { pkg: target, relIdMap: new Map<string, string>(), dropRelIds };
  };
  if (kept.length === 0) return { pkg: target, relIdMap, dropRelIds };
  const names = newNames(target, parts);
  const storyTargets = kept.map((record) => {
    const resolved = resolveInternalTarget(record.ownerPart, record.rawTarget);
    return resolved.ok ? names.get(resolved.partName) : undefined;
  });
  if (storyTargets.some((name) => name === undefined)) return failed();
  const story = withRelationships(
    withRelationshipsPartFor(target, targetOwner),
    targetOwner,
    kept.map(
      (record, index) => [record.type, relativeTarget(targetOwner, storyTargets[index]!)] as const
    )
  );
  if (!story.ok) return failed();
  kept.forEach((record, index) => relIdMap.set(record.id, story.ids[index]!));
  let pkg = withBinaryParts(
    story.pkg,
    [...parts.bytes].map(([source, { bytes, contentType }]) => ({
      partName: names.get(source)!,
      bytes,
      contentType,
    }))
  );
  for (const [source, entry] of parts.xml) {
    const written = withPastedXmlPart(
      pkg,
      names.get(source)!,
      entry,
      parts.relationships.get(source) ?? [],
      names,
      relIdMap
    );
    if (written === null) {
      return failed();
    }

    pkg = written;
  }

  return { pkg, relIdMap, dropRelIds };
}
