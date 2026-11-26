// The field table of the saved item list (checkpoint version 6, design D14): every field of every
// type of the model, stored with its codec or rebuilt by the open. Each table covers every key of
// its type, so a new field fails the type check here until it has an entry. The writer, the reader
// and the compare all go through this one table.

import {
  IMAGE_WRAP_TARGETS,
  type DrawingLocksInput,
  type DrawingPositionInput,
  type SourceCrop,
} from '../../../store/package/drawing-projection.ts';
import type { OoxmlProperty, TabStopWrite } from '../../../store/store/tree-op-types.ts';
import { at, Invalid, isForbiddenKey, plainFields } from '../change/read.ts';
import { sha256 } from '../change/resources.ts';
import { FRAMES_H, FRAMES_V } from '../change/kinds/pictures-read.ts';
import { href, readNamespace } from '../change/values.ts';
import type { OrderedFragmentPartType } from '../change/fragment-names-objects.ts';
import type { OrderedComment, OrderedCommentRange } from './comments.ts';
import type { OrderedObjectOp, OrderedObjectRecord } from './objects.ts';
import {
  ORDERED_MEDIA_TYPES,
  type OrderedFileMedia,
  type OrderedFileStore,
  type OrderedListBullet,
  type OrderedListLevel,
  type OrderedRelationshipResource,
  type OrderedRelationshipType,
  type OrderedResource,
  type OrderedResources,
  type OrderedSettingResource,
  type OrderedStoryRole,
} from './resources.ts';
import type { OrderedSectionOp, OrderedSectionRecord } from './sections.ts';
import { createOrderedItem } from './document.ts';
import {
  after,
  attributes,
  bool,
  choice,
  count,
  element,
  exactly,
  id,
  ID_MAX,
  int,
  item,
  list,
  map,
  maybe,
  node,
  nullable,
  num,
  part,
  person,
  readXmlString,
  rebuilt,
  record,
  set,
  sortedMap,
  stored,
  text,
  variants,
  xmlString,
  type SnapshotCodec,
  type SnapshotFields,
} from './snapshot-codec.ts';
import type {
  OrderedAttribution,
  OrderedContainerToken,
  OrderedCounters,
  OrderedDocument,
  OrderedElementRecord,
  OrderedFieldPiece,
  OrderedFormatRecord,
  OrderedItem,
  OrderedItemAddress,
  OrderedItemKind,
  OrderedNoteProperties,
  OrderedNotePropertiesOp,
  OrderedParagraphOp,
  OrderedParagraphOpRecord,
  OrderedPartRevision,
  OrderedPropertyRecord,
  OrderedPropertyTarget,
  OrderedRevisionIds,
  OrderedStory,
  OrderedStyleWrite,
  OrderedSuggestion,
  OrderedUnitRange,
  OrderedWithin,
} from './types.ts';

/** Every value of a string union, once: a value the union gains fails the type check. */
const valuesOf = <T extends string>(values: Record<T, true>): T[] => Object.keys(values) as T[];

/** A name or a text that may be empty. */
const name = xmlString(0, ID_MAX);

/** A coordinate, a twips value or an offset in a write: any safe integer. */
const signed = int(Number.MIN_SAFE_INTEGER);

/** An item's model text: characters, or the one character that stands for an atom. */
const unitText: SnapshotCodec<string> = {
  write: (value) => value,
  read: (json, path) => {
    if (typeof json !== 'string') {
      throw new Invalid('invalid-type', path, 'not a string');
    }

    return json;
  },
  diff: (a, b, path) => (a === b ? null : path),
};

/** A list level, 0 to 8. */
const level = int(0, 8);

const ids = list(id);

/** String keys to string values (a property's attributes, the document properties). */
const strings: SnapshotCodec<Readonly<Record<string, string>>> = {
  write: (value) => value,
  read: (json, path) => {
    const f = plainFields(json, path);
    return Object.fromEntries(
      Object.keys(json as object).map((key) => {
        if (isForbiddenKey(key)) {
          throw new Invalid('forbidden-key', at(path, key));
        }

        return [
          readXmlString(key, path, 1, ID_MAX),
          readXmlString(f.get(key), at(path, key), 0, ID_MAX),
        ];
      })
    );
  },
  diff: (a, b, path) => {
    const keys = Object.keys(a);
    const same = keys.length === Object.keys(b).length && keys.every((key) => a[key] === b[key]);
    return same ? null : path;
  },
};

const PROPERTY = record<OoxmlProperty>({
  localName: stored(id),
  attributes: { stored: strings, optional: true },
});

const properties = list(PROPERTY);

/** One of two fields, by which one the object has. */
function either<T extends object>(keys: readonly [string, string]): SnapshotCodec<T> {
  return {
    write: (value) => value,
    read: (json, path) => {
      const f = plainFields(json, path);
      const held = keys.filter((key) => f.has(key));
      if (held.length !== 1 || Object.keys(json as object).length !== 1) {
        throw new Invalid('invalid-value', path, `not one of ${keys.join(', ')}`);
      }

      const key = held[0]!;
      return { [key]: readXmlString(f.get(key), at(path, key), 1, ID_MAX) } as T;
    },
    diff: (a, b, path) => (JSON.stringify(a) === JSON.stringify(b) ? null : path),
  };
}

const STYLE_WRITE = either<OrderedStyleWrite>(['set', 'remove']);

// ---- items ----

const KINDS = valuesOf<OrderedItemKind>({
  text: true,
  atom: true,
  zero: true,
  mark: true,
  open: true,
  close: true,
});

const TOKENS = valuesOf<OrderedContainerToken>({
  hyperlink: true,
  sdt: true,
  smartTag: true,
  customXml: true,
  dir: true,
  bdo: true,
  fldSimple: true,
  tbl: true,
  tr: true,
  tc: true,
  sdtBlock: true,
  customXmlBlock: true,
  note: true,
});

const TARGETS = valuesOf<OrderedPropertyTarget>({
  pPr: true,
  numPr: true,
  tblPr: true,
  tblGrid: true,
  tblPrEx: true,
  trPr: true,
  tcPr: true,
  sdtPr: true,
  sectPr: true,
});

type PropertiesFormat = Extract<OrderedFormatRecord, { readonly kind: 'properties' }>;
type FaceFormat = Extract<OrderedFormatRecord, { readonly kind: 'face' }>;

export const PROPERTIES_FORMAT: SnapshotFields<PropertiesFormat> = {
  kind: stored(exactly('properties')),
  id: stored(id),
  suggestion: stored(id),
  seq: stored(count),
  properties: stored(properties),
  remove: maybe(ids),
  style: maybe(STYLE_WRITE),
  replaced: maybe(id),
  replacedEarlier: maybe(ids),
  over: maybe(ids),
  untracked: maybe(bool),
  change: maybe(id),
};

export const FACE_FORMAT: SnapshotFields<FaceFormat> = {
  kind: stored(exactly('face')),
  id: stored(id),
  suggestion: stored(id),
  seq: stored(count),
  face: stored(nullable(element)),
  fileRecord: maybe(element),
  replaced: maybe(id),
  replacedEarlier: maybe(ids),
};

const FORMAT_RECORD = variants<OrderedFormatRecord, 'kind'>('kind', {
  properties: PROPERTIES_FORMAT,
  face: FACE_FORMAT,
});

export const ELEMENT_RECORD: SnapshotFields<OrderedElementRecord> = {
  kind: stored(exactly('element')),
  suggestion: stored(id),
  seq: stored(count),
  target: stored(choice(TARGETS)),
  element: stored(nullable(element)),
  fileRecord: maybe(element),
  recordUnchanged: maybe(bool),
  untracked: maybe(bool),
  change: maybe(id),
  acceptedOnly: maybe(bool),
  settled: maybe(bool),
  tableCells: maybe(ids),
  replaced: maybe(id),
  replacedEarlier: maybe(ids),
  over: maybe(ids),
  merge: maybe(ids),
};

const TAB_STOP = record<TabStopWrite>({
  positionTwips: stored(signed),
  alignment: stored(choice(['left', 'center', 'right', 'decimal', 'bar'])),
  leader: maybe(choice(['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'])),
});

const PARAGRAPH_OP = variants<OrderedParagraphOp, 'op'>('op', {
  setParagraphProperties: {
    op: stored(exactly('setParagraphProperties')),
    properties: stored(properties),
    remove: maybe(ids),
  },
  setListLevel: { op: stored(exactly('setListLevel')), level: stored(level) },
  setListNumbering: {
    op: stored(exactly('setListNumbering')),
    numId: stored(nullable(id)),
    level: maybe(level),
  },
  setParagraphTabStops: {
    op: stored(exactly('setParagraphTabStops')),
    stops: stored(list(TAB_STOP)),
    inForcePositionsTwips: maybe(list(signed)),
  },
});

export const PARAGRAPH_OP_RECORD: SnapshotFields<OrderedParagraphOpRecord> = {
  kind: stored(exactly('paragraphOp')),
  suggestion: stored(id),
  seq: stored(count),
  target: stored(exactly('pPr')),
  op: stored(PARAGRAPH_OP),
  untracked: maybe(bool),
  replaced: maybe(id),
  replacedEarlier: maybe(ids),
  over: maybe(ids),
};

export const NOTE_PROPERTIES: SnapshotFields<OrderedNoteProperties> = {
  position: maybe(id),
  numFmt: maybe(id),
  numStart: maybe(signed),
  numRestart: maybe(id),
};

const NOTE_PROPERTIES_OP: SnapshotFields<OrderedNotePropertiesOp> = {
  op: stored(exactly('setNoteProperties')),
  footnote: maybe(record(NOTE_PROPERTIES)),
  endnote: maybe(record(NOTE_PROPERTIES)),
};

const SECTION_OP = variants<OrderedSectionOp, 'op'>('op', {
  page: {
    op: stored(exactly('page')),
    pageWidthTwips: maybe(signed),
    pageHeightTwips: maybe(signed),
    orientation: maybe(choice(['portrait', 'landscape'])),
    marginTopTwips: maybe(signed),
    marginRightTwips: maybe(signed),
    marginBottomTwips: maybe(signed),
    marginLeftTwips: maybe(signed),
  },
  furniture: {
    op: stored(exactly('furniture')),
    titlePage: maybe(bool),
    headerDistanceTwips: maybe(signed),
    footerDistanceTwips: maybe(signed),
  },
  reference: {
    op: stored(exactly('reference')),
    kind: stored(choice(['header', 'footer'])),
    variant: stored(choice(['default', 'first', 'even'])),
    id: stored(nullable(id)),
  },
  type: { op: stored(exactly('type')), type: stored(choice(['nextPage', 'continuous'])) },
  mark: { op: stored(exactly('mark')), sectPr: stored(element) },
  setNoteProperties: NOTE_PROPERTIES_OP,
});

export const SECTION_RECORD: SnapshotFields<OrderedSectionRecord> = {
  kind: stored(exactly('sectionOp')),
  suggestion: stored(id),
  seq: stored(count),
  target: stored(exactly('sectPr')),
  op: stored(SECTION_OP),
  untracked: maybe(bool),
  keptOnReject: maybe(bool),
  replaced: maybe(id),
  replacedEarlier: maybe(ids),
  over: maybe(ids),
};

const OBJECT_OP = variants<OrderedObjectOp, 'op'>('op', {
  resize: { op: stored(exactly('resize')), cx: stored(count), cy: stored(count) },
  crop: {
    op: stored(exactly('crop')),
    crop: stored(
      record<SourceCrop>({
        left: stored(num),
        top: stored(num),
        right: stored(num),
        bottom: stored(num),
      })
    ),
  },
  position: {
    op: stored(exactly('position')),
    position: stored(
      record<DrawingPositionInput>({
        mode: maybe(choice(['frame', 'simple'])),
        horizontalEmu: maybe(signed),
        verticalEmu: maybe(signed),
        relativeToH: maybe(choice(FRAMES_H)),
        relativeToV: maybe(choice(FRAMES_V)),
      })
    ),
  },
  wrap: { op: stored(exactly('wrap')), wrap: stored(choice(IMAGE_WRAP_TARGETS)) },
  metadata: {
    op: stored(exactly('metadata')),
    title: stored(text),
    description: stored(text),
    link: maybe(nullable(id)),
  },
  locks: {
    op: stored(exactly('locks')),
    locks: stored(
      record<DrawingLocksInput>({
        select: maybe(bool),
        move: maybe(bool),
        resize: maybe(bool),
        changeAspect: maybe(bool),
      })
    ),
  },
  transform: {
    op: stored(exactly('transform')),
    action: stored(choice(['rotateCW', 'rotateCCW', 'flipH', 'flipV'])),
  },
  image: { op: stored(exactly('image')), relationship: stored(id) },
});

export const OBJECT_RECORD: SnapshotFields<OrderedObjectRecord> = {
  kind: stored(exactly('object')),
  suggestion: stored(id),
  seq: stored(count),
  target: stored(exactly('object')),
  op: stored(OBJECT_OP),
  replaced: maybe(id),
};

const PROPERTY_RECORD = variants<OrderedPropertyRecord, 'kind'>('kind', {
  element: ELEMENT_RECORD,
  paragraphOp: PARAGRAPH_OP_RECORD,
  sectionOp: SECTION_RECORD,
  object: OBJECT_RECORD,
});

export const WITHIN: SnapshotFields<OrderedWithin> = { by: stored(ids), dels: stored(ids) };

export const PIECE: SnapshotFields<OrderedFieldPiece> = {
  node: stored(node),
  sourceRun: maybe(id),
  recordFaces: maybe(map(id, nullable(element))),
  recordFiles: maybe(map(id, element)),
};

export const ITEM: SnapshotFields<OrderedItem> = {
  leaf: rebuilt('`addOrderedStory` puts the item in its story tree'),
  space: stored(id),
  offset: stored(count),
  length: stored(int(1)),
  kind: stored(choice(KINDS)),
  text: stored(unitText),
  node: stored(nullable(node)),
  textShell: maybe(element),
  pieces: maybe(list(record(PIECE))),
  face: stored(nullable(element)),
  run: stored(name),
  inRun: maybe(bool),
  bare: maybe(bool),
  block: maybe(bool),
  instr: maybe(bool),
  rebuilt: maybe(bool),
  prompt: maybe(bool),
  token: maybe(choice(TOKENS)),
  pair: maybe(id),
  home: maybe(id),
  unitCount: maybe(count),
  fieldSlot: maybe(record({ field: stored(id), before: stored(count) })),
  moveRange: maybe(bool),
  by: stored(nullable(id)),
  dels: stored(ids),
  fmts: stored(list(FORMAT_RECORD)),
  cancelledFmts: stored(list(FORMAT_RECORD)),
  props: stored(list(PROPERTY_RECORD)),
  takenBackLendings: stored(list(record(ELEMENT_RECORD))),
  move: stored(nullable(id)),
  retracted: stored(bool),
  taken: maybe(nullable(id)),
  gone: stored(bool),
  quiet: stored(bool),
  within: stored(record(WITHIN)),
};

// ---- suggestions, comments and the document ----

export const ATTRIBUTION: SnapshotFields<OrderedAttribution> = {
  author: stored(name),
  person: stored(nullable(person)),
  date: maybe(name),
  attributes: maybe(attributes),
};

export const SUGGESTION: SnapshotFields<OrderedSuggestion> = {
  id: stored(id),
  attribution: stored(record(ATTRIBUTION)),
  status: stored(choice(['pending', 'accepted', 'rejected'])),
  seq: stored(count),
  version: stored(count),
  direct: maybe(bool),
};

/**
 * What a new item holds past its address and kind (`createOrderedItem`): an item's field that
 * holds it is left out, and an absent one reads as it.
 */
const { space: _space, kind: _kind, ...NEW_ITEM } = createOrderedItem({ space: '', kind: 'text' });

/** An item, made as a new item is made. */
const ITEM_CODEC = after(record(ITEM, NEW_ITEM), createOrderedItem);

const ITEMS = list(ITEM_CODEC);

/** A story's tree, as its items in order. */
const TREE: SnapshotCodec<OrderedStory['tree']> = {
  write: (tree, w) => ITEMS.write(tree.toArray(), w),
  read: (json, path, r) => r.tree(ITEMS.read(json, path, r)),
  diff: (a, b, path) => ITEMS.diff(a.toArray(), b.toArray(), path),
};

export const STORY: SnapshotFields<OrderedStory> = {
  name: stored(id),
  part: stored(part),
  tree: stored(TREE),
  spaces: rebuilt('`addOrderedStory` indexes the items by space'),
};

export const ADDRESS: SnapshotFields<OrderedItemAddress> = {
  space: stored(id),
  offset: stored(count),
};

export const UNIT_RANGE: SnapshotFields<OrderedUnitRange> = {
  story: stored(id),
  space: stored(id),
  from: stored(count),
  to: stored(count),
};

export const COUNTERS: SnapshotFields<OrderedCounters> = {
  suggestion: stored(count),
  insertion: stored(count),
  seq: stored(count),
};

export const REVISION_IDS: SnapshotFields<OrderedRevisionIds> = {
  byRole: stored(map(id, id)),
  attributesByRole: stored(map(id, attributes)),
  fileIds: stored(set(id)),
  reattributed: stored(map(id, record(ATTRIBUTION))),
  listed: stored(map(id, ids)),
};

export const PART_REVISION: SnapshotFields<OrderedPartRevision> = {
  part: stored(id),
  element: stored(id),
  revision: stored(record({ id: stored(id), author: stored(name), date: maybe(id) })),
  styleId: maybe(id),
};

export const COMMENT: SnapshotFields<OrderedComment> = {
  id: stored(id),
  author: stored(name),
  initials: stored(nullable(name)),
  date: stored(nullable(id)),
  paraIds: stored(ids),
  paraId: stored(nullable(id)),
  parent: stored(nullable(id)),
  done: stored(bool),
  element: stored(element),
  person: stored(nullable(person)),
};

export const COMMENT_RANGE: SnapshotFields<OrderedCommentRange> = {
  story: stored(id),
  start: stored(nullable(record(ADDRESS))),
  end: stored(nullable(record(ADDRESS))),
  reference: stored(nullable(record(ADDRESS))),
};

// ---- resources ----

const SHA256: SnapshotCodec<string> = {
  write: (value) => value,
  read: (json, path) => sha256(json, path),
  diff: (a, b, path) => (a === b ? null : path),
};

export const FILE_MEDIA: SnapshotFields<OrderedFileMedia> = {
  part: stored(id),
  contentType: stored(choice(ORDERED_MEDIA_TYPES)),
  sha256: stored(SHA256),
  size: stored(int(1)),
  entry: stored(count),
};

export const FILE_STORE: SnapshotFields<OrderedFileStore> = {
  storeItemId: stored(id),
  root: stored(id),
  boundNodes: stored(set(id)),
};

const ROLES = valuesOf<OrderedStoryRole>({
  header: true,
  footer: true,
  footnotes: true,
  endnotes: true,
});

const PART_TYPES = valuesOf<OrderedFragmentPartType>({
  chart: true,
  chartEx: true,
  chartStyle: true,
  chartColorStyle: true,
  chartUserShapes: true,
  themeOverride: true,
  diagramData: true,
  diagramLayout: true,
  diagramQuickStyle: true,
  diagramColors: true,
  diagramDrawing: true,
});

const RELATIONSHIP_TYPES = valuesOf<OrderedRelationshipType>({
  image: true,
  hyperlink: true,
  header: true,
  footer: true,
  footnotes: true,
  endnotes: true,
  package: true,
  oleObject: true,
  ink: true,
  ...Object.fromEntries(PART_TYPES.map((type) => [type, true] as const)),
} as Record<OrderedRelationshipType, true>);

/** A relationship's target: a resource's key, or an address outside, checked as a link is. */
const RELATIONSHIP_TARGET: SnapshotCodec<OrderedRelationshipResource['target']> = {
  write: (value) => value,
  read: (json, path) => {
    const f = plainFields(json, path);
    if (Object.keys(json as object).length !== 1) {
      throw new Invalid('invalid-value', path, 'not one target');
    }

    if (f.has('resource')) {
      return { resource: readXmlString(f.get('resource'), at(path, 'resource'), 1, ID_MAX) };
    }

    if (f.has('external')) {
      return { external: href(f.get('external'), at(path, 'external')) };
    }

    throw new Invalid('invalid-value', path, 'not a target');
  },
  diff: (a, b, path) => (JSON.stringify(a) === JSON.stringify(b) ? null : path),
};

export const LIST_LEVEL_VALUE: SnapshotFields<OrderedListLevel> = {
  level: stored(level),
  start: stored(signed),
  format: stored(id),
  text: stored(name),
  justification: stored(choice(['left', 'center', 'right'])),
  indent: stored(record({ left: stored(signed), hanging: stored(signed) })),
  font: maybe(id),
};

export const LIST_BULLET: SnapshotFields<OrderedListBullet> = {
  id: stored(id),
  element: stored(element),
  images: stored(list(record({ id: stored(id), media: stored(id) }))),
};

const SETTING_OP = variants<OrderedSettingResource['op'], 'op'>('op', {
  setDocumentProtection: {
    op: stored(exactly('setDocumentProtection')),
    enforce: stored(bool),
  },
  setSectionFurnitureOptions: {
    op: stored(exactly('setSectionFurnitureOptions')),
    evenAndOddHeaders: stored(bool),
  },
  setDocumentProperties: {
    op: stored(exactly('setDocumentProperties')),
    values: stored(strings),
  },
  setNoteProperties: NOTE_PROPERTIES_OP,
});

/** What every resource has. */
const RESOURCE_BASE = { key: stored(id), suggestion: stored(id) };

type ResourceFields<K extends OrderedResource['kind']> = SnapshotFields<
  Extract<OrderedResource, { readonly kind: K }>
>;

export const RESOURCES_BY_KIND: { readonly [K in OrderedResource['kind']]: ResourceFields<K> } = {
  media: {
    kind: stored(exactly('media')),
    ...RESOURCE_BASE,
    part: stored(id),
    contentType: stored(choice(ORDERED_MEDIA_TYPES)),
    sha256: stored(SHA256),
    size: stored(count),
  },
  story: {
    kind: stored(exactly('story')),
    ...RESOURCE_BASE,
    part: stored(id),
    role: stored(choice(ROLES)),
    relationshipsOf: maybe(id),
  },
  part: {
    kind: stored(exactly('part')),
    ...RESOURCE_BASE,
    part: stored(id),
    type: stored(choice(PART_TYPES)),
    root: stored(element),
  },
  relationship: {
    kind: stored(exactly('relationship')),
    ...RESOURCE_BASE,
    owner: stored(id),
    id: stored(id),
    type: stored(choice(RELATIONSHIP_TYPES)),
    target: stored(RELATIONSHIP_TARGET),
  },
  list: {
    kind: stored(exactly('list')),
    ...RESOURCE_BASE,
    numId: stored(id),
    abstractNumId: stored(id),
    levels: stored(list(record(LIST_LEVEL_VALUE))),
    definition: maybe(record({ abstract: stored(element), overrides: stored(list(element)) })),
    bullets: maybe(list(record(LIST_BULLET))),
  },
  'list-level': {
    kind: stored(exactly('list-level')),
    ...RESOURCE_BASE,
    numId: stored(id),
    level: stored(level),
    format: stored(choice(['bullet', 'ordered'])),
  },
  style: {
    kind: stored(exactly('style')),
    ...RESOURCE_BASE,
    styleId: stored(id),
    type: stored(choice(['paragraph', 'character', 'table'])),
    name: stored(name),
    basedOn: maybe(id),
    next: maybe(id),
    link: maybe(id),
    uiPriority: maybe(signed),
    unhideWhenUsed: maybe(bool),
    paragraph: stored(properties),
    run: stored(properties),
    element: maybe(element),
  },
  setting: { kind: stored(exactly('setting')), ...RESOURCE_BASE, op: stored(SETTING_OP) },
  'custom-xml': {
    kind: stored(exactly('custom-xml')),
    ...RESOURCE_BASE,
    namespaceUri: stored(id),
    root: stored(id),
    storeItemId: stored(id),
    nodeId: stored(id),
    label: stored(text),
    data: stored(text),
    control: stored(item),
  },
};

export const RESOURCES: SnapshotFields<OrderedResources> = {
  fileParts: stored(set(id)),
  fileRelationships: stored(map(id, set(id))),
  fileLists: stored(set(id)),
  fileStores: stored(map(id, record(FILE_STORE))),
  fileMedia: rebuilt('the media `ordered/base.docx` names by reference (`recordFileMedia`)'),
  added: stored(map(id, variants<OrderedResource, 'kind'>('kind', RESOURCES_BY_KIND))),
};

/** The person of each namespace that typed, by its namespace. */
export const NAMESPACE: SnapshotCodec<string> = {
  write: (value) => value,
  read: (json, path) => readNamespace(json, path),
  diff: (a, b, path) => (a === b ? null : path),
};

/** A change id's counter: from 1, at most 15 digits. */
const CHANGE_COUNTER = int(1, 10 ** 15 - 1);

export const DOCUMENT: SnapshotFields<OrderedDocument> = {
  file: rebuilt("the checkpoint's `file`: the number its loaded spaces hold, kept at every save"),
  // Read first: the trees sum by the suggestions' status, and later fields name the items
  suggestions: stored(after(map(id, record(SUGGESTION)), (read, r) => r.suggestions(read))),
  stories: stored(
    map(
      id,
      after(record(STORY), (read, r) => r.story(read))
    )
  ),
  labelled: rebuilt('`addOrderedStory` indexes the items by their labels'),
  ownEdited: rebuilt("`addOrderedStory` indexes the items by their author's own edits"),
  added: stored(map(id, list(record(UNIT_RANGE)))),
  counters: stored(record(COUNTERS)),
  runShells: stored(map(id, element)),
  originals: stored(map(id, node)),
  revisionIds: stored(record(REVISION_IDS)),
  placeholders: stored(map(name, list(node))),
  movesWithRanges: stored(set(id)),
  composeIds: stored(record({ next: stored(int(1)) })),
  paraIds: stored(set(id)),
  commentIds: stored(set(id)),
  comments: stored(map(id, record(COMMENT))),
  commentRanges: stored(map(id, record(COMMENT_RANGE))),
  numbering: stored(nullable(set(id))),
  // The file media are those `ordered/base.docx` names, which the open records
  resources: stored(after(record(RESOURCES), (read) => ({ ...read, fileMedia: new Map() }))),
  partRevisions: stored(map(id, record(PART_REVISION))),
  namespaces: stored(map(NAMESPACE, person)),
  lastCounters: stored(sortedMap(NAMESPACE, CHANGE_COUNTER)),
  undo: rebuilt('null: no edit runs while a room saves or opens'),
};
