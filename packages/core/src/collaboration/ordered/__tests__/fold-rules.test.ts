// How a fold keeps each fact of the room's model (design D14). A room writes its document as a
// checkpoint and the next opening reads it back, so each field of the model needs a rule: read
// back from the Word file (`file`), kept in the checkpoint's manifest (`manifest`), made again
// from other facts (`derived`), or not needed after a fold (`dropped`). A field with no honest
// rule is `open`, with its backlog row. Each table covers every key of its type, so a new field
// fails the type check here until it has a rule. `fold-check.ts` and the shape tests compare the
// room live and after a fold; this file only names the rules and checks the manifest's sections.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { readOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { CHECKPOINT_MANIFEST, readCheckpoint, type CheckpointRead } from '../model/checkpoint.ts';
import type { OrderedComment, OrderedCommentRange } from '../model/comments.ts';
import type { OrderedObjectRecord } from '../model/objects.ts';
import type {
  OrderedFileMedia,
  OrderedFileStore,
  OrderedResource,
  OrderedResources,
} from '../model/resources.ts';
import type { OrderedSectionRecord } from '../model/sections.ts';
import type {
  OrderedAttribution,
  OrderedCounters,
  OrderedDocument,
  OrderedElementRecord,
  OrderedFieldPiece,
  OrderedFormatRecord,
  OrderedItem,
  OrderedParagraphOpRecord,
  OrderedPartRevision,
  OrderedRevisionIds,
  OrderedStory,
  OrderedSuggestion,
  OrderedWithin,
} from '../model/types.ts';
import { addressOf, apply, APPLY_FILE, DIRECT, EDIT, opens } from './change-apply-fixture.ts';
import {
  PASTED_LIST,
  ROOM_FACTS,
  roomBlobs,
  SPLIT_RUN,
  STYLE_RECORDS,
  TABLE_RECORDS,
  TYPED_BARE,
  TYPED_RUNS,
  type CheckpointRoom,
} from './checkpoint-fixture.ts';
import { checkpointBytes } from './fold-check.ts';

/** A read checkpoint. */
type CheckpointOpened = Extract<CheckpointRead, { readonly ok: true }>;

/** The manifest's sections of room facts, each with the number of entries a read gives back. */
const SECTIONS = {
  media: (read: CheckpointOpened) => read.media.length,
  suggestions: (read: CheckpointOpened) => read.facts.suggestions.length,
  comments: (read: CheckpointOpened) => read.facts.commentPersons.size,
  takeovers: (read: CheckpointOpened) => read.facts.takeovers.length,
  namespaces: (read: CheckpointOpened) => read.facts.namespacePersons.size,
  untracked: (read: CheckpointOpened) => read.facts.untracked.length,
  changeRecords: (read: CheckpointOpened) => read.facts.changeRecords.length,
  lists: (read: CheckpointOpened) => read.facts.lists.length,
  styles: (read: CheckpointOpened) => read.facts.styles.length,
};

type ManifestSection = keyof typeof SECTIONS;

/** The manifest's fields that are no section of room facts. */
const HEADER: ReadonlySet<string> = new Set(['format', 'version']);

/** How a fold keeps one field. `how` says which load step, section or facts. */
type FoldRule =
  | { readonly rule: 'file'; readonly how: string }
  | {
      readonly rule: 'manifest';
      readonly section: ManifestSection;
      /** The field in the section's entries, dotted for a field of a nested entry. */
      readonly field: string;
      readonly how: string;
    }
  | { readonly rule: 'derived'; readonly how: string }
  | { readonly rule: 'dropped'; readonly how: string }
  | { readonly rule: 'open'; readonly row: string; readonly how: string };

const file = (how: string): FoldRule => ({ rule: 'file', how });
const manifest = (section: ManifestSection, field: string, how: string): FoldRule => ({
  rule: 'manifest',
  section,
  field,
  how,
});
const derived = (how: string): FoldRule => ({ rule: 'derived', how });
const dropped = (how: string): FoldRule => ({ rule: 'dropped', how });
const open = (row: string, how: string): FoldRule => ({ rule: 'open', row, how });

/** Every key of `T`, each with its rule. */
type Rules<T> = Record<keyof T, FoldRule>;

// Rules that many records share
const SUGGESTION = manifest('suggestions', 'revisions', "the change record's `w:id` names it");
const RECORD_SEQ = derived('the load numbers records as it reads them, taken ones first');
const TAKEN = manifest('takeovers', 'taken', 'the change record that took it lists it');
const UNDO_ONLY = dropped(
  'only the undo of the change that made it reads it; a fold keeps no undo'
);

const ITEM: Rules<OrderedItem> = {
  leaf: derived('the tree sets it when `addOrderedStory` builds the story'),
  space: file('load/state.ts names a loaded space by the node id; addresses are new (D8)'),
  offset: file('load/state.ts, with the space'),
  length: file('load/runs.ts: the units of the loaded text'),
  kind: file('load/paragraph.ts and load/blocks.ts, by the element'),
  text: file('load/runs.ts'),
  node: file('load/paragraph.ts and load/blocks.ts: the element or its shell'),
  textShell: file('load/runs.ts: the text element'),
  pieces: file('load/runs.ts, load/paragraph.ts and load/equations.ts: a whole field'),
  face: file("load/runs.ts: the run's `w:rPr`; client/room-runs.ts joins the pieces of a room run"),
  run: file('load/runs.ts: the run, or the room run of `oc:run` (client/room-runs.ts)'),
  inRun: file('load/runs.ts'),
  bare: file('load/paragraph.ts'),
  block: file('load/blocks.ts and load/sections.ts'),
  instr: file('load/runs.ts: a field instruction'),
  rebuilt: file('load/state.ts `withFileLabels`: each item the file labels'),
  prompt: file('load/paragraph.ts and load/blocks.ts: the control and its prompt'),
  token: file('load/state.ts: the container element'),
  pair: file('load/state.ts: the open item of the container'),
  home: file('load/paragraph.ts; client/room-runs.ts gives a joined piece the first one'),
  unitCount: file('load/paragraph.ts'),
  fieldSlot: file('load/paragraph.ts and load/moves.ts'),
  moveRange: file('load/moves.ts'),
  by: SUGGESTION,
  dels: SUGGESTION,
  fmts: file("load/runs.ts reads the run's `w:rPrChange`; taken records come from `takeovers`"),
  cancelledFmts: dropped(
    'they draw nothing and hold no card; only the undo of the press reads them'
  ),
  props: file('load/properties.ts and load/sections.ts read each change record of properties'),
  move: file('load/moves.ts: `w:moveFrom`, `w:moveTo` and the range name'),
  retracted: dropped('taken back by its author: compose writes none of it'),
  taken: UNDO_ONLY,
  gone: dropped('its container went: compose writes none of it'),
  quiet: file('load/notes.ts and load/text-boxes.ts; a new row gives it `within`'),
  within: file('load/state.ts: the row and cell revision marks around it'),
};

const WITHIN: Rules<OrderedWithin> = {
  by: file("load/state.ts: the row's `w:ins` and the cell's `w:cellIns`"),
  dels: file("load/state.ts: the row's `w:del` and the cell's `w:cellDel`"),
};

const PIECE: Rules<OrderedFieldPiece> = {
  node: file('load/runs.ts'),
  sourceRun: file('load/runs.ts'),
  recordFaces: file('load/runs.ts; client/takeovers.ts adds the faces under taken records'),
  recordFiles: file("load/runs.ts: the run's `w:rPrChange`"),
};

type PropertiesFormat = Extract<OrderedFormatRecord, { readonly kind: 'properties' }>;
type FaceFormat = Extract<OrderedFormatRecord, { readonly kind: 'face' }>;

const PROPERTIES_FORMAT: Rules<PropertiesFormat> = {
  kind: file("the run's `w:rPrChange` loads as a `face` record (fold-check.ts)"),
  id: derived('the load mints one for each change record; only pieces name it'),
  suggestion: SUGGESTION,
  seq: RECORD_SEQ,
  properties: file("load/runs.ts: the run's `w:rPr`"),
  remove: file("load/runs.ts: the run's `w:rPr` without them"),
  style: file("load/runs.ts: the run's `w:rStyle`"),
  replaced: TAKEN,
  over: open('111', "Word's one `w:rPrChange` holds it, so a take-back of the record takes it too"),
  untracked: file("written as its author's new words' own `w:rPr`; a reject takes the words"),
  change: UNDO_ONLY,
};

const FACE_FORMAT: Rules<FaceFormat> = {
  kind: file("load/runs.ts: the run's `w:rPrChange`"),
  id: derived('the load mints one for each change record; only pieces name it'),
  suggestion: SUGGESTION,
  seq: RECORD_SEQ,
  face: file("load/runs.ts: the `w:rPr` of the run's `w:rPrChange`"),
  fileRecord: file("load/runs.ts: the run's `w:rPrChange`"),
  replaced: TAKEN,
};

const ELEMENT_RECORD: Rules<OrderedElementRecord> = {
  kind: file('load/properties.ts: a change record of properties'),
  suggestion: SUGGESTION,
  seq: RECORD_SEQ,
  target: file('load/properties.ts: the properties element'),
  element: file("load/properties.ts: the container's properties"),
  fileRecord: file('load/properties.ts: the change record'),
  recordUnchanged: manifest('untracked', 'records.recordUnchanged', 'by the container label'),
  untracked: manifest('untracked', 'records.untracked', 'by the container label'),
  change: UNDO_ONLY,
  acceptedOnly: dropped('no write sets it'),
  tableCells: manifest('changeRecords', 'records.tableCells', 'the cells by their labels'),
  replaced: TAKEN,
  over: manifest('untracked', 'records.over', "Word's one change record holds the others"),
  merge: manifest('changeRecords', 'records.merge', 'by the container label'),
};

const PARAGRAPH_OP_RECORD: Rules<OrderedParagraphOpRecord> = {
  kind: file("the paragraph's `w:pPrChange` loads as an `element` record (fold-check.ts)"),
  suggestion: SUGGESTION,
  seq: RECORD_SEQ,
  target: file('load/properties.ts'),
  op: file("load/properties.ts: the paragraph's `w:pPr` and its `w:pPrChange`"),
  untracked: file("written as its author's new paragraph's own `w:pPr`; a reject takes it"),
  replaced: TAKEN,
  over: open('111', "Word's one `w:pPrChange` holds it, so a take-back of the record takes it too"),
};

const SECTION_RECORD: Rules<OrderedSectionRecord> = {
  kind: file("the section's `w:sectPrChange` loads as an `element` record (fold-check.ts)"),
  suggestion: SUGGESTION,
  seq: RECORD_SEQ,
  target: file('load/sections.ts'),
  op: file('load/sections.ts: the `w:sectPr` and its `w:sectPrChange`'),
  untracked: file("written as its author's new mark's own `w:sectPr`; a reject takes the mark"),
  keptOnReject: file("written as the section's own `w:type`, which every decision keeps"),
  replaced: TAKEN,
  over: open('111', "Word's one `w:sectPrChange` holds it, so a take-back takes it too"),
};

const OBJECT_RECORD: Rules<OrderedObjectRecord> = {
  kind: file('a direct edit: written in the drawing, which load/runs.ts reads as the node'),
  suggestion: dropped("a direct edit, accepted: only its author's undo reads it"),
  seq: dropped('a direct edit, accepted: written in the drawing'),
  target: file('load/runs.ts: the drawing'),
  op: file('load/runs.ts: the drawing as the op left it'),
  replaced: dropped('never set: a direct edit takes over no record'),
};

const SUGGESTION_FIELDS: Rules<OrderedSuggestion> = {
  id: manifest('suggestions', 'id', 'the load makes it under its id'),
  attribution: manifest('suggestions', 'author', 'with its date, attributes and person'),
  status: manifest('suggestions', 'status', 'pending when absent'),
  seq: derived("the load makes the listed suggestions first, in the manifest's order"),
  version: manifest('suggestions', 'version', 'the open sets it after its other steps'),
  direct: dropped("accepted when made; only its author's undo turns its status"),
};

const ATTRIBUTION: Rules<OrderedAttribution> = {
  author: manifest('suggestions', 'author', "a mark's `w:author` must be the same"),
  person: manifest('suggestions', 'person', "its Word name must be the mark's author"),
  date: manifest('suggestions', 'date', "a mark's `w:date` must be the same"),
  attributes: manifest('suggestions', 'attributes', "the mark's other attributes"),
};

const STORY: Rules<OrderedStory> = {
  name: file('load/load.ts `orderedStoryParts`'),
  part: file('load/load.ts: the story part'),
  tree: file('load/load.ts `loadStories`, and load/text-boxes.ts'),
  spaces: derived('`addOrderedStory` indexes the items by space'),
};

const DOCUMENT: Rules<OrderedDocument> = {
  file: derived("the open's file number (D8)"),
  stories: file('load/load.ts `loadStories`, and load/text-boxes.ts'),
  suggestions: manifest(
    'suggestions',
    'id',
    "listed first; the marks of no listed one are the file's"
  ),
  labelled: derived('indexed from the labels as the load makes the items'),
  ownEdited: UNDO_ONLY,
  added: UNDO_ONLY,
  counters: derived('see `OrderedCounters`'),
  runShells: file('load/runs.ts'),
  originals: file('load/load.ts `collectOriginals`'),
  revisionIds: derived('see `OrderedRevisionIds`'),
  placeholders: file('load/load.ts: the glossary building blocks'),
  movesWithRanges: file('load/moves.ts: the range markers compose wrote'),
  composeIds: dropped("every open starts it again; the ids it minted are the file's node ids"),
  paraIds: file("the stories hold the new paragraphs' `w14:paraId`, which a new id avoids"),
  commentIds: file('load/load.ts `collectOriginals`'),
  comments: file('load/comments.ts; see `OrderedComment`'),
  commentRanges: file('load/state.ts `noteCommentMarker`'),
  numbering: file('load/load.ts `recordFileLists`'),
  resources: derived('see `OrderedResources`'),
  partRevisions: file('load/part-revisions.ts; a listed `record` keeps its suggestion id'),
  namespaces: manifest('namespaces', 'person', 'each namespace an element names as a typer'),
  undo: dropped('set only while an edit runs'),
};

const COUNTERS: Rules<OrderedCounters> = {
  suggestion: derived('`newSuggestion` passes over every id the document holds'),
  insertion: dropped('the model does not read it'),
  seq: derived('the load counts the suggestions and records it makes'),
};

const REVISION_IDS: Rules<OrderedRevisionIds> = {
  byRole: file("load/state.ts: each mark's `w:id`, which compose wrote by role"),
  attributesByRole: file("load/state.ts: each mark's other attributes"),
  fileIds: file('load/load.ts `collectOriginals`'),
  reattributed: dropped('compose wrote the new attribution on every mark'),
  listed: manifest('suggestions', 'revisions', 'by listed suggestion, in order'),
};

const PART_REVISION: Rules<OrderedPartRevision> = {
  part: file('load/part-revisions.ts'),
  element: file('load/part-revisions.ts'),
  revision: file('load/part-revisions.ts'),
  styleId: file('load/part-revisions.ts'),
};

const COMMENT: Rules<OrderedComment> = {
  id: file('load/comments.ts'),
  author: file('load/comments.ts'),
  initials: file('load/comments.ts'),
  date: file('load/comments.ts'),
  paraIds: file('load/comments.ts'),
  paraId: file('load/comments.ts'),
  parent: file('load/comments.ts: the thread entries'),
  done: file('load/comments.ts: the thread entries'),
  element: file('load/comments.ts'),
  person: manifest('comments', 'person', "by the comment's `w:id`"),
};

const COMMENT_RANGE: Rules<OrderedCommentRange> = {
  story: file('load/state.ts `noteCommentMarker`'),
  start: file('load/state.ts `noteCommentMarker`'),
  end: file('load/state.ts `noteCommentMarker`'),
  reference: file('load/state.ts `noteCommentMarker`'),
};

const RESOURCES: Rules<OrderedResources> = {
  fileParts: file('load/load.ts `recordFileNames`'),
  fileRelationships: file('load/load.ts `recordFileNames`'),
  fileLists: file('load/load.ts `recordFileLists`'),
  fileStores: file('load/load.ts `recordFileStores`'),
  fileMedia: manifest('media', 'part', 'model/file-media.ts `recordFileMedia`'),
  added: derived('see each kind of resource'),
};

const FILE_MEDIA: Rules<OrderedFileMedia> = {
  part: manifest('media', 'part', 'by its part name'),
  contentType: manifest('media', 'contentType', 'checked against the content types'),
  sha256: manifest('media', 'sha256', 'the room storage serves the bytes by it'),
  size: manifest('media', 'size', 'checked when a client reads the bytes'),
  entry: manifest('media', 'entry', 'its place among the entries'),
};

const FILE_STORE: Rules<OrderedFileStore> = {
  storeItemId: file('load/load.ts `recordFileStores`'),
  root: file('load/load.ts `recordFileStores`'),
  boundNodes: file('load/load.ts `recordFileStores`'),
};

/** The rules of one kind of resource. */
type ResourceRules<K extends OrderedResource['kind']> = Rules<
  Extract<OrderedResource, { readonly kind: K }>
>;

// What every resource has: after a fold a resource is the file's, but for a list or a style
const KEY = dropped(
  'the fold writes the resource under the names its key gave; a load reads those'
);
const KIND = file('written as a part, a relationship or a definition of the file');

const MEDIA: ResourceRules<'media'> = {
  kind: KIND,
  key: KEY,
  suggestion: derived("the file's after a fold: it goes with the relationship that names it"),
  part: manifest('media', 'part', 'a reference of the file'),
  contentType: manifest('media', 'contentType', 'a reference of the file'),
  sha256: manifest('media', 'sha256', 'a reference of the file'),
  size: manifest('media', 'size', 'a reference of the file'),
};

const STORY_RESOURCE: ResourceRules<'story'> = {
  kind: KIND,
  key: KEY,
  suggestion: dropped(
    "a header's or a footer's is a direct edit's; a new notes part stays while the list holds it"
  ),
  part: file('load/load.ts: a story part of the file'),
  role: file("load/load.ts: the part's root and content type"),
  relationshipsOf: file("the part's own relationships are written"),
};

const PART_RESOURCE: ResourceRules<'part'> = {
  kind: KIND,
  key: KEY,
  suggestion: open('new', "the file's after a fold: a reject of the pasted chart keeps the part"),
  part: file('a part of the file'),
  type: file("the part's content type and relationship type"),
  root: file('the part as written'),
};

const RELATIONSHIP: ResourceRules<'relationship'> = {
  kind: KIND,
  key: KEY,
  suggestion: open('new', "the file's after a fold: only an image's goes with its last use"),
  owner: file('load/load.ts `recordFileNames`'),
  id: file('load/load.ts `recordFileNames`: the id minted from the key'),
  type: file('the relationship as written'),
  target: file('the relationship as written'),
};

const LIST: ResourceRules<'list'> = {
  kind: KIND,
  key: manifest('lists', 'key', 'client/lists.ts gives it back as a resource'),
  suggestion: manifest('lists', 'suggestion', 'client/lists.ts'),
  numId: derived('from the key (`orderedListNumber`)'),
  abstractNumId: derived('from the key (`orderedListNumber`)'),
  levels: file('client/lists.ts: the definition the numbering part holds'),
  definition: file('client/lists.ts: the definition the numbering part holds'),
  bullets: open('new', "a list with picture bullets stays the file's: a reject keeps it"),
};

const LIST_LEVEL: ResourceRules<'list-level'> = {
  kind: KIND,
  key: KEY,
  suggestion: open('new', "the file's after a fold: a reject of the indent keeps the level"),
  numId: file('the numbering part'),
  level: file('the numbering part'),
  format: file('the numbering part'),
};

const STYLE: ResourceRules<'style'> = {
  kind: KIND,
  key: manifest('styles', 'key', 'client/lists.ts gives it back as a whole style'),
  suggestion: manifest('styles', 'suggestion', 'client/lists.ts'),
  styleId: manifest('styles', 'styleId', 'client/lists.ts'),
  type: file('client/lists.ts: the whole `w:style`'),
  name: file('client/lists.ts: the whole `w:style`'),
  basedOn: file('client/lists.ts: the whole `w:style`'),
  next: file('client/lists.ts: the whole `w:style`'),
  link: file('client/lists.ts: the whole `w:style`'),
  uiPriority: file('client/lists.ts: the whole `w:style`'),
  unhideWhenUsed: file('client/lists.ts: the whole `w:style`'),
  paragraph: file('client/lists.ts: the whole `w:style`'),
  run: file('client/lists.ts: the whole `w:style`'),
  element: file('client/lists.ts: the whole `w:style`'),
};

const SETTING: ResourceRules<'setting'> = {
  kind: KIND,
  key: KEY,
  suggestion: dropped("a direct edit, accepted: only its author's undo reads it"),
  op: file('the settings part as the op left it'),
};

const CUSTOM_XML: ResourceRules<'custom-xml'> = {
  kind: KIND,
  key: KEY,
  suggestion: file("the store's node; compose takes it out once no control binds it"),
  namespaceUri: file('load/load.ts `recordFileStores`'),
  root: file('load/load.ts `recordFileStores`'),
  storeItemId: file('load/load.ts `recordFileStores`'),
  nodeId: file("the store's node, which the control binds"),
  label: file("the control's text"),
  data: file("the store's node"),
  control: file('load/blocks.ts and load/paragraph.ts: the control'),
};

/** Every table, by the type it covers. */
const TABLES: Readonly<Record<string, Readonly<Record<string, FoldRule>>>> = {
  OrderedItem: ITEM,
  OrderedWithin: WITHIN,
  OrderedFieldPiece: PIECE,
  'OrderedFormatRecord properties': PROPERTIES_FORMAT,
  'OrderedFormatRecord face': FACE_FORMAT,
  OrderedElementRecord: ELEMENT_RECORD,
  OrderedParagraphOpRecord: PARAGRAPH_OP_RECORD,
  OrderedSectionRecord: SECTION_RECORD,
  OrderedObjectRecord: OBJECT_RECORD,
  OrderedSuggestion: SUGGESTION_FIELDS,
  OrderedAttribution: ATTRIBUTION,
  OrderedStory: STORY,
  OrderedDocument: DOCUMENT,
  OrderedCounters: COUNTERS,
  OrderedRevisionIds: REVISION_IDS,
  OrderedPartRevision: PART_REVISION,
  OrderedComment: COMMENT,
  OrderedCommentRange: COMMENT_RANGE,
  OrderedResources: RESOURCES,
  OrderedFileMedia: FILE_MEDIA,
  OrderedFileStore: FILE_STORE,
  'OrderedResource media': MEDIA,
  'OrderedResource story': STORY_RESOURCE,
  'OrderedResource part': PART_RESOURCE,
  'OrderedResource relationship': RELATIONSHIP,
  'OrderedResource list': LIST,
  'OrderedResource list-level': LIST_LEVEL,
  'OrderedResource style': STYLE,
  'OrderedResource setting': SETTING,
  'OrderedResource custom-xml': CUSTOM_XML,
};

const RULES: readonly FoldRule[] = Object.values(TABLES).flatMap((table) => Object.values(table));

/** Whether a JSON value holds a field at `path`: in a list, in any of its entries. */
function holdsPath(value: unknown, path: readonly string[]): boolean {
  if (Array.isArray(value)) {
    return value.some((entry) => holdsPath(entry, path));
  }

  const [key, ...rest] = path;
  if (key === undefined) {
    return value !== undefined;
  }

  if (typeof value !== 'object' || value === null || !Object.hasOwn(value, key)) {
    return false;
  }

  return holdsPath((value as Record<string, unknown>)[key], rest);
}

/**
 * Ana's row height, the organiser's column, a direct row height and a direct column beside the new
 * cell: the direct column is made over Ana's grid record (section `untracked`, `over`).
 */
const DIRECT_OVER: CheckpointRoom = {
  name: 'direct column over a row height',
  source: APPLY_FILE,
  edit: (doc) => {
    const height = (row: number) =>
      ({
        kind: 'table-property',
        on: [addressOf(opens(doc, 'tr')[row]!)],
        op: { op: 'setTableRowHeight', heightTwips: 500 },
      }) as const;
    const right = (k: number) =>
      ({ cell: addressOf(opens(doc, 'tc')[k]!), where: 'right' }) as const;
    apply(doc, 'ana:1', { ...height(1), ...EDIT }, 'Ana');
    apply(doc, 'olive:1', { kind: 'insert-column', ...EDIT, ...right(3) }, 'Olive');
    apply(doc, 'olive:2', { ...height(0), ...DIRECT }, 'Olive');
    apply(doc, 'olive:3', { kind: 'insert-column', ...DIRECT, ...right(4) }, 'Olive');
  },
};

/** The rooms whose checkpoints hold a fact of each section. */
const ROOMS: readonly CheckpointRoom[] = [
  ROOM_FACTS,
  STYLE_RECORDS,
  TYPED_RUNS,
  TYPED_BARE,
  SPLIT_RUN,
  PASTED_LIST,
  TABLE_RECORDS,
  DIRECT_OVER,
];

/** The checkpoint of each room: its manifest as JSON, and as the reader reads it. */
async function roomCheckpoints(): Promise<
  { readonly json: Record<string, unknown>; readonly read: CheckpointOpened }[]
> {
  const out = [];
  for (const room of ROOMS) {
    const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
    room.edit(doc);
    const bytes = await checkpointBytes(doc, base, roomBlobs(base));
    const entry = unzipSync(bytes)[CHECKPOINT_MANIFEST.slice(1)];
    const pkg = readOoxmlPackage(bytes, {});
    if (entry === undefined || !pkg.ok) {
      throw new Error(`${room.name}: no checkpoint`);
    }

    const read = readCheckpoint(pkg.package);
    if (!read.ok) {
      throw new Error(`${room.name}: ${read.reason}`);
    }

    out.push({ json: JSON.parse(strFromU8(entry)) as Record<string, unknown>, read });
  }

  return out;
}

const CHECKPOINTS = roomCheckpoints();

describe('every fact of the model has a fold rule', () => {
  test('each manifest rule names a field the writer writes and the reader reads', async () => {
    const checkpoints = await CHECKPOINTS;
    const missing: string[] = [];
    for (const rule of RULES) {
      if (rule.rule !== 'manifest') {
        continue;
      }

      const path = rule.field.split('.');
      const written = checkpoints.some(({ json }) => holdsPath(json[rule.section], path));
      const read = checkpoints.some(({ read }) => SECTIONS[rule.section](read) > 0);
      if (!written || !read) {
        missing.push(`${rule.section}.${rule.field}`);
      }
    }

    expect(missing).toEqual([]);
  });

  test('each section the writer writes has a rule', async () => {
    const [checkpoint] = await CHECKPOINTS;
    const written = Object.keys(checkpoint!.json).filter((key) => !HEADER.has(key));
    const named = RULES.flatMap((rule) => (rule.rule === 'manifest' ? [rule.section] : []));
    const sections = new Set(Object.keys(SECTIONS));
    expect(new Set(written)).toEqual(sections);
    expect(new Set<string>(named)).toEqual(sections);
  });
});
