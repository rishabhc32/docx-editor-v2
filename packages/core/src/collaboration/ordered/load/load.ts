// Story parts as an ordered document. Every value comes from a file, so the walk is bounded:
// by depth (the XML reader's own ceiling) and by item count (the engine's element budget).

import { buildingBlocksOf } from '../../../store/package/building-blocks.ts';
import { boundCustomXmlNodeIdsInPackage } from '../../../store/package/custom-node-payloads.ts';
import { customXmlDataParts } from '../../../store/package/custom-xml-part.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { DEFAULT_LIMITS } from '../../../store/runtime/limits.ts';
import { addOrderedStory, createOrderedDocument } from '../model/document.ts';
import type { OrderedPerson } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { PROMPT_RECORD_ROLES, reserveFileRevisionId } from '../model/revision-ids.ts';
import {
  BLOCK_TOKENS,
  type OrderedAttribution,
  type OrderedDocument,
  type OrderedItem,
  type OrderedSuggestionStatus,
} from '../model/types.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';
import { loadBlocks } from './blocks.ts';
import { loadComments } from './comments.ts';
import { recordFurniture, withLoadedSections } from './sections.ts';
import { linkNewNotes, linkStruckNotes } from './notes.ts';
import { loadPartRevisions } from './part-revisions.ts';
import { loadTextBoxStories } from './text-boxes.ts';
import {
  itemsWithPrompts,
  type LoadBudget,
  MARK_ROLES,
  newLoadBudget,
  newLoadState,
  NOTE_MARK_ROLES,
  RevisionReader,
  withFileLabels,
} from './state.ts';

/** Bounds on a load. A caller may lower them, never raise them. @public */
export interface OrderedLoadOptions {
  /**
   * The room's Word file number. Every space the load makes holds it, so an address from
   * another Word file of the room never resolves in this document (design D8). The room and
   * every client load one file with the same number, so their addresses agree. Default 0.
   */
  readonly file?: number;
  /** Most items in the whole document. Default and ceiling: the engine's element budget. */
  readonly maxItems?: number;
  /** Deepest nesting of blocks and inline containers. Default and ceiling: the XML depth cap. */
  readonly maxDepth?: number;
  /**
   * The glossary's placeholder blocks by name: the prompts controls name through
   * `w:placeholder`. {@link loadOrderedPackage} reads them from the package.
   */
  readonly placeholders?: ReadonlyMap<string, readonly OoxmlNode[]>;
  /**
   * A room checkpoint's pending suggestions, and the decided ones its other records name, each
   * with the `w:id`s the file writes for its marks (`openOrderedCheckpoint` reads them). The load
   * makes them first, in this order, each under its own id, and reads each mark with one of these
   * ids, the author and the date as that one suggestion. Default: none, so every suggestion is the
   * file's own, one for each revision triple, with no person, under a new id.
   */
  readonly suggestions?: readonly {
    // Spelled out: the manifest's own type is not public
    readonly id: string;
    readonly attribution: OrderedAttribution;
    readonly revisions: readonly string[];
    /**
     * The local name of the styles part's change record the suggestion is, whose `w:id` is the
     * one in `revisions`: the record group it names keeps the suggestion's id.
     */
    readonly record?: string;
    /** Pending when absent. */
    readonly status?: OrderedSuggestionStatus;
  }[];
  /**
   * The persons of a room checkpoint's comments by comment `w:id` (`openOrderedCheckpoint` reads
   * them): a comment with one of these ids is that person's. Default: none, so every comment is
   * the file's own, with no person.
   */
  readonly commentPersons?: ReadonlyMap<string, OrderedPerson>;
}

const STORY_ROOTS = new Set(['document', 'hdr', 'ftr', 'footnotes', 'endnotes']);

/** Whether a part is a story: the main document, a header, a footer, or a notes part. @public */
export function isOrderedStoryPart(part: OoxmlPart): boolean {
  return STORY_ROOTS.has(part.root.localName) && isWml(part.root, part.root.localName);
}

/** The story parts of a package, main document first, then the others by name. @public */
export function orderedStoryParts(pkg: OoxmlPackage): Map<string, OoxmlPart> {
  const names = [...pkg.parts.keys()].filter((name) => {
    const part = pkg.parts.get(name);
    return part !== undefined && isOrderedStoryPart(part);
  });
  names.sort(byStoryOrder(pkg.mainDocumentPart));
  const out = new Map<string, OoxmlPart>();
  for (const name of names) {
    const part = pkg.parts.get(name);
    if (part) out.set(name, part);
  }

  return out;
}

/** Compares story part names: the main document first, then the others by name. */
export function byStoryOrder(main: string): (a: string, b: string) => number {
  return (a, b) => {
    if (a === main) {
      return -1;
    }

    if (b === main) {
      return 1;
    }

    if (a < b) {
      return -1;
    }

    return a > b ? 1 : 0;
  };
}

/** The blocks of a story part: a document's body, a header's or footer's root, the notes. */
export function storyBlocksOf(part: OoxmlPart): readonly OoxmlNode[] {
  if (part.root.localName !== 'document') return part.root.children;
  const body = part.root.children.find((child) => isWml(child, 'body'));
  if (!body || body.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', `${part.name} has no body`);
  return body.children;
}

const LOADED = new WeakMap<OrderedDocument, OoxmlPackage>();

/**
 * The package a document was loaded from, which the settings writes go over
 * (`compose/settings.ts`); null for a document loaded from none.
 */
export function loadedPackageOf(doc: OrderedDocument): OoxmlPackage | null {
  return LOADED.get(doc) ?? null;
}

/**
 * Load a package's stories as one ordered document, with the prompts its glossary holds.
 *
 * @public
 */
export function loadOrderedPackage(
  pkg: OoxmlPackage,
  options: Omit<OrderedLoadOptions, 'placeholders'> = {}
): OrderedDocument {
  const placeholders = new Map<string, readonly OoxmlNode[]>();
  for (const block of buildingBlocksOf(pkg))
    if (!placeholders.has(block.name)) placeholders.set(block.name, block.blocks);
  const stories = orderedStoryParts(pkg);
  const budget = newLoadBudget(options.maxItems, options.maxDepth);
  const doc = loadStories(stories, { ...options, placeholders }, budget);
  recordFileNames(doc, pkg);
  recordFileLists(doc, pkg);
  recordFileStores(doc, pkg);
  recordFurniture(doc, pkg);
  LOADED.set(doc, pkg);

  // After the stories, so their suggestion ids stay as they were
  loadPartRevisions(doc, pkg, options.suggestions);
  for (const part of pkg.parts.values()) {
    // The comments part is no story: only its comment and revision ids
    if (isWml(part.root, 'comments')) {
      collectOriginals(doc, part.root, { indexReusable: false });
    }
  }

  loadComments(doc, pkg, stories.values(), budget, options.commentPersons);
  return doc;
}

/**
 * The names the loaded file holds, which a resource a change adds must not take (D14): part
 * names (OPC names ignore case) and relationship ids by owner.
 */
function recordFileNames(doc: OrderedDocument, pkg: OoxmlPackage): void {
  const { fileParts, fileRelationships } = doc.resources;
  for (const name of [...pkg.parts.keys(), ...pkg.partBytes.keys()])
    fileParts.add(name.toLowerCase());
  const owned = (owner: string): Set<string> => {
    let ids = fileRelationships.get(owner);
    if (!ids) {
      ids = new Set();
      fileRelationships.set(owner, ids);
    }

    return ids;
  };
  for (const [owner, records] of pkg.relationships)
    for (const record of records) owned(owner).add(record.id);
  for (const target of pkg.externalTargets) owned(target.ownerPart).add(target.id);
}

/** Each numbering element whose id a change must not take: its role, and its id attribute. */
const NUMBERING_IDS: ReadonlyMap<string, { readonly role: string; readonly id: string }> = new Map([
  ['num', { role: 'num', id: 'numId' }],
  ['abstractNum', { role: 'abstract', id: 'abstractNumId' }],
  ['numPicBullet', { role: 'bullet', id: 'numPicBulletId' }],
]);

/**
 * The list numbers, list definition ids and picture bullet ids the file's numbering parts hold,
 * which a change must not take (D14), and the list numbers an edit may name (`doc.numbering`).
 */
function recordFileLists(doc: OrderedDocument, pkg: OoxmlPackage): void {
  const numbering = new Set<string>();
  for (const part of pkg.parts.values()) {
    if (!isWml(part.root, 'numbering')) {
      continue;
    }

    for (const child of part.root.children) {
      if (child.kind === 'textValue' || !isWml(child, child.localName)) {
        continue;
      }

      const entry = NUMBERING_IDS.get(child.localName);
      const id = entry && attributeOf(child, entry.id);
      if (!entry || id === undefined) {
        continue;
      }

      doc.resources.fileLists.add(`${entry.role} ${id}`);
      if (entry.role === 'num') {
        numbering.add(id);
      }
    }
  }

  doc.numbering = numbering;
}

/** The custom XML data stores the main document part relates, which a custom node's binds. */
function recordFileStores(doc: OrderedDocument, pkg: OoxmlPackage): void {
  for (const store of customXmlDataParts(pkg, pkg.mainDocumentPart)) {
    const root = pkg.parts.get(store.partName)?.root.localName;
    if (root === undefined || doc.resources.fileStores.has(store.namespaceUri)) continue;
    doc.resources.fileStores.set(store.namespaceUri, {
      storeItemId: store.itemId.toUpperCase(),
      root,
      boundNodes: boundCustomXmlNodeIdsInPackage(pkg, store.itemId),
    });
  }
}

/**
 * Load story parts, in the order given, as one ordered document. Revisions in the parts become
 * pending suggestions, one for each revision triple (`w:id`, `w:author`, `w:date`). Refuses a
 * part that is not a story, nests past the depth limit, or holds more items than the limit.
 *
 * @public
 */
export function loadOrderedDocument(
  parts: ReadonlyMap<string, OoxmlPart>,
  options: OrderedLoadOptions = {}
): OrderedDocument {
  return loadStories(parts, options, newLoadBudget(options.maxItems, options.maxDepth));
}

/** {@link loadOrderedDocument} within `budget`, which the rest of a package's load shares. */
function loadStories(
  parts: ReadonlyMap<string, OoxmlPart>,
  options: OrderedLoadOptions,
  budget: LoadBudget
): OrderedDocument {
  const doc = createOrderedDocument(options.file ?? 0);
  for (const [name, blocks] of options.placeholders ?? []) {
    doc.placeholders.set(name, blocks);
  }

  // No fixed suggestion: each file revision is its own, or the suggestion a checkpoint lists
  const revisions = new RevisionReader(doc, undefined, options.suggestions);

  // Every part's revision ids first, so an id the load mints (a move's range markers that the
  // file did not write) never repeats one a later part holds
  for (const [name, part] of parts) {
    if (!isOrderedStoryPart(part)) {
      throw new OrderedRefusal('malformed-story', `${name} is not a story`);
    }

    collectOriginals(doc, part.root);
  }

  for (const [name, part] of parts) {
    const notes = isWml(part.root, 'footnotes') || isWml(part.root, 'endnotes');
    const state = newLoadState(doc, name, revisions, budget, notes ? NOTE_MARK_ROLES : MARK_ROLES);
    loadBlocks(state, storyBlocksOf(part), 1);
    const items = withFileLabels(withLoadedSections(state, part, itemsWithPrompts(state)));
    checkEndsWithMark(name, items);

    // The main document's text boxes are stories of their own (group 9.9)
    loadTextBoxStories(doc, addOrderedStory(doc, name, part, items), revisions, budget);
  }

  // A prompt the glossary kept may hold a change record (a format on the prompt before content
  // filled its control): each is a suggestion, read as the stories' marks are, so a decision
  // reaches it when the prompt shows again
  for (const blocks of doc.placeholders.values()) {
    for (const block of blocks) {
      readPromptRecords(revisions, block);
    }
  }

  linkNewNotes(doc);
  linkStruckNotes(doc);
  return doc;
}

/** The suggestion of each change record in a prompt the glossary kept. */
function readPromptRecords(revisions: RevisionReader, node: OoxmlNode): void {
  if (node.kind === 'textValue') {
    return;
  }

  const role = PROMPT_RECORD_ROLES.get(node.localName);
  if (role !== undefined && isWml(node, node.localName)) {
    revisions.suggestionFor(node, role);
  }

  for (const child of node.children) {
    readPromptRecords(revisions, child);
  }
}

/** Every inline item must belong to a paragraph: nothing but blocks after the last mark. */
function checkEndsWithMark(name: string, items: readonly OrderedItem[]): void {
  for (let k = items.length - 1; k >= 0; k -= 1) {
    const item = items[k];
    if (!item || item.kind === 'mark') return;
    const blockToken =
      (item.kind === 'open' || item.kind === 'close') &&
      item.token !== undefined &&
      BLOCK_TOKENS.has(item.token);
    if (!blockToken && !(item.kind === 'zero' && item.block))
      throw new OrderedRefusal(
        'malformed-story',
        `${name} has inline content after its last paragraph`
      );
  }
}

/** Nodes compose may reuse when untouched: paragraphs, tables, rows, cells, controls, notes. */
const REUSABLE_KINDS = new Set<OoxmlNode['kind']>([
  'paragraph',
  'table',
  'tableRow',
  'tableCell',
  'contentControl',
  'note',
]);
const REUSABLE_NAMES = new Set(['sdt', 'footnote', 'endnote', 'customXml']);

/** Revision elements whose `w:id` a minted id must not repeat. */
const REVISION_NAMES = new Set([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'moveFromRangeStart',
  'moveToRangeStart',
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblPrExChange',
  'tblGridChange',
  'trPrChange',
  'tcPrChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'numberingChange',
  'customXmlInsRangeStart',
  'customXmlDelRangeStart',
]);

/** Elements that carry a comment's `w:id`. */
const COMMENT_NAMES = new Set([
  'commentRangeStart',
  'commentRangeEnd',
  'commentReference',
  'comment',
]);

/**
 * Index the reusable nodes, and note every revision `w:id` and comment id: one bounded walk of
 * the part.
 */
function collectOriginals(
  doc: OrderedDocument,
  root: OoxmlElement,
  { indexReusable = true }: { readonly indexReusable?: boolean } = {}
): void {
  const stack: OoxmlNode[] = [root];
  let visited = 0;
  const ceiling = DEFAULT_LIMITS.maxElementCount;
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    visited += 1;
    if (visited > ceiling) throw new OrderedRefusal('too-many-items', 'part too large to index');
    if (indexReusable && (REUSABLE_KINDS.has(node.kind) || REUSABLE_NAMES.has(node.localName))) {
      doc.originals.set(node.id, node);
    }

    if (REVISION_NAMES.has(node.localName)) {
      const id = attributeOf(node, 'id');
      if (id !== undefined) reserveFileRevisionId(doc, id);
    } else if (COMMENT_NAMES.has(node.localName)) {
      const id = attributeOf(node, 'id');
      if (id !== undefined) doc.commentIds.add(id);
    }

    for (let k = node.children.length - 1; k >= 0; k -= 1) {
      const child = node.children[k];
      if (child !== undefined && child.kind !== 'textValue') stack.push(child);
    }
  }
}
