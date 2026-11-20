// The document and its stories: creation, the label index, lookups by address, and what shows.

import { SumTree, type SummaryOps } from '../tree/sum-tree.ts';
import { sameMaker, type OrderedMaker } from './person.ts';
import { indexPromptItem } from './prompt-marks.ts';
import { OrderedRefusal } from './refusal.ts';
import { createOrderedResources } from './resources.ts';
import {
  loadedSpace,
  type OrderedDocument,
  type OrderedItem,
  type OrderedItemAddress,
  type OrderedItemSummary,
  type OrderedStory,
  type OrderedSuggestion,
  type OrderedSuggestionStatus,
  type OrderedWithin,
} from './types.ts';
import type { OoxmlPart } from '../../../store/package/ooxml-tree.ts';

/**
 * A new, empty document, for the room's Word file `file` (0 when the room has no numbers).
 *
 * @throws RangeError when `file` is not a safe integer from 0.
 * @public
 */
export function createOrderedDocument(file = 0): OrderedDocument {
  if (!Number.isSafeInteger(file) || file < 0)
    throw new RangeError('A Word file number is a safe integer from 0.');
  return {
    file,
    stories: new Map(),
    suggestions: new Map(),
    labelled: new Map(),
    ownEdited: new Map(),
    added: new Map(),
    counters: { suggestion: 0, insertion: 0, seq: 0 },
    runShells: new Map(),
    originals: new Map(),
    revisionIds: {
      byRole: new Map(),
      attributesByRole: new Map(),
      fileIds: new Set(),
      listed: new Map(),
    },
    placeholders: new Map(),
    movesWithRanges: new Set(),
    composeIds: { next: 1 },
    paraIds: new Set(),
    commentIds: new Set(),
    comments: new Map(),
    commentRanges: new Map(),
    numbering: null,
    resources: createOrderedResources(),
    partRevisions: new Map(),
    namespaces: new Map(),
    undo: null,
  };
}

/** Shared empty label list: labels are replaced, never pushed to. */
const NONE: readonly never[] = Object.freeze([]);

/** Shared empty container labels. */
export const NOT_WITHIN: OrderedWithin = Object.freeze({ by: NONE, dels: NONE });

/** The parts of an item a caller must give; every other field takes its default. */
export type OrderedItemInit = Partial<Omit<OrderedItem, 'leaf'>> &
  Pick<OrderedItem, 'space' | 'kind'>;

/** A new item, with no labels unless given. It is in no story yet. @public */
export function createOrderedItem(init: OrderedItemInit): OrderedItem {
  return {
    offset: 0,
    length: 1,
    text: '',
    node: null,
    face: null,
    run: '',
    by: null,
    dels: NONE,
    fmts: NONE,
    cancelledFmts: NONE,
    props: NONE,
    move: null,
    retracted: false,
    gone: false,
    quiet: false,
    within: NOT_WITHIN,
    ...init,
    leaf: undefined,
  };
}

// ---- visibility ----

/** A suggestion's status; null for the file's own content (no suggestion). @public */
export function statusOf(
  doc: OrderedDocument,
  suggestion: string | null
): OrderedSuggestionStatus | null {
  return suggestion === null ? null : (doc.suggestions.get(suggestion)?.status ?? null);
}

/** Whether `id` names a pending suggestion. */
export const isPending = (doc: OrderedDocument, id: string | null): id is string =>
  id !== null && statusOf(doc, id) === 'pending';

/** The innermost pending new row or cell around an item, or null. */
export function innermostNewContainer(doc: OrderedDocument, item: OrderedItem): string | null {
  for (let k = item.within.by.length - 1; k >= 0; k -= 1) {
    const id = item.within.by[k] ?? null;
    if (isPending(doc, id)) {
      return id;
    }
  }

  return null;
}

/**
 * The pending suggestion that inserted an item, if any: its own, else the innermost new row or
 * cell around it. Word marks a new row and its cells' paragraph marks, not the cells, so a cell
 * loads with no inserter of its own; the host reads the row as covering its paragraphs whole
 * (`isInside`), so a change on the cell or in its text sits inside the row.
 */
export function inserter(doc: OrderedDocument, item: OrderedItem): string | null {
  if (isPending(doc, item.by)) {
    return item.by;
  }

  return innermostNewContainer(doc, item);
}

/** Whether `id` names a pending suggestion that someone other than `maker` made. */
export const pendingByOther = (
  doc: OrderedDocument,
  id: string | null,
  maker: OrderedMaker
): id is string => isPending(doc, id) && !madeBy(doc, id, maker);

/** Whether `maker` made the suggestion `id`: the same person ID (`sameMaker`). */
export function madeBy(doc: OrderedDocument, id: string, maker: OrderedMaker): boolean {
  const found = doc.suggestions.get(id);
  return found !== undefined && sameMaker(found.attribution, maker);
}

/** The suggestion with this id, or null. */
export function suggestionOrNull(
  doc: OrderedDocument,
  id: string | null | undefined
): OrderedSuggestion | null {
  return id === null || id === undefined ? null : (doc.suggestions.get(id) ?? null);
}

/** Whether one maker made the suggestions `a` and `b` (`sameMaker`). */
export function oneMaker(doc: OrderedDocument, a: string, b: string): boolean {
  const found = doc.suggestions.get(b);
  return found !== undefined && madeBy(doc, a, found.attribution);
}

/** The suggestion with this id when it is pending and `maker` made it, else null. */
export function pendingBy(
  doc: OrderedDocument,
  id: string,
  maker: OrderedMaker
): OrderedSuggestion | null {
  const found = doc.suggestions.get(id);
  return found && found.status === 'pending' && sameMaker(found.attribution, maker) ? found : null;
}

/**
 * Hidden for good: the item shows in no state of the document that is still possible. Its
 * container went, its author took it back, the suggestion that inserted it was rejected, or a
 * suggestion that struck it was accepted. Pending struck text shows struck, and pending
 * inserted text shows inserted.
 */
export function isHidden(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.gone || item.retracted) return true;
  for (const inserter of item.within.by) if (statusOf(doc, inserter) === 'rejected') return true;
  for (const deleter of item.within.dels) if (statusOf(doc, deleter) === 'accepted') return true;

  // A split's edge stays when its Enter is rejected, and goes with its container or its own removal
  if (isSplitEdge(item)) {
    if (insertionTakenBack(doc, item) || splitContainerHidden(doc, item)) {
      return true;
    }
  } else if (item.by !== null && statusOf(doc, item.by) === 'rejected') {
    return true;
  }

  for (const deleter of item.dels) if (statusOf(doc, deleter) === 'accepted') return true;
  return false;
}

/**
 * A close or open item Enter put inside a link or a control to split it (its space is not the
 * container's). A decision on the Enter keeps the two containers, as the engine and a reopened
 * Word file do; only taking the Enter back (its author's, or the organiser's undo of a direct
 * Enter) leaves one container again. When the container itself goes, both halves go.
 */
export function isSplitEdge(item: OrderedItem): boolean {
  return (item.kind === 'open' || item.kind === 'close') && item.space !== item.pair;
}

/** Whether the unit's insertion was taken back: by its author, or a direct edit's revert. */
export function insertionTakenBack(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.retracted) return true;
  const found = item.by === null ? undefined : doc.suggestions.get(item.by);
  return found?.direct === true && found.status === 'rejected';
}

/**
 * A split's edge belongs to the container it split: when that container goes whole (its
 * insertion rejected or taken back, both its ends deleted), both halves go with it, whatever the
 * Enter. A removal of one half strikes that half's own edges, so the other half keeps its own.
 */
function splitContainerHidden(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.pair === undefined) return false;
  for (const story of doc.stories.values()) {
    const ends = story.spaces.get(item.pair);
    if (ends?.length) return ends.every((end) => isHidden(doc, end));
  }

  return false;
}

/** The suggestions that struck the item and are still pending, in order. */
export function pendingDeleters(doc: OrderedDocument, item: OrderedItem): string[] {
  return item.dels.filter((deleter) => isPending(doc, deleter));
}

/**
 * The suggestion that moved an item away: of its deleters, the one that inserted the move's
 * destination, or the first one when the file holds the source without its destination. Null
 * when no move holds the item.
 */
export function moverOf(doc: OrderedDocument, item: OrderedItem): string | null {
  if (item.move === null || item.by !== null) return null;
  for (const deleter of item.dels)
    for (const labelled of doc.labelled.get(deleter) ?? [])
      if (labelled.by === deleter && labelled.move === item.move) return deleter;
  return item.dels[0] ?? null;
}

/**
 * Whether `deleter` strikes the item as its move's source. A move's labels stay on its words
 * after its decision, so a later deletion of a rejected move's words is a plain deletion.
 */
export function movedAwayBy(doc: OrderedDocument, item: OrderedItem, deleter: string): boolean {
  return moverOf(doc, item) === deleter;
}

/** Whether a pending suggestion inserted the item. */
export function insertedPending(doc: OrderedDocument, item: OrderedItem): boolean {
  return isPending(doc, item.by);
}

/** The item's width in the engine's paragraph offsets: characters and atoms count. */
export function widthOf(item: OrderedItem): number {
  switch (item.kind) {
    case 'text':
      return item.length;
    case 'atom':
      return 1;
    default:
      return 0;
  }
}

/**
 * Whether any label names a suggestion (or the item went, or the file labelled it), so its
 * paragraph is drawn again.
 */
export function isLabelled(item: OrderedItem): boolean {
  return (
    item.by !== null ||
    item.dels.length > 0 ||
    item.fmts.length > 0 ||
    item.props.length > 0 ||
    item.retracted ||
    item.gone ||
    item.within.by.length > 0 ||
    item.within.dels.length > 0 ||
    item.rebuilt === true
  );
}

function summaryOps(doc: OrderedDocument): SummaryOps<OrderedItem, OrderedItemSummary> {
  return {
    zero: () => ({ units: 0, shown: 0, marks: 0 }),
    add: (into, other) => {
      into.units += other.units;
      into.shown += other.shown;
      into.marks += other.marks;
    },
    of: (item) => {
      const visible = !isHidden(doc, item);
      return {
        units: item.length,
        shown: visible ? widthOf(item) : 0,
        marks: visible && item.kind === 'mark' ? 1 : 0,
      };
    },
  };
}

// ---- stories ----

/**
 * Add a story of items, in order. The items must be in no story yet.
 *
 * @public
 */
export function addOrderedStory(
  doc: OrderedDocument,
  name: string,
  part: OoxmlPart,
  items: readonly OrderedItem[]
): OrderedStory {
  if (doc.stories.has(name)) throw new RangeError(`The story ${name} is already loaded.`);
  for (const item of items) if (item.leaf !== undefined) throw new OrderedRefusal('item-in-tree');
  const tree = new SumTree<OrderedItem, OrderedItemSummary>(summaryOps(doc), items);
  const story: OrderedStory = { name, part, tree, spaces: new Map() };
  for (const item of items) indexItem(doc, story, item);
  doc.stories.set(name, story);
  return story;
}

/** The story by name, or a refusal. */
export function storyNamed(doc: OrderedDocument, name: string): OrderedStory {
  const story = doc.stories.get(name);
  if (!story) throw new OrderedRefusal('unknown-story', name);
  return story;
}

/** The story that holds an address's space, or null. */
export function storyOfAddress(
  doc: OrderedDocument,
  address: OrderedItemAddress
): OrderedStory | null {
  for (const story of doc.stories.values()) if (story.spaces.has(address.space)) return story;
  return null;
}

/** The main document's story, or null when the list has none. */
export function mainStory(doc: OrderedDocument): OrderedStory | null {
  for (const story of doc.stories.values()) {
    if (story.part.root.localName === 'document') {
      return story;
    }
  }

  return null;
}

/** The story whose tree holds the item, or null. */
export function storyOfItem(doc: OrderedDocument, item: OrderedItem): OrderedStory | null {
  const tree = item.leaf?.tree;
  if (tree === undefined) return null;
  for (const story of doc.stories.values()) if (story.tree === tree) return story;
  return null;
}

/**
 * Why an address names no unit of the document: `other-file` when a load of another Word file
 * of the room made its space (a change made against an older file), else `unknown-address`.
 */
export function missingAddress(
  doc: OrderedDocument,
  address: OrderedItemAddress
): 'other-file' | 'unknown-address' {
  const loaded = loadedSpace(address.space);
  return loaded !== null && loaded.file !== doc.file ? 'other-file' : 'unknown-address';
}

/** The item that holds a unit, or null. A binary search in the unit's space. @public */
export function itemAt(story: OrderedStory, address: OrderedItemAddress): OrderedItem | null {
  const list = story.spaces.get(address.space);
  if (!list) return null;
  let lo = 0;
  let hi = list.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const item = list[mid];
    if (item === undefined) return null;
    if (address.offset < item.offset) hi = mid - 1;
    else if (address.offset >= item.offset + item.length) lo = mid + 1;
    else return item;
  }

  return null;
}

// ---- indexes ----

/** Put an item in its space's list, by offset, in the label index and in the prompt index. */
export function indexItem(doc: OrderedDocument, story: OrderedStory, item: OrderedItem): void {
  const list = story.spaces.get(item.space);
  if (!list) story.spaces.set(item.space, [item]);
  else {
    let at = list.length;
    while (at > 0 && (list[at - 1]?.offset ?? 0) > item.offset) at -= 1;
    list.splice(at, 0, item);
  }

  labelIndex(doc, item);
  indexPromptItem(doc, item);
}

/** Take an item out of its space's list and the label index. */
export function unindexItem(doc: OrderedDocument, story: OrderedStory, item: OrderedItem): void {
  const list = story.spaces.get(item.space);
  if (list) {
    const at = list.indexOf(item);
    if (at >= 0) list.splice(at, 1);
    if (list.length === 0) story.spaces.delete(item.space);
  }

  labelUnindex(doc, item);
}

/** The suggestions an item's labels name. */
export function suggestionsOf(item: OrderedItem): Set<string> {
  const ids = new Set<string>();
  if (item.by !== null) ids.add(item.by);
  for (const deleter of item.dels) ids.add(deleter);
  for (const record of item.fmts) ids.add(record.suggestion);
  for (const record of item.cancelledFmts) ids.add(record.suggestion);
  for (const record of item.props) ids.add(record.suggestion);
  for (const id of item.within.by) ids.add(id);
  for (const id of item.within.dels) ids.add(id);
  return ids;
}

/** Put an item in an index's entry. */
function addTo(index: Map<string, Set<OrderedItem>>, key: string, item: OrderedItem): void {
  let set = index.get(key);
  if (!set) {
    set = new Set();
    index.set(key, set);
  }

  set.add(item);
}

/** Take an item out of an index's entry; an emptied entry goes. */
function removeFrom(index: Map<string, Set<OrderedItem>>, key: string, item: OrderedItem): void {
  const set = index.get(key);
  if (!set) return;
  set.delete(item);
  if (set.size === 0) index.delete(key);
}

/** The changes of their author's own edits that an item's labels name (`ownEdited`). */
function ownEditsOf(item: OrderedItem): Set<string> {
  const changes = new Set<string>();
  if (item.taken) changes.add(item.taken);
  for (const record of item.fmts)
    if (record.kind === 'properties' && record.change !== undefined) changes.add(record.change);
  return changes;
}

/** Record which suggestions label an item, and which own edits changed it. */
export function labelIndex(doc: OrderedDocument, item: OrderedItem): void {
  for (const id of suggestionsOf(item)) addTo(doc.labelled, id, item);
  for (const change of ownEditsOf(item)) addTo(doc.ownEdited, change, item);
}

/** Forget which suggestions label an item, and which own edits changed it. */
export function labelUnindex(doc: OrderedDocument, item: OrderedItem): void {
  for (const id of suggestionsOf(item)) removeFrom(doc.labelled, id, item);
  for (const change of ownEditsOf(item)) removeFrom(doc.ownEdited, change, item);
}

/** The items that hold units a change added to its author's earlier pending suggestions. */
export function itemsAddedBy(doc: OrderedDocument, change: string): OrderedItem[] {
  const out: OrderedItem[] = [];
  for (const range of doc.added.get(change) ?? []) {
    const items = doc.stories.get(range.story)?.spaces.get(range.space) ?? [];
    for (const item of items)
      if (item.offset < range.to && item.offset + item.length > range.from) out.push(item);
  }

  return out;
}
