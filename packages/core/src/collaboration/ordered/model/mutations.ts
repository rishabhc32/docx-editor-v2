// Every write to an ordered document. Each one records how to take it back when an undo record
// is open, so an edit built from these writes is taken back exactly.

import {
  indexItem,
  labelIndex,
  labelUnindex,
  storyOfItem,
  suggestionsOf,
  unindexItem,
  itemAt,
} from './document.ts';
import { OrderedRefusal } from './refusal.ts';
import { recordUndo } from './undo-record.ts';
import { addTextBoxStories } from '../load/text-boxes.ts';
import type {
  OrderedAttribution,
  OrderedDocument,
  OrderedItem,
  OrderedItemLabels,
  OrderedStory,
  OrderedSuggestion,
  OrderedSuggestionStatus,
  OrderedUnitRange,
} from './types.ts';

const isHighSurrogate = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLowSurrogate = (code: number) => code >= 0xdc00 && code <= 0xdfff;

/**
 * Split a text item so that a piece starts at `offset` (a unit offset in its space). Returns
 * the two pieces, or null when `offset` is at an edge of the item (nothing to split).
 *
 * @public
 */
export function splitItem(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  offset: number
): [OrderedItem, OrderedItem] | null {
  const k = offset - item.offset;
  if (k <= 0 || k >= item.length) return null;
  if (item.kind !== 'text') throw new OrderedRefusal('not-text', item.kind);

  // A character outside the basic plane is two UTF-16 units: the engine refuses a place between
  // them, and one half alone is not valid XML
  if (isHighSurrogate(item.text.charCodeAt(k - 1)) && isLowSurrogate(item.text.charCodeAt(k)))
    throw new OrderedRefusal('splits-surrogate-pair', `${item.space}|${offset}`);
  const left: OrderedItem = { ...item, length: k, text: item.text.slice(0, k), leaf: undefined };
  const right: OrderedItem = {
    ...item,
    offset,
    length: item.length - k,
    text: item.text.slice(k),
    leaf: undefined,
  };
  story.tree.replace(item, [left, right]);
  unindexItem(doc, story, item);
  indexItem(doc, story, left);
  indexItem(doc, story, right);
  recordUndo(doc, () => {
    story.tree.insertAfter(left, [item]);
    for (const piece of [left, right]) {
      story.tree.remove(piece);
      unindexItem(doc, story, piece);
    }

    indexItem(doc, story, item);
  });
  return [left, right];
}

/**
 * The items that hold exactly the units `[start, end)` of a space, split out as needed.
 *
 * @public
 */
export function isolateUnits(
  doc: OrderedDocument,
  story: OrderedStory,
  space: string,
  start: number,
  end: number
): OrderedItem[] {
  const out: OrderedItem[] = [];
  for (let at = start; at < end; ) {
    let item = itemAt(story, { space, offset: at });
    if (!item) throw new OrderedRefusal('unknown-address', `${space}|${at}`);
    if (item.offset < at) item = splitItem(doc, story, item, at)?.[1] ?? item;
    if (item.offset + item.length > end) item = splitItem(doc, story, item, end)?.[0] ?? item;
    out.push(item);
    at = item.offset + item.length;
  }

  return out;
}

/**
 * A write added, removed or changed items these suggestions label, or their labels: each
 * suggestion's version goes up once.
 */
function touch(doc: OrderedDocument, ids: ReadonlySet<string>): void {
  for (const id of ids) {
    const suggestion = doc.suggestions.get(id);
    if (!suggestion) continue;
    suggestion.version += 1;
    recordUndo(doc, () => {
      suggestion.version -= 1;
    });
  }
}

/** The suggestions any of the items' labels name. */
function namedBy(items: readonly OrderedItem[]): Set<string> {
  const ids = new Set<string>();
  for (const item of items) for (const id of suggestionsOf(item)) ids.add(id);
  return ids;
}

const LABEL_KEYS = [
  'by',
  'dels',
  'fmts',
  'cancelledFmts',
  'props',
  'move',
  'retracted',
  'taken',
  'gone',
  'quiet',
  'within',
] as const;

function copyLabel<K extends keyof OrderedItemLabels>(
  into: Partial<OrderedItemLabels>,
  from: Partial<OrderedItemLabels>,
  key: K
): void {
  into[key] = from[key];
}

/**
 * Change an item's labels in place; the tree's sums and the label index follow.
 *
 * @public
 */
export function relabel(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  change: Partial<OrderedItemLabels>
): void {
  const old: Partial<OrderedItemLabels> = {};
  const next: Partial<OrderedItemLabels> = {};
  for (const key of LABEL_KEYS) {
    if (change[key] === undefined) continue;
    copyLabel(old, item, key);
    copyLabel(next, change, key);
  }

  const write = (labels: Partial<OrderedItemLabels>): void => {
    labelUnindex(doc, item);
    Object.assign(item, labels);
    labelIndex(doc, item);
    story.tree.update(item);
  };
  const named = suggestionsOf(item);
  write(next);
  recordUndo(doc, () => write(old));
  for (const id of suggestionsOf(item)) named.add(id);
  touch(doc, named);
}

/**
 * Put new items right after `after` (at the very start when null). An object that holds a text
 * box also adds that text box's story.
 *
 * @public
 */
export function insertItems(
  doc: OrderedDocument,
  story: OrderedStory,
  after: OrderedItem | null,
  items: readonly OrderedItem[]
): void {
  story.tree.insertAfter(after, items);
  for (const item of items) indexItem(doc, story, item);
  recordUndo(doc, () => {
    for (const item of items) {
      story.tree.remove(item);
      unindexItem(doc, story, item);
    }
  });
  touch(doc, namedBy(items));

  // An object that holds a text box brings the text box's story (group 9.9)
  addTextBoxStories(doc, story, items);
}

/**
 * Move consecutive items to right after `after`, as they are: no unit changes its address.
 */
export function moveItems(
  doc: OrderedDocument,
  story: OrderedStory,
  items: readonly OrderedItem[],
  after: OrderedItem
): void {
  const [first] = items;
  if (first === undefined) {
    return;
  }

  const before = story.tree.prev(first);
  const put = (at: OrderedItem | null): void => {
    for (const item of items) {
      story.tree.remove(item);
    }

    story.tree.insertAfter(at, items);
  };
  put(after);
  recordUndo(doc, () => put(before));
  touch(doc, namedBy(items));
}

/**
 * Put `next` in the place of `old`: consecutive items of one space whose units `next` holds
 * again, cut another way (typing that grows an item).
 * No unit changes its address.
 */
export function replaceItems(
  doc: OrderedDocument,
  story: OrderedStory,
  old: readonly OrderedItem[],
  next: readonly OrderedItem[]
): void {
  const [first, ...rest] = old;
  const last = next[next.length - 1];
  if (first === undefined || last === undefined) throw new RangeError('Nothing to replace.');
  story.tree.replace(first, next);
  for (const item of rest) story.tree.remove(item);
  for (const item of old) unindexItem(doc, story, item);
  for (const item of next) indexItem(doc, story, item);
  recordUndo(doc, () => {
    story.tree.insertAfter(last, old);
    for (const item of next) {
      story.tree.remove(item);
      unindexItem(doc, story, item);
    }

    for (const item of old) indexItem(doc, story, item);
  });
  touch(doc, namedBy([...old, ...next]));
}

/**
 * A new pending suggestion. Its id is `id`, or minted from the document's counter past every id
 * the document holds (a checkpoint keeps each suggestion's id, a minted one too).
 *
 * @throws OrderedRefusal `name-taken` when the document holds `id`: a change id used again after
 *   a fold, which the room no longer knows, mints the ids it minted before.
 * @public
 */
export function newSuggestion(
  doc: OrderedDocument,
  attribution: OrderedAttribution,
  id?: string
): OrderedSuggestion {
  if (id !== undefined && doc.suggestions.has(id)) {
    throw new OrderedRefusal('name-taken', id);
  }

  doc.counters.suggestion += 1;
  let sid = id ?? `S${doc.counters.suggestion}`;
  while (doc.suggestions.has(sid)) {
    doc.counters.suggestion += 1;
    sid = `S${doc.counters.suggestion}`;
  }

  const suggestion: OrderedSuggestion = {
    id: sid,
    attribution,
    status: 'pending',
    seq: nextSeq(doc),
    version: 0,
  };
  doc.suggestions.set(sid, suggestion);
  recordUndo(doc, () => doc.suggestions.delete(sid));
  return suggestion;
}

/** The next creation order number. Counters are restored when a record is taken back. */
export function nextSeq(doc: OrderedDocument): number {
  doc.counters.seq += 1;
  return doc.counters.seq;
}

/**
 * Change a suggestion's status. Every item it labels gets its sums again; nothing moves.
 *
 * @public
 */
export function setStatus(
  doc: OrderedDocument,
  suggestionId: string,
  status: OrderedSuggestionStatus
): void {
  const suggestion = doc.suggestions.get(suggestionId);
  if (!suggestion) throw new OrderedRefusal('unknown-suggestion', suggestionId);
  if (suggestion.status === status) return;
  const old = suggestion.status;
  const write = (value: OrderedSuggestionStatus): void => {
    suggestion.status = value;
    for (const item of doc.labelled.get(suggestionId) ?? [])
      storyOfItem(doc, item)?.tree.update(item);
  };
  write(status);
  recordUndo(doc, () => write(old));
}

/**
 * Accept a direct edit's suggestion when it is made, and mark it as a direct edit's: its
 * author can take it back exactly (`revert`).
 */
export function settleDirect(doc: OrderedDocument, suggestionId: string): void {
  const suggestion = doc.suggestions.get(suggestionId);
  if (!suggestion) throw new OrderedRefusal('unknown-suggestion', suggestionId);
  if (suggestion.direct !== true) {
    suggestion.direct = true;
    recordUndo(doc, () => {
      delete suggestion.direct;
    });
  }

  setStatus(doc, suggestionId, 'accepted');
}

/**
 * Every pending suggestion that labels an item `suggestionId` labels moves up a version: what
 * its card shows changed (a direct edit taken back or made again shows or hides those items).
 */
export function touchNeighbours(doc: OrderedDocument, suggestionId: string): void {
  const ids = namedBy([...(doc.labelled.get(suggestionId) ?? [])]);
  ids.delete(suggestionId);
  touch(doc, ids);
}

/**
 * Record units a change added to an earlier pending suggestion of its author (`doc.added`), for
 * the change's undo. An undo record takes it back with the edit.
 */
export function recordAdded(
  doc: OrderedDocument,
  change: string,
  ranges: readonly OrderedUnitRange[]
): void {
  const before = doc.added.get(change);
  doc.added.set(change, [...(before ?? []), ...ranges]);
  recordUndo(doc, () => {
    if (before) doc.added.set(change, before);
    else doc.added.delete(change);
  });
}

/** Forget the units a change added (`doc.added`): its undo took them back. */
export function forgetAdded(doc: OrderedDocument, change: string): void {
  const before = doc.added.get(change);
  if (!before) return;
  doc.added.delete(change);
  recordUndo(doc, () => doc.added.set(change, before));
}
