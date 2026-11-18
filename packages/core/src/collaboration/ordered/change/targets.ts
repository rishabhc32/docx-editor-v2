// What a change touches in a document: the items it edits in place, the carets it types at,
// and the suggestions it names. It checks that every address exists, in the change's part, on
// an item of the right kind, and that a range runs forward. It writes nothing.
//
// The room's policy reads only these items (design D9), and the room refuses a change whose
// targets do not resolve before the policy or the edit runs.

import { isHidden, itemAt, missingAddress, movedAwayBy, statusOf } from '../model/document.ts';
import { addressLabel } from '../model/refusal.ts';
import { itemsInside } from '../placement.ts';
import {
  isUnit,
  type OrderedContainerToken,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
  type OrderedSuggestion,
} from '../model/types.ts';
import { orderedKindSpecOf } from './kinds/registry.ts';
import type { OrderedTargetReader } from './kinds/spec.ts';
import { tableOpToken } from './table-op.ts';
import {
  isOrderedEditChange,
  type OrderedChange,
  type OrderedChangeAddress,
  type OrderedChangePlace,
  type OrderedCommentChange,
} from './types.ts';

/**
 * Why a change's targets do not resolve. Stable.
 *
 * - `unknown-story`: the part is not a story of the document.
 * - `unknown-address`: no unit at an address in the part.
 * - `other-file`: an address from another Word file of the room: the change was made against an
 *   older file (design D8).
 * - `unknown-suggestion`: a suggestion the change names is not in the document.
 * - `unknown-comment`: a comment a reply or a resolve names is not in the document.
 * - `wrong-target`: an address names a unit of the wrong kind (a join's address is not a
 *   paragraph mark, a row's is not a row's open item) or not its first unit.
 * - `invalid-range`: a range whose end is not after its start.
 *
 * @public
 */
export type OrderedTargetErrorCode =
  | 'unknown-story'
  | 'unknown-address'
  | 'other-file'
  | 'unknown-suggestion'
  | 'unknown-comment'
  | 'wrong-target'
  | 'invalid-range';

/**
 * A caret a change types at: where its new items go, as the item after them and the shown item
 * before them. A place inside struck text puts them past the struck text's end, as the engine
 * moves the caret there (placement rule 5).
 *
 * @public
 */
export interface OrderedCaret {
  /** The item holding the unit at the caret, or the paragraph's mark for an end place. */
  readonly at: OrderedItem;
  /**
   * The item holding the unit right before the caret, skipping markers with no width and hidden
   * items: `at` itself when the caret is inside it. Null at the story's start.
   */
  readonly before: OrderedItem | null;
}

/** What a change touches. @public */
export interface OrderedTouched {
  /** The change's story; null for a decision and a take-back. */
  readonly story: OrderedStory | null;
  /**
   * Items the change edits in place: a range's items (hidden ones too), target marks and tokens,
   * a column change's cell with the rows of its table, and a control value's control or a link
   * target's link with the units in it that the change strikes.
   */
  readonly items: readonly OrderedItem[];
  /** Places the change types or inserts at. */
  readonly carets: readonly OrderedCaret[];
  /** Suggestions the change names (`into`, a decision's, a take-back's). */
  readonly suggestions: readonly OrderedSuggestion[];
}

/** A change's targets, or why they do not resolve. @public */
export type OrderedTouchedResult =
  | { readonly ok: true; readonly touched: OrderedTouched }
  | { readonly ok: false; readonly code: OrderedTargetErrorCode; readonly detail: string };

class Unresolved extends Error {
  readonly code: OrderedTargetErrorCode;
  readonly detail: string;

  constructor(code: OrderedTargetErrorCode, detail: string) {
    super(code);
    this.code = code;
    this.detail = detail;
  }
}

function unitItem(
  doc: OrderedDocument,
  story: OrderedStory,
  address: OrderedChangeAddress
): OrderedItem {
  const item = itemAt(story, address);
  if (!item) throw new Unresolved(missingAddress(doc, address), addressLabel(address));
  return item;
}

/** The item at an address that must be the first unit of an item of `kind` (and `tokens`). */
function targetItem(
  doc: OrderedDocument,
  story: OrderedStory,
  address: OrderedChangeAddress,
  kind: OrderedItem['kind'],
  tokens?: readonly OrderedContainerToken[]
): OrderedItem {
  const item = unitItem(doc, story, address);
  const tokenOk = tokens === undefined || (item.token !== undefined && tokens.includes(item.token));
  if (item.kind !== kind || item.offset !== address.offset || !tokenOk)
    throw new Unresolved('wrong-target', addressLabel(address));
  return item;
}

/** The previous item that shows and has width or is a mark or a container edge. */
function shownBefore(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): OrderedItem | null {
  for (let prev = story.tree.prev(item); prev !== null; prev = story.tree.prev(prev))
    if (prev.kind !== 'zero' && !isHidden(doc, prev)) return prev;
  return null;
}

/**
 * A unit the view shows struck: a pending deletion's, not a pending move's source (it reads
 * plain).
 */
function shownStruck(doc: OrderedDocument, item: OrderedItem): boolean {
  if (!isUnit(item) || isHidden(doc, item)) return false;
  const deleter = item.dels.find((id) => statusOf(doc, id) === 'pending');
  return deleter !== undefined && !movedAwayBy(doc, item, deleter);
}

/** The shown unit right before an item in its paragraph, past markers and container edges. */
function unitBefore(doc: OrderedDocument, story: OrderedStory, item: OrderedItem) {
  for (let prev = story.tree.prev(item); prev !== null; prev = story.tree.prev(prev)) {
    if (prev.kind === 'mark') return null;
    if (isUnit(prev) && !isHidden(doc, prev)) return prev;
  }

  return null;
}

/**
 * Past struck text, as the engine moves a caret that has struck units on both sides of it
 * (placement rule 5, `siteAt`): right after the last struck unit of the paragraph's run of
 * them, which container edges do not end.
 */
function pastStruck(doc: OrderedDocument, story: OrderedStory, from: OrderedItem): OrderedCaret {
  let last = from;
  for (let next = story.tree.next(from); next !== null; next = story.tree.next(next)) {
    if (next.kind === 'zero' || isHidden(doc, next)) continue;
    if (shownStruck(doc, next)) last = next;
    else if (next.kind !== 'open' && next.kind !== 'close') break;
  }

  for (let next = story.tree.next(last); next !== null; next = story.tree.next(next))
    if (next.kind !== 'zero' && !isHidden(doc, next)) return { at: next, before: last };
  return { at: last, before: last };
}

/** A caret at a place: a binary search in its space and a few steps back in the tree. */
function caret(doc: OrderedDocument, story: OrderedStory, place: OrderedChangePlace): OrderedCaret {
  if ('mark' in place) {
    const mark = targetItem(doc, story, place.mark, 'mark');
    return { at: mark, before: shownBefore(doc, story, mark) };
  }

  const item = unitItem(doc, story, place);
  const inside = place.offset > item.offset;
  const previous = inside ? item : unitBefore(doc, story, item);
  if (previous && shownStruck(doc, previous) && shownStruck(doc, item))
    return pastStruck(doc, story, item);
  return { at: item, before: inside ? item : shownBefore(doc, story, item) };
}

/**
 * A column change's cell and the rows of its table: a column puts a cell in, or strikes one in,
 * every row. Rows of a table nested in a cell are that table's, not this one's.
 */
function columnItems(story: OrderedStory, cell: OrderedItem, out: OrderedItem[]): void {
  out.push(cell);
  let table: OrderedItem | null = null;
  let depth = 0;
  for (let at = story.tree.prev(cell); at !== null && table === null; at = story.tree.prev(at)) {
    if (at.token !== 'tbl') continue;
    if (at.kind === 'close') depth += 1;
    else if (at.kind === 'open' && depth === 0) table = at;
    else if (at.kind === 'open') depth -= 1;
  }

  if (table === null) return;
  depth = 0;
  for (let at = story.tree.next(table); at !== null; at = story.tree.next(at)) {
    if (at.token === 'tbl' && at.kind === 'close' && depth === 0) break;
    if (at.token === 'tbl') depth += at.kind === 'open' ? 1 : -1;
    else if (depth === 0 && at.kind === 'open' && at.token === 'tr') out.push(at);
  }
}

/**
 * A control value's control, or a link target's link, and the units in it the change strikes
 * (`setControlValue`, `changeLinkTarget`).
 */
function controlItems(story: OrderedStory, open: OrderedItem, out: OrderedItem[]): void {
  out.push(open);
  for (const item of itemsInside(story, open))
    if (isUnit(item) && !item.retracted && item.prompt !== true) out.push(item);
}

/** A place's position in units from the story's start. */
function position(
  doc: OrderedDocument,
  story: OrderedStory,
  place: OrderedChangePlace
): { item: OrderedItem; units: number } {
  const item =
    'mark' in place ? targetItem(doc, story, place.mark, 'mark') : unitItem(doc, story, place);
  const inside = 'mark' in place ? 0 : place.offset - item.offset;
  return { item, units: story.tree.before(item).units + inside };
}

/**
 * Add the items a range covers: from the one holding its first unit to the one before its end.
 * One step per item it covers.
 */
function rangeItems(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedChangePlace,
  to: OrderedChangePlace,
  out: OrderedItem[]
): void {
  const start = position(doc, story, from);
  const end = position(doc, story, to);
  if (end.units <= start.units)
    throw new Unresolved('invalid-range', 'the end is not after the start');
  const last = 'mark' in to || to.offset === end.item.offset ? story.tree.prev(end.item) : end.item;
  for (const item of story.tree.items(start.item)) {
    out.push(item);
    if (item === last) break;
  }
}

/** A new comment's range, or its caret when it covers no unit. */
function commentTargets(
  doc: OrderedDocument,
  story: OrderedStory,
  change: OrderedCommentChange,
  items: OrderedItem[],
  carets: OrderedCaret[]
): void {
  if (position(doc, story, change.from).units === position(doc, story, change.to).units) {
    carets.push(caret(doc, story, change.from));
    return;
  }

  rangeItems(doc, story, change.from, change.to, items);
}

function storyOf(doc: OrderedDocument, part: string): OrderedStory {
  const story = doc.stories.get(part);
  if (!story) throw new Unresolved('unknown-story', part.slice(0, 160));
  return story;
}

function suggestion(doc: OrderedDocument, id: string): OrderedSuggestion {
  const found = doc.suggestions.get(id);
  if (!found) throw new Unresolved('unknown-suggestion', id);
  return found;
}

function resolve(doc: OrderedDocument, change: OrderedChange): OrderedTouched {
  if (change.kind === 'decide' || change.kind === 'retract' || change.kind === 'revert') {
    const ids =
      change.kind === 'decide'
        ? change.cards.flatMap((card) => card.suggestions.map((seen) => seen.suggestion))
        : change.suggestions;
    return {
      story: null,
      items: [],
      carets: [],
      suggestions: ids.map((id) => suggestion(doc, id)),
    };
  }

  // A reply or a resolve names a thread, not places
  if (change.kind === 'comment-reply' || change.kind === 'comment-resolve') {
    if (!doc.comments.has(change.comment)) {
      throw new Unresolved('unknown-comment', change.comment.slice(0, 64));
    }

    return { story: null, items: [], carets: [], suggestions: [] };
  }

  const story = storyOf(doc, change.part);
  const items: OrderedItem[] = [];
  const carets: OrderedCaret[] = [];
  const suggestions =
    change.kind !== 'comment' && change.into !== undefined ? [suggestion(doc, change.into)] : [];
  switch (change.kind) {
    case 'insert':
    case 'insert-tab':
    case 'insert-break':
    case 'split':
    case 'insert-table':
    case 'insert-note':
      carets.push(caret(doc, story, change.place));
      break;
    case 'delete':
    case 'format':
      rangeItems(doc, story, change.from, change.to, items);
      break;
    case 'comment':
      commentTargets(doc, story, change, items, carets);
      break;
    case 'move':
      rangeItems(doc, story, change.from, change.to, items);
      carets.push(caret(doc, story, change.place));
      break;
    case 'join':
      items.push(targetItem(doc, story, change.mark, 'mark'));
      break;
    case 'paragraph-property':
    case 'mark-format':
      for (const mark of change.marks) items.push(targetItem(doc, story, mark, 'mark'));
      break;
    case 'table-property':
      for (const open of change.on)
        items.push(targetItem(doc, story, open, 'open', [tableOpToken(change.op)]));
      break;
    case 'insert-row':
    case 'delete-row':
      items.push(targetItem(doc, story, change.row, 'open', ['tr']));
      break;
    case 'insert-column':
    case 'delete-column':
      columnItems(story, targetItem(doc, story, change.cell, 'open', ['tc']), items);
      break;
    case 'control-value':
      controlItems(
        story,
        targetItem(doc, story, change.control, 'open', ['sdt', 'sdtBlock']),
        items
      );
      break;
    case 'link-target':
      // A new target strikes the link's words and writes them again (`changeLinkTarget`)
      controlItems(story, targetItem(doc, story, change.link, 'open', ['hyperlink']), items);
      break;
    default:
      // A kind a group added names its targets through its spec
      touchedByGroup(change, readerOf(doc, story, items, carets));
  }

  return { story, items, carets, suggestions };
}

/** What a change of a kind a group added touches, through its spec. */
function touchedByGroup(change: OrderedChange, reader: OrderedTargetReader): void {
  const spec = orderedKindSpecOf(change.kind);
  if (!spec || !isOrderedEditChange(change))
    throw new Unresolved('wrong-target', `no targets for ${change.kind.slice(0, 64)}`);
  spec.touches(change, reader);
}

/** The reader a group's spec names its targets through. */
function readerOf(
  doc: OrderedDocument,
  story: OrderedStory,
  items: OrderedItem[],
  carets: OrderedCaret[]
): OrderedTargetReader {
  return {
    caret: (place) => void carets.push(caret(doc, story, place)),
    range: (from, to) => rangeItems(doc, story, from, to, items),
    item: (address, kind, tokens) => {
      const found = targetItem(doc, story, address, kind, tokens);
      items.push(found);
      return found;
    },
  };
}

/**
 * What a change touches in a document, or why its targets do not resolve.
 *
 * @public
 */
export function orderedChangeTargets(
  doc: OrderedDocument,
  change: OrderedChange
): OrderedTouchedResult {
  try {
    return { ok: true, touched: resolve(doc, change) };
  } catch (error) {
    if (!(error instanceof Unresolved)) throw error;
    return { ok: false, code: error.code, detail: error.detail };
  }
}
