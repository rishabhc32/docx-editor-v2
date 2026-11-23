// One suggestion's review card, found from its own items and their neighbours, so a decision
// costs what its card and the units around it hold, not what the document holds.

import { isNotesStory, leftOutNoteItems } from './compose/note-references.ts';
import { promptUnitHidden } from './compose/paragraph.ts';
import {
  innermostNewContainer,
  inserter,
  isHidden,
  isPending,
  oneMaker,
  storyOfItem,
} from './model/document.ts';
import { OrderedRefusal } from './model/refusal.ts';
import { textBoxGone } from './model/text-boxes.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from './model/types.ts';

/** One review card: the pending suggestions one decision applies to. @public */
export interface OrderedCard {
  /** The card's first suggestion, by creation order. */
  readonly id: string;
  /** Its suggestions, by creation order. */
  readonly suggestions: readonly string[];
  /** The first suggestion's author. */
  readonly author: string;
}

/** The pending suggestions that change an item without inserting it. */
export function changers(doc: OrderedDocument, item: OrderedItem): string[] {
  const out = new Set<string>();
  for (const id of item.dels) {
    if (isPending(doc, id)) {
      out.add(id);
    }
  }

  for (const record of item.fmts) {
    if (isPending(doc, record.suggestion)) {
      out.add(record.suggestion);
    }
  }

  for (const record of item.props) {
    if (isPending(doc, record.suggestion)) {
      out.add(record.suggestion);
    }
  }

  return [...out];
}

export const authorOf = (doc: OrderedDocument, id: string): string =>
  doc.suggestions.get(id)?.attribution.author ?? '';

/**
 * A unit a card can sit on: what shows, with a width or a paragraph's or container's edge. A
 * link's edge is none: Word and the editor mark a link's words, never the link (no `w:ins` holds
 * one), so a card reads the same live, where the room labels the edges, and after a reopen.
 */
export const isCardUnit = (doc: OrderedDocument, item: OrderedItem): boolean =>
  item.kind !== 'zero' && item.token !== 'hyperlink' && !isOrderedItemHidden(doc, item);

/**
 * Whether an item shows nowhere now: hidden for good (`isHidden`: its container went, its author
 * took it back, a decision took it away), or a unit of a prompt that content took the place of
 * (`hiddenPrompt`). Cards read units by this rule, and a fold, which keeps such a prompt in the
 * glossary only, changes nothing it says.
 *
 * @public
 */
export function isOrderedItemHidden(doc: OrderedDocument, item: OrderedItem): boolean {
  return isHidden(doc, item) || hiddenPrompt(doc, item);
}

/**
 * A unit of a prompt that content took the place of. Word and the engine drop a prompt's runs,
 * with any suggestion on them, when content fills the control (placement rule 7), and a saved
 * file keeps the prompt in the glossary only, so a suggestion there shows on no card. Taking the
 * content back shows the prompt, and the card, again. A prompt's marks hide in the model
 * (`gone`). A control that such a prompt held goes with it, its prompt too.
 */
function hiddenPrompt(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.prompt !== true || (item.kind !== 'text' && item.kind !== 'atom')) return false;
  const story = storyOfItem(doc, item);
  return story !== null && promptUnitHidden(doc, story, item);
}

/**
 * One read of a document's cards: the units a card sits on, the same for the list
 * (`orderedCards`) and for one card (`orderedCardOf`). A note compose leaves out (its references
 * show nowhere) and a text box whose object went hold no unit, as a saved file has neither. Those
 * notes are read once, and only when a search reaches an item of a notes part.
 */
export class CardRead {
  /** The one pending insertion every unit of a suggestion sits in, by suggestion (`insideOne`). */
  readonly inside = new Map<string, string | null>();
  #leftOut: ReadonlySet<OrderedItem> | null = null;
  /** Whether each story read so far is a text box whose object went (`textBoxGone`). */
  readonly #boxesGone = new Map<OrderedStory, boolean>();

  constructor(readonly doc: OrderedDocument) {}

  /**
   * Whether an item is in a note compose leaves out (`leftOutNoteItems`) or in a text box whose
   * object went (`textBoxGone`).
   */
  isLeftOut(item: OrderedItem, story = storyOfItem(this.doc, item)): boolean {
    if (story === null) {
      return false;
    }

    if (isNotesStory(story)) {
      this.#leftOut ??= leftOutNoteItems(this.doc);
      return this.#leftOut.has(item);
    }

    let gone = this.#boxesGone.get(story);
    if (gone === undefined) {
      gone = textBoxGone(this.doc, story.name);
      this.#boxesGone.set(story, gone);
    }

    return gone;
  }

  /** A unit a card can sit on: what shows (`isCardUnit`), outside what compose leaves out. */
  isUnit(item: OrderedItem, story?: OrderedStory | null): boolean {
    return isCardUnit(this.doc, item) && !this.isLeftOut(item, story);
  }

  /** Whether a suggestion is on no unit only because its items are in what compose leaves out. */
  onlyLeftOut(suggestion: string): boolean {
    let leftOut = false;
    for (const item of this.doc.labelled.get(suggestion) ?? []) {
      if (this.isUnit(item)) {
        return false;
      }

      leftOut ||= this.isLeftOut(item);
    }

    return leftOut;
  }
}

/**
 * Why a decision on a pending suggestion that is on no card is refused: it is only in a note
 * whose references show nowhere or in a text box whose object went, so neither the list nor a
 * saved file shows it.
 */
export const noCardRefusal = (suggestion: string): OrderedRefusal =>
  new OrderedRefusal(
    'hidden-target',
    `${suggestion} is only in a note or a text box that shows nowhere`
  );

const TABLE_CONTAINERS: ReadonlySet<string> = new Set(['tbl', 'tr', 'tc']);

/**
 * The innermost pending new row or cell around a unit, when the same author inserted it with
 * another suggestion than the unit's own (join 4). A table container's own edge is passed over:
 * a new column's cell in a new row keeps the column's card.
 */
export function ownContainer(doc: OrderedDocument, item: OrderedItem): string | null {
  const by = item.by;
  if (!isPending(doc, by)) {
    return null;
  }

  if (item.token !== undefined && TABLE_CONTAINERS.has(item.token)) return null;
  const container = innermostNewContainer(doc, item);
  if (container === null || container === by) {
    return null;
  }

  return oneMaker(doc, container, by) ? container : null;
}

/**
 * Whether a run of one suggestion's units starts a whole row: a new row is a change of its own,
 * never inside the rows around it (join 3), as the host reads a row as its own paragraphs.
 */
export const startsRow = (item: OrderedItem): boolean =>
  item.kind === 'open' && item.token === 'tr';

/** The one pending insertion every unit of `id` sits in, or null (kept in `read.inside`). */
function insideOne(read: CardRead, id: string) {
  const known = read.inside.get(id);
  if (known !== undefined) return known;
  let one: string | null | undefined;
  for (const item of read.doc.labelled.get(id) ?? []) {
    if (!read.isUnit(item)) continue;
    const by = inserter(read.doc, item);
    one = one === undefined || one === by ? by : null;
    if (one === null) break;
  }

  read.inside.set(id, one ?? null);
  return one ?? null;
}

/**
 * Whether change `id` on `unit` joins the card of `by`, the unit's inserter (join 1). A table's,
 * a row's or a cell's property record has no text, so the host reads it as a collapsed range,
 * which joins an insertion only when the whole change is inside it (`isInside`: a change on a
 * new row or on new cells only). A row height over another person's new cell keeps its own card.
 */
export function joinsInsertion(read: CardRead, unit: OrderedItem, id: string, by: string): boolean {
  const propertyOnly =
    unit.kind === 'open' &&
    unit.token !== undefined &&
    TABLE_CONTAINERS.has(unit.token) &&
    !unit.dels.includes(id) &&
    !unit.fmts.some((record) => record.suggestion === id);
  return !propertyOnly || insideOne(read, id) === by;
}

function unitBefore(read: CardRead, story: OrderedStory, item: OrderedItem) {
  for (let prev = story.tree.prev(item); prev; prev = story.tree.prev(prev))
    if (read.isUnit(prev, story)) return prev;
  return null;
}

function unitAfter(read: CardRead, story: OrderedStory, item: OrderedItem) {
  for (let next = story.tree.next(item); next; next = story.tree.next(next))
    if (read.isUnit(next, story)) return next;
  return null;
}

/** The last unit of the run of units `by` inserted that starts or goes on at `from`. */
function runEnd(read: CardRead, story: OrderedStory, from: OrderedItem, by: string) {
  let end = from;
  for (let next = unitAfter(read, story, end); next && inserter(read.doc, next) === by; )
    [end, next] = [next, unitAfter(read, story, next)];
  return end;
}

function runStart(read: CardRead, story: OrderedStory, from: OrderedItem, by: string) {
  let start = from;
  for (let prev = unitBefore(read, story, start); prev && inserter(read.doc, prev) === by; )
    [start, prev] = [prev, unitBefore(read, story, prev)];
  return start;
}

/** One card search: the units it visited, and how a suggestion it finds joins the card. */
interface CardSearch {
  readonly read: CardRead;
  /** The first unit of each run of one suggestion's units that join 3 looked at. */
  readonly runs: Set<OrderedItem>;
  readonly add: (id: string) => void;
}

/** One unit of suggestion `x` that a search visits, with its inserter, changers and neighbours. */
interface UnitVisit {
  readonly story: OrderedStory;
  readonly unit: OrderedItem;
  readonly x: string;
  readonly by: string | null;
  readonly changed: readonly string[];
  readonly before: OrderedItem | null;
  readonly after: OrderedItem | null;
}

/**
 * The suggestions one unit of `x` joins `x` with. The joins, over the shown units of a story
 * (zero-width markers and hidden items are passed over), as the host's `cardsOf` rule reads them:
 *
 *  1. On one unit: a change by one author on another author's pending insertion (a table
 *     property change only when all of it is inside that insertion, `joinsInsertion`).
 *  2. On two units side by side: a deletion next to an insertion by the same author.
 *  3. Over a run of units one suggestion inserted: when the units on both sides of the run are
 *     another author's one pending insertion, the run joins it. A run that starts a row does not.
 *  4. On one unit: what an author inserted inside their own pending new row or cell, in a later
 *     change, joins that row's or cell's card (`ownContainer`), as the editor alone shows a new
 *     row and the words typed in it as one card, and as 2.x decided a row with what it holds.
 */
function joinsAt(search: CardSearch, story: OrderedStory, unit: OrderedItem, x: string): void {
  const { read, add } = search;
  const doc = read.doc;

  // 4. x's words in the same author's new row or cell, or that author's words in x's (the label
  // index lists a unit under every new container around it, so both sides visit it)
  const own = ownContainer(doc, unit);
  if (own === x && unit.by !== null) {
    add(unit.by);
  } else if (own !== null && unit.by === x) {
    add(own);
  }

  const visit: UnitVisit = {
    story,
    unit,
    x,
    by: inserter(doc, unit),
    changed: changers(doc, unit),
    before: unitBefore(read, story, unit),
    after: unitAfter(read, story, unit),
  };
  if (visit.by === x) {
    joinsAsInserter(search, visit);
  }

  if (visit.changed.includes(x)) {
    joinsAsChanger(search, visit);
  }
}

/** Joins 1 to 3 of a unit that `x` inserted. */
function joinsAsInserter(search: CardSearch, visit: UnitVisit): void {
  const { read, add, runs } = search;
  const doc = read.doc;
  const { story, unit, x, after } = visit;

  // 1. Another author's change on x's insertion
  for (const id of visit.changed) {
    if (!oneMaker(doc, id, x) && joinsInsertion(read, unit, id, x)) {
      add(id);
    }
  }

  // 2. The same author's deletion right before or after x's insertion
  for (const side of [visit.before, after]) {
    for (const id of side?.dels ?? []) {
      if (isPending(doc, id) && oneMaker(doc, id, x)) {
        add(id);
      }
    }
  }

  // 3. x's run inside another author's insertion
  const start = runStart(read, story, unit, x);
  if (!runs.has(start) && !startsRow(start)) {
    runs.add(start);
    const outerItem = unitBefore(read, story, start);
    const afterItem = unitAfter(read, story, runEnd(read, story, start, x));
    const outer = outerItem ? inserter(doc, outerItem) : null;
    if (
      outer !== null &&
      afterItem &&
      inserter(doc, afterItem) === outer &&
      !oneMaker(doc, outer, x)
    ) {
      add(outer);
    }
  }

  // 3. Another author's run inside x's insertion, starting right after this unit
  const next = after ? inserter(doc, after) : null;
  if (after && next !== null && next !== x && !oneMaker(doc, next, x) && !startsRow(after)) {
    const close = unitAfter(read, story, runEnd(read, story, after, next));
    if (close && inserter(doc, close) === x) {
      add(next);
    }
  }
}

/** Joins 1 and 2 of a unit that `x` changes without inserting it. */
function joinsAsChanger(search: CardSearch, visit: UnitVisit): void {
  const { read, add } = search;
  const doc = read.doc;
  const { unit, x, by } = visit;

  // 1. x's change on another author's insertion
  if (by !== null && !oneMaker(doc, by, x) && joinsInsertion(read, unit, x, by)) {
    add(by);
  }

  // 2. x's deletion right before or after the same author's insertion
  if (!unit.dels.includes(x)) {
    return;
  }

  for (const side of [visit.before, visit.after]) {
    const sideBy = side ? inserter(doc, side) : null;
    if (sideBy !== null && oneMaker(doc, sideBy, x)) {
      add(sideBy);
    }
  }
}

/** A card of these suggestions, by creation order: the first one's id and author. */
export function cardFrom(doc: OrderedDocument, ids: readonly string[]): OrderedCard | null {
  const seqOf = (id: string) => doc.suggestions.get(id)?.seq ?? 0;
  const sorted = [...ids].sort((a, b) => seqOf(a) - seqOf(b));
  const first = sorted[0];
  if (first === undefined) {
    return null;
  }

  return { id: first, suggestions: sorted, author: authorOf(doc, first) };
}

/**
 * The review card that holds a pending suggestion: the one `orderedCards` gives, or the
 * suggestion alone when it shows nowhere. Null when the suggestion is not pending, or when it is
 * only in a note whose references show nowhere or in a text box whose object went, which
 * `orderedCards` leaves off every card too. The search visits the card's own items and the units
 * around them only (and, for a note, which notes show).
 *
 * @public
 */
export function orderedCardOf(doc: OrderedDocument, suggestion: string): OrderedCard | null {
  const read = new CardRead(doc);
  if (!isPending(doc, suggestion) || read.onlyLeftOut(suggestion)) {
    return null;
  }

  const members = new Set([suggestion]);
  const queue = [suggestion];
  const search: CardSearch = {
    read,
    runs: new Set(),
    add: (id) => {
      if (members.has(id)) {
        return;
      }

      members.add(id);
      queue.push(id);
    },
  };
  for (let x = queue.pop(); x !== undefined; x = queue.pop()) {
    for (const item of doc.labelled.get(x) ?? []) {
      if (!isCardUnit(doc, item)) {
        continue;
      }

      const story = storyOfItem(doc, item);
      if (story && !read.isLeftOut(item, story)) {
        joinsAt(search, story, item, x);
      }
    }
  }

  return cardFrom(doc, [...members]);
}
