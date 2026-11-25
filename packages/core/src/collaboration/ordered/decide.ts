// Decisions and review cards. A decision changes statuses, and writes what it changed into the
// items as the room's does (`change/apply-decide.ts`). A card is what one decision applies to
// (`joinsAt` in card-of.ts).

import {
  authorOf,
  CardRead,
  cardFrom,
  changers,
  insertionBeside,
  isCardUnit,
  joinsInsertion,
  noCardRefusal,
  orderedCardOf,
  ownContainer,
  startsRow,
  type OrderedCard,
} from './card-of.ts';
import { decideCard } from './change/apply-decide.ts';
import { orderedChangeId, orderedIdMinter } from './change/ids.ts';
import { inserter, isHidden, isPending, oneMaker } from './model/document.ts';
import { OrderedRefusal } from './model/refusal.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from './model/types.ts';

export type { OrderedCard } from './card-of.ts';

/** Union-find over suggestion ids, with the order the walk first tracked each one. */
class Groups {
  private readonly parent = new Map<string, string>();
  private readonly seen = new Set<string>();
  /** The suggestions with a card, in the order the walk first met them. */
  readonly tracked: string[] = [];

  private add(id: string): void {
    if (!this.parent.has(id)) this.parent.set(id, id);
  }

  /** Add a suggestion that has a card. */
  track(id: string): void {
    this.add(id);
    if (!this.seen.has(id)) {
      this.seen.add(id);
      this.tracked.push(id);
    }
  }

  find(id: string): string {
    let root = id;
    for (
      let up = this.parent.get(root);
      up !== undefined && up !== root;
      up = this.parent.get(root)
    )
      root = up;

    // Path compression keeps later finds short
    for (let at = id; at !== root; ) {
      const up = this.parent.get(at) ?? root;
      this.parent.set(at, root);
      at = up;
    }

    return root;
  }

  join(a: string, b: string): void {
    this.add(a);
    this.add(b);
    const ra = this.find(a);
    const rb = this.find(b);
    if (ra !== rb) this.parent.set(rb, ra);
  }
}

/**
 * The review cards of every pending suggestion that shows, in document order of their first
 * unit, then one card for each pending tracked change outside the stories (a style's), in the
 * part's order.
 *
 * @public
 */
export function orderedCards(doc: OrderedDocument): OrderedCard[] {
  const groups = new Groups();
  const read = new CardRead(doc);
  for (const story of doc.stories.values()) {
    const units = cardUnits(read, story, groups);
    joinNeighbours(read, units, groups);
    joinRunsInside(doc, units, groups);
  }

  const cards = cardsOfGroups(doc, groups);

  // A change record outside the stories has no items: each is a card of its own
  for (const id of doc.partRevisions.keys()) {
    if (isPending(doc, id)) {
      cards.push({ id, suggestions: [id], author: authorOf(doc, id) });
    }
  }

  return cards;
}

/**
 * The units of a story a card can sit on. Cards are read from what shows now: items hidden for
 * good, the notes left out with their references and the text boxes gone with their objects are
 * passed over, as a saved and reopened file has none of them, so a card is the same live and
 * after a reopen.
 */
function cardUnits(read: CardRead, story: OrderedStory, groups: Groups): OrderedItem[] {
  const doc = read.doc;
  const units: OrderedItem[] = [];
  for (const item of story.tree.items()) {
    if (read.isLeftOut(item, story)) {
      continue;
    }

    if (isCardUnit(doc, item)) {
      units.push(item);
    } else if (item.kind === 'zero' && !isHidden(doc, item)) {
      // A run with nothing in it but its properties has no width: a file's change of them is a
      // card (of its own, unless its revision also changes words), as the editor alone shows it
      for (const record of item.fmts) {
        if (isPending(doc, record.suggestion)) {
          groups.track(record.suggestion);
        }
      }
    }
  }

  return units;
}

/** Joins on one unit and on two units side by side (joins 1, 2 and 4). */
function joinNeighbours(read: CardRead, units: readonly OrderedItem[], groups: Groups): void {
  const doc = read.doc;
  for (const [k, item] of units.entries()) {
    const by = inserter(doc, item);
    if (by !== null) {
      groups.track(by);
    }

    // What a person inserts in their own new row or cell joins its card
    const own = ownContainer(doc, item);
    if (own !== null && by !== null) {
      groups.join(own, by);
    }

    for (const id of changers(doc, item)) {
      groups.track(id);

      // A change on another person's pending insertion joins its card
      if (by !== null && !oneMaker(doc, id, by) && joinsInsertion(read, item, id, by)) {
        groups.join(by, id);
      }
    }

    const next = units[k + 1];
    if (!next) {
      continue;
    }

    // A deletion next to an insertion by the same author: one card
    for (const deleter of item.dels) {
      const nextBy = insertionBeside(doc, deleter, next);
      if (nextBy !== null) {
        groups.join(deleter, nextBy);
      }
    }

    for (const deleter of next.dels) {
      const beside = insertionBeside(doc, deleter, item);
      if (beside !== null) {
        groups.join(beside, deleter);
      }
    }
  }
}

/** Text typed between two units of another person's insertion joins that insertion's card. */
function joinRunsInside(doc: OrderedDocument, units: readonly OrderedItem[], groups: Groups): void {
  for (let start = 0; start < units.length; ) {
    const head = units[start];
    const by = head ? inserter(doc, head) : null;
    let end = start + 1;
    while (end < units.length && by !== null && inserter(doc, units[end] as OrderedItem) === by) {
      end += 1;
    }

    const before = units[start - 1];
    const after = units[end];
    const outer = before ? inserter(doc, before) : null;
    if (
      by !== null &&
      head &&
      !startsRow(head) &&
      outer !== null &&
      after &&
      inserter(doc, after) === outer &&
      outer !== by &&
      !oneMaker(doc, outer, by)
    ) {
      groups.join(outer, by);
    }

    start = end;
  }
}

/** One card for each group, in the order the walk first tracked one of its suggestions. */
function cardsOfGroups(doc: OrderedDocument, groups: Groups): OrderedCard[] {
  const members = new Map<string, string[]>();
  for (const id of groups.tracked) {
    const root = groups.find(id);
    const list = members.get(root);
    if (list) {
      list.push(id);
    } else {
      members.set(root, [id]);
    }
  }

  const cards: OrderedCard[] = [];
  for (const list of members.values()) {
    const card = cardFrom(doc, list);
    if (card) {
      cards.push(card);
    }
  }

  return cards;
}

/**
 * Decide the card that holds `suggestion`, as the room's decision does (`decideCard`): every
 * pending suggestion in it takes the status, and the items hold what the decision changed. A card
 * the editor alone does not decide is refused with the engine's reason (`engine-refused`).
 * Returns the suggestions decided. A pending suggestion that shows nowhere is its own card, but
 * one only in a note whose references show nowhere or in a text box whose object went is on no
 * card and is refused (`hidden-target`). The card is found from the suggestion's own items
 * (`orderedCardOf`), so a decision costs what its card holds. Inside an open undo record, the
 * record takes the decision back.
 *
 * @public
 */
export function decideOrdered(
  doc: OrderedDocument,
  suggestion: string,
  decision: 'accept' | 'reject'
): readonly string[] {
  const found = doc.suggestions.get(suggestion);
  if (!found) throw new OrderedRefusal('unknown-suggestion', suggestion);
  if (found.status !== 'pending') throw new OrderedRefusal('not-pending', suggestion);
  const card = orderedCardOf(doc, suggestion);
  if (card === null) throw noCardRefusal(suggestion);
  const status = decision === 'accept' ? 'accepted' : 'rejected';

  // As the room's decision: one change id per decision, from the creation counter it moves
  const ids = orderedIdMinter(orderedChangeId('decide', doc.counters.seq + 1));
  return decideCard(doc, card, status, ids);
}
