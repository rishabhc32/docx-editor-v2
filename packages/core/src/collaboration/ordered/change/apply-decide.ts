// A decision change: the organiser accepts or rejects review cards as they saw them. A decision
// only changes statuses, so decisions give one document in every order.

import { reviewItemKey } from '../../../store/store/review-items.ts';
import { revisionItemsOf } from '../../../store/store/review-reads.ts';
import { decidedAsBatch, planRevisionBatch } from '../../../store/store/revision-batch.ts';
import { revisionKey } from '../../../store/store/tree-op-tracked-adjacency.ts';
import { noCardRefusal, orderedCardOf, type OrderedCard } from '../card-of.ts';
import { composeStoryPart } from '../compose/compose.ts';
import { itemAt, storyOfItem } from '../model/document.ts';
import { setStatus } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import { textBoxHostOf } from '../model/text-boxes.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import type { OrderedIdMinter } from './ids.ts';
import { settleDecidedTables } from './settle-columns.ts';
import { settleHiddenNotes } from './settle-notes.ts';
import { settleLooseWords } from './settle-paragraphs.ts';
import type { OrderedDecideChange, OrderedSeenCard, OrderedSeenSuggestion } from './types.ts';

/** The cards a decision decides, and why the first card it leaves stays. */
interface DecisionVerdict {
  readonly cards: readonly OrderedCard[];
  readonly left: OrderedRefusal | null;
}

/** The card as it is now, when it is the card the organiser saw; else why it is not. */
function cardAsSeen(doc: OrderedDocument, seen: OrderedSeenCard): OrderedCard | OrderedRefusal {
  for (const { suggestion, version } of seen.suggestions) {
    const found = doc.suggestions.get(suggestion);
    if (found?.status !== 'pending') return new OrderedRefusal('not-pending', suggestion);
    if (found.version !== version)
      return new OrderedRefusal(
        'stale-decision',
        `${suggestion} is at version ${found.version}, not ${version}`
      );
  }

  const first = seen.suggestions[0]?.suggestion;
  const card = first === undefined ? null : orderedCardOf(doc, first);

  // Every suggestion is pending here, so no card means the list shows the first on none
  if (first !== undefined && card === null) return noCardRefusal(first);
  const ids = new Set(seen.suggestions.map((s) => s.suggestion));
  if (!card || card.suggestions.length !== ids.size || card.suggestions.some((id) => !ids.has(id)))
    return new OrderedRefusal('stale-decision', 'the card holds other suggestions now');
  return card;
}

/**
 * The part an item draws in, with the item of that part's story it draws with: for an item of a
 * text box, its object in the host part. Null for an item of no story, or of a box that went.
 */
function drawnAt(
  doc: OrderedDocument,
  item: OrderedItem
): { readonly part: string; readonly item: OrderedItem } | null {
  const story = storyOfItem(doc, item);
  if (story === null) {
    return null;
  }

  const box = textBoxHostOf(doc, story.name);
  if (box === null) {
    return { part: story.name, item };
  }

  const holder = doc.stories.get(box.host);
  const object = holder ? itemAt(holder, box.atom) : null;
  return object ? { part: box.host, item: object } : null;
}

/**
 * The cards the engine's own plan leaves, each with the engine's reason: the editor alone
 * refuses to decide them (a column deletion whose struck cells hold another record of its
 * author), so the room refuses them too, and no copy decides what another cannot. One plan for
 * each part the cards are in selects every item of the cards, as the decision takes them all. A
 * card stays when the plan leaves an item the editor decides with its dependencies
 * (`decidedAsBatch`); an item the editor resolves on its own sites is decided as it decides it.
 * A card a content control's lock keeps stays too (`locked`), as Word's Accept All leaves it.
 * Each part is drawn only where the plan reads it (`compose/region.ts`): the blocks the cards'
 * items are in, the block after each, and the blocks with a move.
 */
function engineLeft(
  doc: OrderedDocument,
  cards: readonly OrderedCard[],
  status: OrderedDecideChange['status']
): Map<OrderedCard, OrderedRefusal> {
  const action = status === 'accepted' ? 'accept' : 'reject';
  const cardOfMark = new Map<string, OrderedCard>();
  const regions = new Map<string, Set<OrderedItem>>();
  for (const card of cards) {
    for (const id of card.suggestions) {
      const { author, date } = doc.suggestions.get(id)?.attribution ?? { author: '' };
      for (const mark of revisionIdsOf(doc, id)) {
        cardOfMark.set(revisionKey(mark, author, date), card);
      }

      for (const item of doc.labelled.get(id) ?? []) {
        const at = drawnAt(doc, item);
        if (at === null) {
          continue;
        }

        const region = regions.get(at.part) ?? new Set<OrderedItem>();
        regions.set(at.part, region.add(at.item));
      }
    }
  }

  const left = new Map<OrderedCard, OrderedRefusal>();
  for (const [name, region] of regions) {
    const part = composeStoryPart(doc, name, { region }).part;
    const cardOfKey = new Map<string, OrderedCard>();
    const batched = new Set<string>();
    for (const item of revisionItemsOf(part)) {
      const marks = item.addresses.map(({ id, author, date }) => revisionKey(id, author, date));
      const card = marks.map((mark) => cardOfMark.get(mark)).find((c) => c !== undefined);
      if (card === undefined) {
        continue;
      }

      const key = reviewItemKey(item);
      cardOfKey.set(key, card);
      if (decidedAsBatch(item, part)) {
        batched.add(key);
      }
    }

    const plan = planRevisionBatch(part, action, [...cardOfKey.keys()], undefined, {
      skipLocked: true,
    });
    for (const { key, reason } of plan.result.skipped) {
      const card = cardOfKey.get(key);
      if (card === undefined || left.has(card)) {
        continue;
      }

      if (reason === 'locked') {
        left.set(
          card,
          new OrderedRefusal('locked', `${action}: a content-locked control's content`)
        );
      } else if (batched.has(key)) {
        left.set(card, new OrderedRefusal('engine-refused', `${action}: ${reason}`));
      }
    }
  }

  return left;
}

/**
 * Each card the change names, judged on the document as it is now, before any status turns.
 * As the host refuses a decision on a card that changed after the organiser saw it, the change
 * names each suggestion with the version seen; an edit moves the version, and joining or
 * leaving a card changes its set. The check is per card: a changed card stays and the others
 * are decided, so one late keystroke does not stop an Accept All.
 */
function judge(doc: OrderedDocument, change: OrderedDecideChange): DecisionVerdict {
  // A suggestion no copy ever held: the change is not one an organiser saw
  for (const card of change.cards)
    for (const { suggestion } of card.suggestions)
      if (!doc.suggestions.has(suggestion))
        throw new OrderedRefusal('unknown-suggestion', suggestion);
  const cards: OrderedCard[] = [];
  let left: OrderedRefusal | null = null;
  for (const seen of change.cards) {
    const card = cardAsSeen(doc, seen);
    if (card instanceof OrderedRefusal) left ??= card;
    else cards.push(card);
  }

  const refused = engineLeft(doc, cards, change.status);
  for (const card of cards) {
    left ??= refused.get(card) ?? null;
  }

  return { cards: cards.filter((card) => !refused.has(card)), left };
}

/**
 * The cards a decision decides on this document now: each card it names whose suggestions are
 * all pending at the versions the organiser saw. Empty when the decision would be refused.
 *
 * @public
 */
export function orderedDecidedCards(
  doc: OrderedDocument,
  change: OrderedDecideChange
): readonly OrderedCard[] {
  try {
    return judge(doc, change).cards;
  } catch (error) {
    if (error instanceof OrderedRefusal) return [];
    throw error;
  }
}

/**
 * Turn every suggestion of the cards to the status, then write into the items what the decision
 * changed: its tables (`settleDecidedTables`), the notes it hid (`settleHiddenNotes`) and the
 * words it left in no paragraph (`settleLooseWords`). Returns the ids decided.
 */
function decideCards(
  doc: OrderedDocument,
  cards: readonly OrderedCard[],
  status: OrderedDecideChange['status'],
  ids: OrderedIdMinter
): string[] {
  const decided = cards.flatMap((card) => card.suggestions);
  for (const id of decided) {
    setStatus(doc, id, status);
  }

  settleDecidedTables(doc, decided, ids);
  settleHiddenNotes(doc, decided);
  settleLooseWords(doc, decided);
  return decided;
}

/**
 * Decide each card the change names that is as the organiser saw it. Returns the ids decided.
 * When no card is as seen, the first card's reason refuses the change (`stale-decision`, or
 * `not-pending` for a card decided already), and the organiser sees the cards again.
 */
export function applyDecision(
  doc: OrderedDocument,
  change: OrderedDecideChange,
  ids: OrderedIdMinter
): string[] {
  const { cards, left } = judge(doc, change);
  if (cards.length === 0)
    throw left ?? new OrderedRefusal('stale-decision', 'the change names no card');
  return decideCards(doc, cards, change.status, ids);
}

/**
 * Decide one card as it is now, as a decision change decides it: refused with the engine's reason
 * when the editor alone does not decide it (`engineLeft`). Returns the ids decided.
 */
export function decideCard(
  doc: OrderedDocument,
  card: OrderedCard,
  status: OrderedDecideChange['status'],
  ids: OrderedIdMinter
): string[] {
  const left = engineLeft(doc, [card], status).get(card);
  if (left !== undefined) {
    throw left;
  }

  return decideCards(doc, [card], status, ids);
}

/**
 * What an organiser sends to decide the card that holds `suggestion`, as their copy shows it
 * now: every suggestion of the card with its version. Null when the suggestion is not pending.
 *
 * @public
 */
export function orderedCardSeen(
  doc: OrderedDocument,
  suggestion: string
): readonly OrderedSeenSuggestion[] | null {
  const card = orderedCardOf(doc, suggestion);
  if (!card) return null;
  return card.suggestions.map((id) => ({
    suggestion: id,
    version: doc.suggestions.get(id)?.version ?? 0,
  }));
}

/**
 * The cards that hold `suggestions`, each once, as a copy shows them now: what an organiser
 * sends to decide several cards in one change. Suggestions that are not pending are passed over.
 */
export function orderedSeenCards(
  doc: OrderedDocument,
  suggestions: readonly string[]
): OrderedSeenCard[] {
  const named = new Set<string>();
  const cards: OrderedSeenCard[] = [];
  for (const suggestion of suggestions) {
    if (named.has(suggestion)) continue;
    const seen = orderedCardSeen(doc, suggestion);
    if (!seen) continue;
    for (const entry of seen) named.add(entry.suggestion);
    cards.push({ suggestions: seen });
  }

  return cards;
}
