// Undo in a room is a new change that takes back the author's own suggestion, so it never
// touches another person's text.

import { itemsAddedBy, storyOfItem, statusOf, suggestionsOf } from '../model/document.ts';
import { forgetAdded, isolateUnits, relabel } from '../model/mutations.ts';
import { sameMaker, type OrderedMaker } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type {
  OrderedDocument,
  OrderedItem,
  OrderedItemLabels,
  OrderedStory,
} from '../model/types.ts';

/** Whether a pending suggestion other than `own` labels the item. */
export function othersPending(doc: OrderedDocument, item: OrderedItem, own: string): boolean {
  const pending = (id: string) => id !== own && statusOf(doc, id) === 'pending';
  return (
    (item.by !== null && pending(item.by)) ||
    item.dels.some(pending) ||
    item.fmts.some((f) => pending(f.suggestion)) ||
    item.props.some((r) => pending(r.suggestion))
  );
}

/**
 * Format or property records without the suggestion `id`'s. A record it took from another
 * suggestion (policy 2) goes back to that suggestion while it is pending.
 */
function recordsWithout<T extends { readonly suggestion: string; readonly replaced?: string }>(
  doc: OrderedDocument,
  records: readonly T[],
  id: string
): T[] {
  const out: T[] = [];
  for (const record of records) {
    if (record.suggestion !== id) {
      out.push(record);
      continue;
    }

    if (record.replaced !== undefined && statusOf(doc, record.replaced) === 'pending') {
      const { replaced, ...rest } = record;

      // Still a T: `replaced` is optional on every record type
      out.push({ ...rest, suggestion: replaced } as T);
    }
  }

  return out;
}

/**
 * Take back `suggestion` for `maker`: what it inserted is taken back (it never existed for
 * anyone else), and its deletions, formats and property records come off the items. Refused
 * when it is not theirs, when it is decided, or when a container it inserted (a row, a cell, a
 * note) holds another author's change.
 */
export function retractSuggestion(
  doc: OrderedDocument,
  suggestion: string,
  maker: OrderedMaker
): void {
  const found = doc.suggestions.get(suggestion);
  if (!found) throw new OrderedRefusal('unknown-suggestion', suggestion);
  if (found.status !== 'pending') throw new OrderedRefusal('not-pending', suggestion);
  if (!sameMaker(found.attribution, maker)) {
    throw new OrderedRefusal('not-own-suggestion', suggestion);
  }

  const items = [...(doc.labelled.get(suggestion) ?? [])];
  const moves = new Set<string>();
  for (const item of items) {
    if (item.by === suggestion && item.move !== null) moves.add(item.move);

    // Another author's change inside a container this suggestion inserted would go with it;
    // an item already taken back holds no change
    if (
      !item.retracted &&
      item.within.by.includes(suggestion) &&
      othersPending(doc, item, suggestion)
    )
      throw new OrderedRefusal(
        'others-in-row',
        "a new container that holds another author's change"
      );
  }

  for (const item of items) {
    const story = storyOfItem(doc, item);
    if (!story) continue;
    const change: Partial<OrderedItemLabels> = {};
    if (item.by === suggestion || item.within.by.includes(suggestion)) {
      change.retracted = true;

      // Taken back whole: no undo of an earlier deletion shows it again
      if (item.taken) change.taken = null;
    }

    if (item.dels.includes(suggestion)) {
      change.dels = item.dels.filter((id) => id !== suggestion);
      if (item.move !== null && moves.has(item.move)) change.move = null;
    }

    if (item.fmts.some((f) => f.suggestion === suggestion))
      change.fmts = recordsWithout(doc, item.fmts, suggestion);

    // The records this suggestion's press cancelled (`withPressFormat`) come back, as the press
    // never was
    if (item.cancelledFmts.some((f) => f.suggestion === suggestion)) {
      const cancelled = item.cancelledFmts.filter((f) => f.suggestion === suggestion);
      change.fmts = [...(change.fmts ?? item.fmts), ...recordsWithout(doc, cancelled, suggestion)];
      change.cancelledFmts = item.cancelledFmts.filter((f) => f.suggestion !== suggestion);
    }

    if (item.props.some((r) => r.suggestion === suggestion))
      change.props = recordsWithout(doc, item.props, suggestion);
    if (item.within.dels.includes(suggestion))
      change.within = {
        by: item.within.by,
        dels: item.within.dels.filter((id) => id !== suggestion),
      };
    relabel(doc, story, item, change);
  }
}

/**
 * Take back exactly what `changes` of an author did to the author's own pending suggestions:
 * the undo of one editor commit, as the editor's undo takes it back. The items they took back
 * show again (`taken`), the units they typed into those suggestions are taken back
 * (`doc.added`), and the formats they gave those words come off (a format record's `change`).
 * The validator keeps `changes` to the sender's own namespace. Refused while a suggestion that
 * inserted an item it changes is decided: a decision is final. Returns the suggestions the items
 * carry.
 */
export function undoOwnEdits(doc: OrderedDocument, changes: readonly string[]): string[] {
  const undone = new Set(changes);
  const edited = new Set(changes.flatMap((change) => [...(doc.ownEdited.get(change) ?? [])]));
  const added = changes.flatMap((change) => itemsAddedBy(doc, change));
  for (const item of [...edited, ...added]) {
    for (const inserter of [item.by, ...item.within.by]) {
      if (inserter !== null && statusOf(doc, inserter) !== 'pending') {
        throw new OrderedRefusal('not-pending', inserter);
      }
    }
  }

  const carried = new Set<string>();
  const write = (story: OrderedStory, item: OrderedItem, change: Partial<OrderedItemLabels>) => {
    relabel(doc, story, item, change);
    for (const id of suggestionsOf(item)) carried.add(id);
  };
  for (const item of edited) {
    const story = storyOfItem(doc, item);
    if (!story) continue;
    const change: Partial<OrderedItemLabels> = {};
    if (item.taken && undone.has(item.taken)) {
      change.retracted = false;
      change.taken = null;
    }

    const fmts = item.fmts.filter(
      (record) => record.kind !== 'properties' || !undone.has(record.change ?? '')
    );
    if (fmts.length !== item.fmts.length) change.fmts = fmts;
    write(story, item, change);
  }

  for (const change of changes) {
    for (const range of doc.added.get(change) ?? []) {
      const story = doc.stories.get(range.story);
      if (!story) continue;
      for (const item of isolateUnits(doc, story, range.space, range.from, range.to))
        if (!item.retracted) write(story, item, { retracted: true });
    }

    forgetAdded(doc, change);
  }

  return [...carried];
}
