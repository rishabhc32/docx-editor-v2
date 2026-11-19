// Who typed the units of a pending insertion: the namespace of the room change that typed them.

import { CardRead } from '../card-of.ts';
import { parseOrderedChangeId } from '../change/ids.ts';
import { byStoryOrder } from '../load/load.ts';
import { itemAt, mainStory } from './document.ts';
import { isTextBoxStoryName, textBoxHostOf } from './text-boxes.ts';
import { isUnit, type OrderedDocument, type OrderedItem, type OrderedStory } from './types.ts';

/** Units of one space that one room change typed: the offsets from `from` up to `to`. */
export interface TypedRun {
  readonly from: number;
  readonly to: number;
  readonly typer: string;
}

/** The typers a checkpoint kept for a document loaded from it: runs of typed units by space. */
const LOADED = new WeakMap<OrderedDocument, ReadonlyMap<string, readonly TypedRun[]>>();

/**
 * The namespace of the room change that typed the unit at `offset` (in the item's space) of an
 * item a pending suggestion inserted, such as an organiser's words typed directly inside a
 * party's insertion. A copy of the room and a copy opened from its checkpoint give the same
 * answer. Null for a unit of any other item, for one inside a new row or note (its container's
 * mark covers it), and for one no room change typed (a file's own).
 *
 * @public
 */
export function orderedUnitTyper(
  doc: OrderedDocument,
  item: OrderedItem,
  offset: number
): string | null {
  const inside = offset >= item.offset && offset < item.offset + item.length;
  if (!isUnit(item) || !inside || item.by === null || item.quiet) {
    return null;
  }

  if (doc.suggestions.get(item.by)?.status !== 'pending') {
    return null;
  }

  const runs = LOADED.get(doc)?.get(item.space);
  const kept = runs?.find((run) => offset >= run.from && offset < run.to);
  if (kept !== undefined) {
    return kept.typer;
  }

  // A change's units live in the space `i:<change id>.<k>`
  const change = item.space.startsWith('i:')
    ? item.space.slice(2, item.space.lastIndexOf('.'))
    : '';
  return parseOrderedChangeId(change)?.namespace ?? null;
}

/**
 * Visit each unit a pending insertion holds that the file writes inside its revision mark, in the
 * order a load of the file reads them, with its suggestion: units hidden for good, in a note left
 * out, or covered by their container's mark (`quiet`) are passed over.
 */
export function eachInsertedUnit(
  doc: OrderedDocument,
  visit: (suggestion: string, item: OrderedItem, offset: number) => void
): void {
  const read = new CardRead(doc);
  for (const story of storiesInFileOrder(doc)) {
    for (const item of story.tree.items()) {
      if (item.by === null || item.quiet || !isUnit(item) || !read.isUnit(item, story)) {
        continue;
      }

      for (let k = 0; k < item.length; k += 1) {
        visit(item.by, item, item.offset + k);
      }
    }
  }
}

/** Keep the typers a checkpoint names for the units of a document loaded from it, by space. */
export function keepLoadedTypers(
  doc: OrderedDocument,
  typers: ReadonlyMap<string, readonly TypedRun[]>
): void {
  LOADED.set(doc, typers);
}

/**
 * The stories in the order a load of the composed file makes them: the main document, the other
 * parts by name, then the text boxes by their objects' places.
 */
function storiesInFileOrder(doc: OrderedDocument): OrderedStory[] {
  const main = mainStory(doc);
  if (main === null) {
    return [];
  }

  const order = byStoryOrder(main.name);
  const parts = [...doc.stories.values()]
    .filter((s) => !isTextBoxStoryName(s.name))
    .sort((a, b) => order(a.name, b.name));
  const boxes = [...doc.stories.values()]
    .filter((s) => isTextBoxStoryName(s.name))
    .map((story) => {
      const host = textBoxHostOf(doc, story.name);
      const hostStory = host === null ? undefined : doc.stories.get(host.host);
      const atom = host === null || hostStory === undefined ? null : itemAt(hostStory, host.atom);
      const position =
        atom === null || hostStory === undefined ? 0 : hostStory.tree.before(atom).units;
      return { story, host: parts.indexOf(hostStory ?? main), position };
    });
  boxes.sort((a, b) => a.host - b.host || a.position - b.position);
  return [...parts, ...boxes.map(({ story }) => story)];
}
