// The list's paragraphs by the editor's node ids.

import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { isHidden, itemAt } from '../model/document.ts';
import {
  markSpace,
  tokenSpace,
  type OrderedDocument,
  type OrderedItem,
  type OrderedItemAddress,
  type OrderedStory,
} from '../model/types.ts';
import type { OrderedChangeAddress, OrderedChangePlace } from '../change/types.ts';
import { OrderedClientRefusal } from './refusal.ts';

/** A shown paragraph of the list, found by its node id. */
export interface ListParagraph {
  readonly story: OrderedStory;
  readonly mark: OrderedItem;
  /** Its index among the story's shown paragraphs. */
  readonly index: number;
  readonly view: OrderedViewParagraph;
}

/** The view's text, one character for each unit, as the editor's paragraph text reads. */
export const viewText = (view: OrderedViewParagraph): string =>
  view.chars.map((char) => char.ch).join('');

/**
 * The mark item of the paragraph with this node id in a story, shown or not, or null. After
 * every publish the editor shows the list's composed paragraphs with the list's own node ids, so
 * an op's `paragraphId` names a paragraph of the list: a loaded paragraph's mark lives in space
 * `m:<file>:<node id>`, and a paragraph a change made has its mark's space as its node id
 * (`k:<change>.<k>`).
 */
export function markOfNode(
  doc: OrderedDocument,
  story: OrderedStory,
  nodeId: string
): OrderedItem | null {
  for (const space of [markSpace(doc.file, nodeId), nodeId]) {
    const item = itemAt(story, { space, offset: 0 });
    if (item && item.kind === 'mark' && item.node?.id === nodeId) return item;
  }

  return null;
}

/** The open item of the container with this node id in a story, or null. */
export function openOfNode(
  doc: OrderedDocument,
  story: OrderedStory,
  nodeId: string
): OrderedItem | null {
  for (const space of [tokenSpace(doc.file, nodeId), nodeId]) {
    const item = itemAt(story, { space, offset: 0 });
    if (item && item.kind === 'open' && item.offset === 0) return item;
  }

  return null;
}

/**
 * The first shown mark at or after an item: the mark that ends the paragraph the item shows in.
 * Null at the story's end, or when `stop` holds for an item before such a mark.
 */
export function nextShownMark(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  stop?: (item: OrderedItem) => boolean
): OrderedItem | null {
  for (let at: OrderedItem | null = item; at; at = story.tree.next(at)) {
    if (at.kind === 'mark' && !isHidden(doc, at)) {
      return at;
    }

    if (stop?.(at)) {
      return null;
    }
  }

  return null;
}

/** The shown paragraph a mark ends, or null when the mark is hidden. */
export function paragraphOfMark(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem
): ListParagraph | null {
  if (isHidden(doc, mark)) return null;
  const index = story.tree.before(mark).marks;
  const view = orderedViewParagraph(doc, story.name, index);
  if (!view || view.mark !== mark) return null;
  return { story, mark, index, view };
}

/**
 * The shown paragraph with this node id, in the story first, then in any story; null when the
 * list shows none.
 */
export function shownParagraphOfNode(
  doc: OrderedDocument,
  nodeId: string,
  storyName?: string
): ListParagraph | null {
  const first = storyName === undefined ? undefined : doc.stories.get(storyName);
  const stories = first ? [first, ...doc.stories.values()] : [...doc.stories.values()];
  for (const story of stories) {
    const mark = markOfNode(doc, story, nodeId);
    const found = mark ? paragraphOfMark(doc, story, mark) : null;
    if (found) {
      return found;
    }
  }

  return null;
}

/**
 * The shown paragraph with this node id, as `shownParagraphOfNode` finds it. Refused by name
 * when the list has no such paragraph: the editor shows one the list never published.
 */
export function paragraphOfNode(
  doc: OrderedDocument,
  nodeId: string,
  storyName?: string
): ListParagraph {
  const found = shownParagraphOfNode(doc, nodeId, storyName);
  if (!found) {
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', `no paragraph ${nodeId}`);
  }

  return found;
}

/**
 * The shown paragraph that holds the paragraph with this node id now: itself, or the one a join
 * took it into (its mark hidden, its text before the next shown mark). The engine keeps the
 * first paragraph's node id across a join, so a later op of the same batch names it by that id.
 */
export function joinedParagraphOfNode(doc: OrderedDocument, nodeId: string): ListParagraph {
  for (const story of doc.stories.values()) {
    const mark = markOfNode(doc, story, nodeId);
    const shown = mark ? nextShownMark(doc, story, mark) : null;
    const found = shown ? paragraphOfMark(doc, story, shown) : null;
    if (found) {
      return found;
    }
  }

  throw new OrderedClientRefusal('ordered-refused', 'out-of-step', `no paragraph ${nodeId}`);
}

/** The shown paragraph at an index of a story, or a refusal. */
export function paragraphAtIndex(
  doc: OrderedDocument,
  story: OrderedStory,
  index: number
): ListParagraph {
  const view = orderedViewParagraph(doc, story.name, index);
  if (!view)
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', `no paragraph ${index}`);
  return { story, mark: view.mark, index, view };
}

/** A unit's or a mark's address as a change names it. */
export const addressOf = (item: OrderedItem): OrderedChangeAddress => ({
  space: item.space,
  offset: item.offset,
});

/**
 * The place at an offset of a paragraph's view, as a change names it: the unit there, or the
 * paragraph's end (`{ mark }`) past its last unit. An offset outside the view is refused.
 */
export function placeAt(paragraph: ListParagraph, offset: number): OrderedChangePlace {
  const { view } = paragraph;
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > view.chars.length)
    throw new OrderedClientRefusal(
      'ordered-refused',
      'out-of-step',
      `offset ${offset} of ${view.chars.length}`
    );
  if (offset === view.chars.length) return { mark: addressOf(view.mark) };
  const place: OrderedItemAddress = orderedPlaceAt(view, offset);
  return { space: place.space, offset: place.offset };
}
