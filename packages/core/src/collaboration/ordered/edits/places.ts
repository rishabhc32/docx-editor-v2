// Places by address. An edit names a place as the unit it comes before; a paragraph mark's
// address is the end of its paragraph. Addresses never change during a room session, so one
// place means the same on every copy, whatever other edits came first; the resolver finds where
// the place is in the view now.

import { findNode } from '../../../store/package/ooxml-edit.ts';
import { isBlockToken } from '../compose/blocks.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { isKeptLink } from '../compose/kept-links.ts';
import { isNotesStory, noteKeyOf, notesWithHiddenReferences } from '../compose/note-references.ts';
import { isHidden, itemAt, missingAddress, storyNamed } from '../model/document.ts';
import { promptMarksOf } from '../model/prompt-marks.ts';
import { addressLabel, OrderedRefusal } from '../model/refusal.ts';
import { textBoxGone } from '../model/text-boxes.ts';
import {
  loadedSpace,
  type OrderedContainerToken,
  type OrderedDocument,
  type OrderedItem,
  type OrderedItemAddress,
  type OrderedItemKind,
  type OrderedStory,
} from '../model/types.ts';

/**
 * A place in a story: right before the unit at this address. The address of a paragraph mark
 * is the end of its paragraph.
 *
 * @public
 */
export type OrderedPlace = OrderedItemAddress;

/** Where a place is in the view now. @public */
export interface OrderedResolvedPlace {
  /** The shown paragraph's index in the story. */
  readonly index: number;
  /** The offset in the paragraph's view: its characters and objects. */
  readonly offset: number;
  readonly paragraph: OrderedViewParagraph;
}

/**
 * The shown paragraph mark at or after an item. A place never leaves its note: in a note that no
 * longer shows (converted to the other kind, or a rejected new one) the next shown paragraph is
 * another note's, so the place's target is hidden.
 */
function shownMarkFrom(doc: OrderedDocument, story: OrderedStory, from: OrderedItem): OrderedItem {
  for (let item: OrderedItem | null = from; item; item = story.tree.next(item)) {
    if (item.kind === 'mark' && !isHidden(doc, item)) return item;
    if (item.kind === 'close' && item.token === 'note')
      throw new OrderedRefusal('hidden-target', 'the note does not show');
  }

  throw new OrderedRefusal('malformed-story', `no shown paragraph after ${from.space}`);
}

/**
 * Refuses an item of a story whose holder no longer shows: a text box whose object went (its
 * insertion rejected or taken back, its deletion accepted), or a note none of whose references
 * shows. Compose writes neither, so a change there would be a pending suggestion that nothing
 * shows. A struck text box or note still shows, as Word shows it, and takes changes.
 */
export function refuseHiddenHolder(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): void {
  if (textBoxGone(doc, story.name)) {
    throw new OrderedRefusal('hidden-target', 'the text box does not show');
  }

  if (!isNotesStory(story)) return;
  for (let back: OrderedItem | null = item; back; back = story.tree.prev(back)) {
    if (back.kind === 'close' && back.token === 'note' && back !== item) return;
    if (back.kind !== 'open' || back.token !== 'note') continue;
    const key = noteKeyOf(back);
    if (key !== undefined && notesWithHiddenReferences(doc).has(key))
      throw new OrderedRefusal('hidden-target', 'the note does not show');
    return;
  }
}

/** An address as a refusal names it. */

/**
 * The item that starts at an address, of a kind (and one of `tokens`), that still shows, in a
 * text box or note that still shows. A link's open item a reject kept shows while its words do.
 */
export function shownItemAt(
  doc: OrderedDocument,
  story: OrderedStory,
  address: OrderedItemAddress,
  kind: OrderedItemKind,
  tokens?: readonly OrderedContainerToken[]
): OrderedItem {
  const item = itemAt(story, address);
  if (!item) {
    throw new OrderedRefusal(missingAddress(doc, address), addressLabel(address));
  }

  const tokenOk = tokens === undefined || (item.token !== undefined && tokens.includes(item.token));
  if (item.kind !== kind || item.offset !== address.offset || !tokenOk) {
    throw new OrderedRefusal('unknown-address', addressLabel(address));
  }

  if (isHidden(doc, item) && !isKeptLink(doc, story, item)) {
    throw new OrderedRefusal('hidden-target', addressLabel(address));
  }

  refuseHiddenHolder(doc, story, item);
  return item;
}

/**
 * A place in a prompt's paragraph that content replaced (`change/settle-prompts.ts`), such as
 * a caret another person put in the prompt's other line at the same time: that paragraph is gone,
 * so the place goes to the start of the control's next shown paragraph, or else to the end of
 * its shown paragraph before, and never out of the control. Null for any other place.
 */
export function replacedPromptPlace(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): { readonly mark: OrderedItem; readonly end: boolean } | null {
  if (promptMarksOf(doc).size === 0) return null;
  let from: OrderedItem | null = item;
  while (from && from.kind !== 'mark') from = story.tree.next(from);
  if (!from || from.prompt !== true || !isHidden(doc, from)) return null;
  const next = shownMarkInside(doc, from, (at) => story.tree.next(at), 'close');
  if (next) return { mark: next, end: false };
  const previous = shownMarkInside(doc, item, (at) => story.tree.prev(at), 'open');
  return previous ? { mark: previous, end: true } : null;
}

/** The first shown mark from `from` one way, before the edge of the container `from` is in. */
function shownMarkInside(
  doc: OrderedDocument,
  from: OrderedItem,
  step: (at: OrderedItem) => OrderedItem | null,
  edge: 'open' | 'close'
): OrderedItem | null {
  let depth = 0;
  for (let at = step(from); at; at = step(at)) {
    if (isBlockToken(at)) {
      if (at.kind !== edge) depth += 1;
      else if (depth === 0) return null;
      else depth -= 1;
      continue;
    }

    if (at.kind === 'mark' && !isHidden(doc, at)) return at;
  }

  return null;
}

/** The view's offset right after the last unit of `item`, or -1. */
function endOffsetOf(paragraph: OrderedViewParagraph, item: OrderedItem): number {
  for (let k = paragraph.chars.length - 1; k >= 0; k -= 1)
    if (paragraph.chars[k]?.item === item) return k + 1;
  return -1;
}

/**
 * Refuses a place whose address the story does not hold, by the reason it is missing: a
 * paragraph inside content kept whole as one item (a text box, a shape) is not covered.
 */
function refuseMissing(doc: OrderedDocument, story: OrderedStory, place: OrderedPlace): never {
  const missing = missingAddress(doc, place);
  const loaded = loadedSpace(place.space);
  if (
    missing === 'unknown-address' &&
    loaded !== null &&
    (loaded.kind === 'b' || loaded.kind === 'm') &&
    findNode(story.part, loaded.node)
  ) {
    throw new OrderedRefusal('not-covered', 'a place inside content kept whole (a text box)');
  }

  throw new OrderedRefusal(missing, addressLabel(place));
}

/** The shown paragraph a mark ends, with its index in the story. */
function shownParagraphOf(
  doc: OrderedDocument,
  name: string,
  story: OrderedStory,
  mark: OrderedItem
): { readonly index: number; readonly paragraph: OrderedViewParagraph } {
  const index = story.tree.before(mark).marks;
  const paragraph = orderedViewParagraph(doc, name, index);
  if (!paragraph) throw new OrderedRefusal('malformed-story', `no paragraph ${index}`);
  return { index, paragraph };
}

/**
 * Where a place is in the view now. A unit that does not show places the caret before the next
 * unit that shows; the end of a paragraph that a join took is where its text ends, inside the
 * paragraph that took it.
 *
 * @public
 */
export function resolveOrderedPlace(
  doc: OrderedDocument,
  name: string,
  place: OrderedPlace
): OrderedResolvedPlace {
  const story = storyNamed(doc, name);
  if (typeof place.space !== 'string' || !Number.isSafeInteger(place.offset))
    throw new OrderedRefusal('unknown-address', 'a place needs a space and an offset');
  const item = itemAt(story, place);
  if (!item) {
    return refuseMissing(doc, story, place);
  }

  refuseHiddenHolder(doc, story, item);
  const replaced = replacedPromptPlace(doc, story, item);
  if (replaced) {
    const { index, paragraph } = shownParagraphOf(doc, name, story, replaced.mark);
    return { index, offset: replaced.end ? paragraph.chars.length : 0, paragraph };
  }

  const mark = shownMarkFrom(doc, story, item);
  const { index, paragraph } = shownParagraphOf(doc, name, story, mark);
  if (item === mark) return { index, offset: paragraph.chars.length, paragraph };
  if (item.kind === 'mark') {
    // A joined paragraph's end: after the last unit before its mark that shows
    for (let back = story.tree.prev(item); back; back = story.tree.prev(back)) {
      if (back.kind === 'mark' && !isHidden(doc, back)) break;
      const end = endOffsetOf(paragraph, back);
      if (end >= 0) return { index, offset: end, paragraph };
    }

    return { index, offset: 0, paragraph };
  }

  const at = paragraph.chars.findIndex(
    (c) => c.address.space === place.space && c.address.offset === place.offset
  );
  if (at >= 0) return { index, offset: at, paragraph };

  // The unit does not show: before the next unit that shows
  for (let next = story.tree.next(item); next && next !== mark; next = story.tree.next(next)) {
    const first = paragraph.chars.findIndex((c) => c.item === next);
    if (first >= 0) return { index, offset: first, paragraph };
  }

  return { index, offset: paragraph.chars.length, paragraph };
}

/**
 * The place at a view offset of a shown paragraph: the unit there, or the paragraph's mark at
 * its end. What a client sends for a caret it reads from the view.
 *
 * @public
 */
export function orderedPlaceAt(paragraph: OrderedViewParagraph, offset: number): OrderedPlace {
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > paragraph.chars.length)
    throw new OrderedRefusal('unknown-address', `offset ${offset} of ${paragraph.chars.length}`);
  const char = paragraph.chars[offset];
  return char ? char.address : { space: paragraph.mark.space, offset: paragraph.mark.offset };
}
