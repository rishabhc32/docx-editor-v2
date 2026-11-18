// A rejected new link that still holds something keeps its element, as Word keeps it.

import { isSplitEdge, statusOf } from '../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { ComposeContext } from './context.ts';

const NONE: ReadonlySet<OrderedItem> = new Set();

/**
 * Whether a decision rejected the insertion of the item, in this drawing. A take-back (its
 * author's, or the organiser's undo of a direct edit) is no decision: the link never was, and it
 * goes.
 */
function rejectedInsertion(ctx: ComposeContext, item: OrderedItem): boolean {
  if (item.by === null || ctx.doc.suggestions.get(item.by)?.direct === true) return false;
  return ctx.rejectedOrDrawnSo(item.by);
}

/** Whether something other than its container's insertion hides the link edge. */
function hiddenOtherwise(ctx: ComposeContext, item: OrderedItem): boolean {
  if (item.gone || item.retracted || ctx.containerWent(item)) return true;
  return item.dels.some((id) => statusOf(ctx.doc, id) === 'accepted');
}

/** Whether a link edge is gone only because a decision rejected the link's insertion. */
function rejectedLinkEdge(ctx: ComposeContext, item: OrderedItem): boolean {
  if (item.token !== 'hyperlink' || item.pair === undefined || hiddenOtherwise(ctx, item))
    return false;
  if (!isSplitEdge(item)) return rejectedInsertion(ctx, item);

  // A split's edge: its Enter taken back joins the halves again; else the link decides
  if (item.by !== null && ctx.revertedDirect(item.by)) return false;
  const open = ctx.story.spaces
    .get(item.pair)
    ?.find((end) => end.kind === 'open' && !isSplitEdge(end));
  return open !== undefined && !hiddenOtherwise(ctx, open) && rejectedInsertion(ctx, open);
}

/** Whether an item is an edge (open or close item) of a link whose rejection alone hides it. */
function isRejectedEdge(ctx: ComposeContext, item: OrderedItem): boolean {
  const edge = item.kind === 'open' || item.kind === 'close';
  return edge && item.pair !== undefined && ctx.gone(item) && rejectedLinkEdge(ctx, item);
}

/**
 * The edges (open and close items) of the rejected new links in a paragraph that still hold
 * something shown: a marker, a unit or a container. Each half of a split link counts alone, as
 * each is its own `w:hyperlink` in the file.
 *
 * Word's Reject All takes a new link's runs out and keeps the `w:hyperlink` (with its
 * relationship) when a comment's start, a bookmark or another person's text is still in it; only
 * a link left with nothing goes (the engine does the same, `tree-op-revisions.ts`). So does a
 * reopened room file, whose link belongs to no suggestion.
 */
export function keptLinkEdges(
  ctx: ComposeContext,
  buffer: readonly OrderedItem[]
): ReadonlySet<OrderedItem> {
  let kept: Set<OrderedItem> | null = null;

  // The open edge of each candidate link the walk is in, by its pair, and whether it holds any
  const inside = new Map<string, { open: OrderedItem; holds: boolean }>();
  for (const item of buffer) {
    if (item.kind === 'mark') continue;
    if (item.pair !== undefined && isRejectedEdge(ctx, item)) {
      if (item.kind === 'open') {
        inside.set(item.pair, { open: item, holds: false });
        continue;
      }

      const link = inside.get(item.pair);
      inside.delete(item.pair);
      if (link?.holds) {
        kept ??= new Set();
        kept.add(link.open).add(item);
      }

      continue;
    }

    if (inside.size === 0 || ctx.gone(item)) continue;
    for (const link of inside.values()) link.holds = true;
  }

  return kept ?? NONE;
}

/**
 * Whether a hidden link's open item is a rejected new link the drawing keeps: edits find it as
 * they find any link that shows, as a reopened file has it (a new address, an unlink).
 */
export function isKeptLink(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): boolean {
  if (open.kind !== 'open' || open.token !== 'hyperlink' || open.pair === undefined) return false;
  const ctx = new ComposeContext(doc, story, {});
  if (!ctx.gone(open) || !rejectedLinkEdge(ctx, open)) return false;

  // As `keptLinkEdges` walks the paragraph: up to this half's end, for something shown
  for (let item = story.tree.next(open); item; item = story.tree.next(item)) {
    const gone = ctx.gone(item);
    if (item.kind === 'mark') {
      if (!gone) return false;
      continue;
    }

    if (item.pair === open.pair && isRejectedEdge(ctx, item)) return false;
    if (!gone) return true;
  }

  return false;
}
