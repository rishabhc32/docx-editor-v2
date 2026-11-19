// Blocks the engine builds on a scratch part (rows, tables, notes, pasted content), loaded as
// items of the edit's suggestion through the same loader as a file, so they have the labels a
// save and a reopen would give them.

import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { composeBlocks } from '../compose/blocks.ts';
import { withStoryBlocks } from '../compose/paragraph-ops.ts';
import { loadBlocks } from '../load/blocks.ts';
import { loadInEdit, newLoadBudget, newLoadState, RevisionReader } from '../load/state.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem, OrderedStory } from '../model/types.ts';
import { attributeOf, childNamed } from '../ooxml-nodes.ts';
import { scratchCompose, withNewIds, type Ctx } from './context.ts';
import { withParagraphIds } from './para-ids.ts';

/** The attribution the engine's template ops run under: every revision becomes the edit's. */
export const TEMPLATE_REVISION = { author: 'template', date: '2026-01-01T00:00:00Z' } as const;

/**
 * The revision for an engine op run under the edit's own author, so the engine knows the
 * author's own records and cells.
 */
export const revisionOf = (ctx: Ctx) => ({
  author: ctx.attribution.author,
  date: TEMPLATE_REVISION.date,
});

/** The story as the view draws it, every pending suggestion shown. */
export function viewPart(ctx: Ctx): OoxmlPart {
  const compose = scratchCompose(ctx.doc, ctx.story, 'view');
  return withStoryBlocks(ctx.story.part, composeBlocks(compose).children);
}

/** A node of the view by id, as an element. */
export function viewNode(view: OoxmlPart, id: string): OoxmlElement {
  const node = findNode(view, id);
  if (!node || node.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', `${id} is not in the view`);
  return node;
}

/** The blocks of a scratch part: its body's children, or its root's for a story with no body. */
export function scratchBlocks(part: OoxmlPart): readonly OoxmlNode[] {
  const body = part.root.localName === 'document' ? childNamed(part.root, 'body') : part.root;
  return body?.children ?? [];
}

/** The id of an open item's container node. */
export function shellId(open: OrderedItem): string {
  const node = open.node;
  if (!node || node.kind === 'textValue')
    throw new OrderedRefusal('malformed-story', `${open.space} has no container node`);
  return node.id;
}

/**
 * The engine's op on a scratch part that holds only `blocks`. The result is the scratch part;
 * a refusal of the engine is a named refusal of the edit.
 */
export function engineScratch(
  view: OoxmlPart,
  blocks: readonly OoxmlNode[],
  op: TreeDocOp
): OoxmlPart {
  const result = applyTreeOp(withStoryBlocks(view, blocks), op, { deferValidation: true });
  if (!result.ok) throw engineRefusal(op.op, result.reason);
  return result.part;
}

/**
 * Items for a block node the engine built, with ids from the edit's id source. Every revision
 * in it is `suggestion`, or as the reader given reads it; the loader gives the labels a reopened
 * file would have. `keep` selects the items the edit uses.
 */
export function templateItems(
  ctx: Ctx,
  node: OoxmlElement,
  suggestion: string | RevisionReader,
  keep: (item: OrderedItem) => boolean = () => true
): OrderedItem[] {
  const template = withParagraphIds(ctx, withNewIds(ctx, node));
  const budget = newLoadBudget();
  const ranged = new Set(ctx.doc.movesWithRanges);
  const state = newLoadState(
    ctx.doc,
    ctx.story.name,
    typeof suggestion === 'string' ? new RevisionReader(ctx.doc, suggestion) : suggestion,
    budget
  );

  // What the loader adds for compose (run shells, moves with range markers) the edit takes back
  // with its record
  loadInEdit(ctx.doc, () => loadBlocks(state, [template], 1));
  const items = withoutCopiedMoves(
    ctx.story,
    ranged,
    state.items
      .map((item, k) => (state.prompts.has(k) ? { ...item, prompt: true } : item))
      .filter(keep)
  );

  // A move whose range markers the template holds has them as items only when the edit keeps
  // them
  const markers = new Set(items.flatMap(rangeName));
  for (const name of [...ctx.doc.movesWithRanges])
    if (!ranged.has(name) && !markers.has(name)) ctx.doc.movesWithRanges.delete(name);
  return items;
}

/**
 * A template's words of a move the document holds (the view drew a party's pending move, and
 * the engine copied it) are a copy, no part of the move: plain words, without the move's range
 * markers, which stay the move's own. Only a move of the template itself (pasted words that hold
 * one) keeps them (seed 1014: a copy that kept them hid the move's own markers on save).
 */
function withoutCopiedMoves(
  story: OrderedStory,
  ranged: ReadonlySet<string>,
  items: OrderedItem[]
): OrderedItem[] {
  const named = new Set(
    items.flatMap((item) => [...(item.move === null ? [] : [item.move]), ...rangeName(item)])
  );
  if (named.size === 0) return items;
  const copied = new Set([...named].filter((name) => ranged.has(name)));
  for (const item of story.tree.items())
    if (item.move !== null && named.has(item.move)) copied.add(item.move);
  if (copied.size === 0) return items;

  // A range's end has no name: its `w:id` is its start's
  const ranges = new Set(
    items.flatMap((item) =>
      rangeName(item).some((name) => copied.has(name)) ? rangeAttribute(item, 'id') : []
    )
  );
  return items.flatMap((item) => {
    if (rangeAttribute(item, 'id').some((id) => ranges.has(id))) return [];
    return item.move !== null && copied.has(item.move) ? [{ ...item, move: null }] : [item];
  });
}

/** An attribute of a move range marker, when the item is one and has it. */
function rangeAttribute(item: OrderedItem, localName: string): string[] {
  if (!item.moveRange) return [];
  const node = item.node;
  const value = node && node.kind !== 'textValue' ? attributeOf(node, localName) : undefined;
  return value === undefined ? [] : [value];
}

/** A move range marker's name, when the item is one and has one. */
const rangeName = (item: OrderedItem): string[] => rangeAttribute(item, 'name');
