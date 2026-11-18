// Repeating section items (task 9.6): the engine's own op on the view adds an item that
// repeats the one before it (its controls back to their prompts) or removes one. A new item is
// loaded as the engine wrote it, under ids from the change; a removed item's items are struck.
// Both are direct edits, as the editor makes them in suggesting mode too: `revert` takes them
// back exactly.

import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { editedProperties } from '../../../store/store/tree-op-content-controls.ts';
import { isHidden } from '../model/document.ts';
import { insertItems } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem } from '../model/types.ts';
import { isWml, withAttributes, withChildren } from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import { engineOnView, strikeBlock } from './control-shell.ts';
import { suggestionFor, type Ctx } from './context.ts';
import { isParagraphIdentity } from './para-ids.ts';
import { shellId, templateItems } from './templates.ts';

const isContent = (node: OoxmlNode): boolean =>
  node.kind === 'contentControlContent' || isWml(node, 'sdtContent');

/** The section's items, by their open items, in order: its block controls one level in. */
function itemsOf(ctx: Ctx, section: OrderedItem): OrderedItem[] {
  const out: OrderedItem[] = [];
  let depth = 0;
  for (const item of itemsInside(ctx.story, section, ctx.doc)) {
    if (item.kind === 'open' && item.token !== undefined && item.token !== 'note') {
      if (depth === 0 && item.token === 'sdtBlock' && !isHidden(ctx.doc, item)) out.push(item);
      depth += 1;
    } else if (item.kind === 'close' && item.token !== undefined && item.token !== 'note')
      depth -= 1;
  }

  return out;
}

/**
 * A new item's element with ids from the change: no `w14:paraId` (the change's minter gives each
 * paragraph its own), and every control's `w:id` minted, so every copy writes the same.
 */
function withChangeIds(ctx: Ctx, node: OoxmlNode, depth = 0): OoxmlNode {
  if (node.kind === 'textValue' || depth > 64) return node;
  const children = node.children.map((child) => withChangeIds(ctx, child, depth + 1));
  if (isWml(node, 'sdtPr'))
    return editedProperties(withChildren(node, children), { id: ctx.ids.decimal() }, () =>
      ctx.ids.node()
    );
  const attributes = node.attributes.filter((a) => !isParagraphIdentity(a));
  return withAttributes(withChildren(node, children), attributes);
}

/** Add an item after the item before `index` (the first item's place for 0), as the engine does. */
export function addRepeatingItem(
  ctx: Ctx,
  section: OrderedItem,
  index: number | undefined
): string {
  if (!ctx.direct)
    throw new OrderedRefusal('invalid-edit', 'a repeating section item: a direct edit only');
  const id = shellId(section);
  const items = itemsOf(ctx, section);
  const { result } = engineOnView(ctx, {
    op: 'addRepeatingSectionItem',
    controlId: id,
    ...(index === undefined ? {} : { index }),
  });
  const at = index ?? items.length;
  const after = findNode(result, id);
  const content = after && after.kind !== 'textValue' ? after.children.find(isContent) : undefined;
  const element = content && content.kind !== 'textValue' ? content.children[at] : undefined;
  if (!element || element.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', 'addRepeatingSectionItem: no new item');
  const suggestion = suggestionFor(ctx);

  // Everything in the new item goes with it through `within`, as a new row's content does: a
  // party's words typed in it later too, when the organiser takes the item back
  const within = { by: [...section.within.by, suggestion.id], dels: section.within.dels };
  const loaded = templateItems(ctx, withChangeIds(ctx, element) as OoxmlElement, suggestion.id);
  const last = loaded.length - 1;
  const labelled = loaded.map(
    (item, k): OrderedItem => ({
      ...item,
      by: item.by ?? suggestion.id,
      within: k === 0 || k === last ? section.within : within,
    })
  );
  const before = items[at - 1];
  const anchor =
    before !== undefined
      ? closeOf(ctx.story, before, ctx.doc)
      : ctx.story.tree.prev(items[0] ?? closeOf(ctx.story, section, ctx.doc));
  insertItems(ctx.doc, ctx.story, anchor, labelled);
  return suggestion.id;
}

/** Remove the item at `index`: all its items struck by the direct edit. */
export function removeRepeatingItem(ctx: Ctx, section: OrderedItem, index: number): string {
  if (!ctx.direct)
    throw new OrderedRefusal('invalid-edit', 'a repeating section item: a direct edit only');
  engineOnView(ctx, { op: 'removeRepeatingSectionItem', controlId: shellId(section), index });
  const item = itemsOf(ctx, section)[index];
  if (!item) throw new OrderedRefusal('unknown-control', `no item ${index}`);
  const suggestion = suggestionFor(ctx);
  strikeBlock(ctx, item, suggestion.id);
  return suggestion.id;
}
