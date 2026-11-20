// The complex fields open at a place of a paragraph, read from their markers (`w:fldChar`) as the
// engine's field parse reads them. A whole field the room holds is one item with its pieces, and
// opens nothing.

import type { Ctx } from './edits/context.ts';
import { isHidden } from './model/document.ts';
import { OrderedRefusal } from './model/refusal.ts';
import type { OrderedItem } from './model/types.ts';
import { attributeOf } from './ooxml-nodes.ts';

/** A field character's type (`w:fldChar/@w:fldCharType`) on a marker, or undefined. */
export function fieldCharType(item: OrderedItem): string | undefined {
  const node = item.node;
  if (item.kind !== 'zero' || node === null || node.kind === 'textValue') {
    return undefined;
  }

  return node.localName === 'fldChar' ? attributeOf(node, 'fldCharType') : undefined;
}

/** Fields open after a marker: one more at a begin, one fewer at an end. */
export function fieldDepthAfter(depth: number, item: OrderedItem): number {
  const type = fieldCharType(item);
  if (type === 'begin') {
    return depth + 1;
  }

  if (type === 'end') {
    return Math.max(0, depth - 1);
  }

  return depth;
}

/**
 * Refuse a whole complex field among `items` after a field that begins in its paragraph and has
 * not ended there. The engine reads no field inside another field's result in that paragraph (its
 * parse stops at a begin with no end there, and a closed field's result is that field's), so a
 * reopen or a fold would hold no unit and no card for it. A simple field is one unit anywhere.
 */
export function refuseFieldInField(ctx: Ctx, items: readonly OrderedItem[]): void {
  for (const item of items) {
    if (item.kind !== 'atom' || item.pieces === undefined || item.bare) {
      continue;
    }

    const before: OrderedItem[] = [];
    for (
      let at = ctx.story.tree.prev(item);
      at && at.kind !== 'mark';
      at = ctx.story.tree.prev(at)
    ) {
      before.push(at);
    }

    let depth = 0;
    for (const at of before.reverse()) {
      if (!isHidden(ctx.doc, at)) {
        depth = fieldDepthAfter(depth, at);
      }
    }

    if (depth > 0) {
      throw new OrderedRefusal('invalid-edit', 'a field inside another field');
    }
  }
}
