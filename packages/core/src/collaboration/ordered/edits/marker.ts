// Markers with no width that edits add: a comment's range and reference, a bookmark, a move's
// range. A marker is no unit: it stays where it was put.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedIdMinter } from '../change/ids.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { createOrderedItem, itemAt, type OrderedItemInit } from '../model/document.ts';
import { moveRangeNodes, whilePending, type MoveHalf } from '../model/move-ranges.ts';
import { splitItem } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem } from '../model/types.ts';
import { wmlAttribute } from '../ooxml-nodes.ts';
import type { Ctx } from './context.ts';
import type { OrderedResolvedPlace } from './places.ts';

/** An element `w:<name>` with `w:` attributes and ids from the change's id minter. */
export function wml(
  ids: OrderedIdMinter,
  kind: OoxmlElement['kind'],
  name: string,
  attributes: Readonly<Record<string, string>>,
  children: readonly OoxmlElement[] = []
): OoxmlElement {
  return build(
    ids.node(),
    kind,
    name,
    Object.entries(attributes).map(([localName, value]) => wmlAttribute(localName, value)),
    children
  );
}

/** A Word decimal from the change's minter that `used` does not hold: the same on every copy. */
export function freshDecimal(ctx: Pick<Ctx, 'ids'>, used: (id: string) => boolean): string {
  for (;;) {
    const id = String(ctx.ids.decimal());
    if (!used(id)) {
      return id;
    }
  }
}

/**
 * The item a marker goes right before: the shown unit at the place (split out of its text
 * item), or its paragraph's mark. Read by address, so an earlier split of the same item is seen.
 */
export function itemBefore(ctx: Ctx, place: OrderedResolvedPlace): OrderedItem {
  const char = place.paragraph.chars[place.offset];
  if (!char) {
    return place.paragraph.mark;
  }

  const item = itemAt(ctx.story, char.address);
  if (!item) {
    throw new OrderedRefusal('unknown-address', char.address.space);
  }

  if (char.address.offset === item.offset) {
    return item;
  }

  return splitItem(ctx.doc, ctx.story, item, char.address.offset)?.[1] ?? item;
}

/** A new marker in a paragraph: the rows and cells around it hold it too (a rejected row). */
export function newMarkerItem(
  ctx: Ctx,
  p: OrderedViewParagraph,
  node: OoxmlNode,
  extra: Partial<OrderedItemInit> = {}
): OrderedItem {
  return createOrderedItem({
    space: ctx.ids.marker(),
    kind: 'zero',
    node,
    within: p.mark.within,
    ...extra,
  });
}

/** A move half's range start and end items in a paragraph, around the moved words. */
export function moveRangeItems(
  ctx: Ctx,
  p: OrderedViewParagraph,
  half: MoveHalf,
  name: string,
  suggestion: string
): readonly [OrderedItem, OrderedItem] {
  const ids = [ctx.ids.node(), ctx.ids.node()] as const;
  const [start, end] = moveRangeNodes(ctx.doc, half, name, suggestion, ids);
  const labels = { moveRange: true, ...whilePending(suggestion) };
  return [newMarkerItem(ctx, p, start, labels), newMarkerItem(ctx, p, end, labels)];
}
