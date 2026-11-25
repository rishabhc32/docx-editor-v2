// Content the engine keeps from edits: a table of contents, a control bound to custom XML, and a
// content-locked control. An edit that reaches it is refused by name, as the engine refuses it.

import type { OoxmlElement } from '../../store/package/ooxml-tree.ts';
import { formattingPassesContentLock } from '../../store/store/content-lock-policy.ts';
import {
  declaredLockOf,
  lockFlagsFromValue,
  type ContentControlLockFlags,
} from '../../store/store/tree-op-nodes.ts';
import type { OoxmlProperty } from '../../store/store/tree-op-types.ts';
import { orderedViewParagraph } from './compose/compose.ts';
import {
  ComposeContext,
  type OrderedViewChar,
  type OrderedViewParagraph,
} from './compose/context.ts';
import { containerPropertiesAfter } from './compose/properties.ts';
import { isControlSegment, openOfSegment, segmentsOf } from './containers.ts';
import type { Ctx } from './edits/context.ts';
import { isHidden } from './model/document.ts';
import { OrderedRefusal } from './model/refusal.ts';
import type { OrderedItem } from './model/types.ts';
import { attributeOf, childNamed, isWml, withChildren } from './ooxml-nodes.ts';
import { reachesTocField } from './toc-field-range.ts';

/** A control's properties (`w:sdtPr`), or undefined. */
function propertiesOf(open: OrderedItem): OoxmlElement | undefined {
  const node = open.node;
  return node && node.kind !== 'textValue' ? childNamed(node, 'sdtPr') : undefined;
}

/** A control whose content comes from custom XML (`w:dataBinding`). */
function isBound(open: OrderedItem): boolean {
  return propertiesOf(open)?.children.some((c) => isWml(c, 'dataBinding')) === true;
}

/** A block control that holds a table of contents (Word's gallery of that name). */
export function isTableOfContents(open: OrderedItem): boolean {
  const properties = propertiesOf(open);
  const docPartObj = properties === undefined ? undefined : childNamed(properties, 'docPartObj');
  const gallery = docPartObj === undefined ? undefined : childNamed(docPartObj, 'docPartGallery');
  return gallery !== undefined && attributeOf(gallery, 'val') === 'Table of Contents';
}

/**
 * A control's own lock (`w:lock`) as it shows now: the file's, or the organiser's later change
 * of it (`control-property`), a record on the open item that compose writes.
 */
function lockOf(ctx: Ctx, open: OrderedItem): ContentControlLockFlags {
  const node = open.node;
  if ((open.token !== 'sdt' && open.token !== 'sdtBlock') || !node || node.kind === 'textValue') {
    return lockFlagsFromValue(undefined);
  }

  if (open.props.length === 0) {
    return declaredLockOf(node);
  }

  const compose = new ComposeContext(ctx.doc, ctx.story, {});
  return declaredLockOf(withChildren(node, containerPropertiesAfter(compose, open, node.children)));
}

/** A control whose lock keeps its content from edits (`contentLocked`, `sdtContentLocked`). */
const contentLocked = (ctx: Ctx) => (open: OrderedItem) => lockOf(ctx, open).content;

/** The inline controls a view character sits in that `kept` names, by their segments. */
function keptControls(char: OrderedViewChar | undefined, kept: (segment: string) => boolean) {
  if (!char) {
    return [];
  }

  return segmentsOf(char.containers).filter(
    (segment) => isControlSegment(segment) && kept(segment)
  );
}

/**
 * Whether an edit of view characters `[start, end)` (a caret when they are equal) reaches inside
 * an inline control `kept` names: a character in it, or a caret with it on both sides. At its
 * edges a caret is outside it (placement rules 1 and 5).
 */
function reachesInside(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  kept: (open: OrderedItem) => boolean
): boolean {
  // Each control once, however many characters of a long range sit in it
  const known = new Map<string, boolean>();
  const isKept = (segment: string): boolean => {
    let answer = known.get(segment);
    if (answer === undefined) {
      const open = openOfSegment(ctx.story, segment);
      answer = open !== null && kept(open);
      known.set(segment, answer);
    }

    return answer;
  };
  if (end > start) {
    return p.chars.slice(start, end).some((char) => keptControls(char, isKept).length > 0);
  }

  const before = keptControls(p.chars[start - 1], isKept);
  return before.length > 0 && keptControls(p.chars[start], isKept).some((c) => before.includes(c));
}

const LOCKED = 'the content of a content-locked control';

/**
 * Refuse a change inside block containers (a paragraph's, or the ones around a table) when one
 * of them is a control whose lock keeps its content (`locked`).
 */
export function refuseLockedBlocks(ctx: Ctx, blocks: readonly OrderedItem[]): void {
  if (blocks.some(contentLocked(ctx))) {
    throw new OrderedRefusal('locked', LOCKED);
  }
}

/**
 * Refuse a change of `items` (a row's, a table's, a paragraph's) when a control among them that
 * shows keeps its content, or, when the change removes them (`removes`), when a control among
 * them may not be removed (`sdtLocked`, `sdtContentLocked`).
 */
export function refuseLockedWithin(
  ctx: Ctx,
  items: readonly OrderedItem[],
  removes: boolean
): void {
  for (const item of items) {
    if (item.kind !== 'open' || isHidden(ctx.doc, item)) {
      continue;
    }

    const lock = lockOf(ctx, item);
    if (lock.content) {
      throw new OrderedRefusal('locked', LOCKED);
    }

    if (removes && lock.wrapper) {
      throw new OrderedRefusal('locked', 'a control that may not be removed');
    }
  }
}

/** The shown paragraph that holds an item (a mark, a unit, a container's edge), as drawn now. */
export function paragraphOfItem(ctx: Ctx, item: OrderedItem): OrderedViewParagraph | null {
  const p = orderedViewParagraph(ctx.doc, ctx.story.name, ctx.story.tree.before(item).marks);
  return p && (p.mark === item || p.items.includes(item)) ? p : null;
}

/**
 * Refuse a change of one item when a control around it keeps its content: a paragraph's mark (a
 * section break), an object (a picture's size, an equation), a note reference (a conversion), a
 * link's open item (its target). For a container's open item only the controls around it count,
 * as the engine asks a node's enclosing controls.
 */
export function refuseLockedAt(ctx: Ctx, item: OrderedItem): void {
  const p = paragraphOfItem(ctx, item);
  if (!p) {
    return;
  }

  refuseLockedBlocks(ctx, p.blocks);
  if (item.kind === 'mark') {
    return;
  }

  const own = `${item.token === 'hyperlink' ? 'h' : 'c'}:${item.space}`;
  const char = p.chars.find((c) =>
    item.kind === 'open' ? segmentsOf(c.containers).includes(own) : c.item === item
  );
  if (!char) {
    return;
  }

  const segments = segmentsOf(char.containers);
  const around = item.kind === 'open' ? segments.slice(0, segments.indexOf(own)) : segments;
  const locked = contentLocked(ctx);
  for (const segment of around) {
    const open = isControlSegment(segment) ? openOfSegment(ctx.story, segment) : null;
    if (open && locked(open)) {
      throw new OrderedRefusal('locked', LOCKED);
    }
  }
}

/**
 * Refuse a change of the table a paragraph is in (a row, a column, a cell's properties) when a
 * control around the table keeps its content. The engine's own op on the table meets the
 * controls in it.
 */
export function refuseLockedTable(ctx: Ctx, p: OrderedViewParagraph): void {
  const table = p.blocks.map((block) => block.token).lastIndexOf('tbl');
  refuseLockedBlocks(ctx, table < 0 ? p.blocks : p.blocks.slice(0, table));
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) whose content a
 * control's lock keeps (ECMA-376 Part 1 §17.5.2.23, the engine's `locked`): the paragraph is in a
 * `contentLocked` or `sdtContentLocked` block control, or the edit reaches inside such an inline
 * control. A lock holds for the controls inside the control too (`resolveContentControlLock`).
 * At an inline control's edge, typed text goes beside it (`promptSite` never enters it).
 */
export function refuseContentLock(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number
): void {
  refuseLockedBlocks(ctx, p.blocks);
  if (reachesInside(ctx, p, start, end, contentLocked(ctx))) {
    throw new OrderedRefusal('locked', LOCKED);
  }
}

/**
 * Whether a format of `properties` passes a content lock: a direct format that the room's policy
 * lets past it, as the engine lets an untracked restyle (`restylePassesContentLock`). A suggested
 * format stays refused, as the engine refuses a tracked restyle there.
 */
export function formatPassesContentLock(ctx: Ctx, properties: readonly OoxmlProperty[]): boolean {
  return ctx.direct && formattingPassesContentLock(properties, ctx.contentLockPolicy);
}

/** Whether typing never enters a control (the engine's `refusesTypedContent`): bound or locked. */
export function refusesTyping(ctx: Ctx, open: OrderedItem): boolean {
  return isBound(open) || lockOf(ctx, open).content;
}

/**
 * Refuse an edit of view characters `[start, end)` (a caret when they are equal) that reaches
 * content the engine keeps from edits: a table of contents, a bound control's content, a locked
 * control's content. A format names its `properties`: a locked control's content may then take
 * it (`formatPassesContentLock`).
 */
export function refuseLocked(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  properties?: readonly OoxmlProperty[]
): void {
  const toc = p.blocks.find(isTableOfContents);
  if (toc && reachesTocField(ctx, toc, p, start, end)) {
    throw new OrderedRefusal('not-covered', 'a table of contents result is generated');
  }

  if (reachesInside(ctx, p, start, end, isBound)) {
    throw new OrderedRefusal('not-covered', 'a control bound to custom XML');
  }

  if (properties === undefined || !formatPassesContentLock(ctx, properties)) {
    refuseContentLock(ctx, p, start, end);
  }
}
