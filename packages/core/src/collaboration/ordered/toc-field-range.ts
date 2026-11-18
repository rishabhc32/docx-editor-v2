// What the TOC field of a table of contents control covers in one paragraph, as the engine's
// read-only ranges give it. Only that range is generated: the control's heading, before the
// field, is ordinary text.

import type { OrderedViewParagraph } from './compose/context.ts';
import { closesContainer } from './containers.ts';
import type { Ctx } from './edits/context.ts';
import { isHidden } from './model/document.ts';
import type { OrderedItem } from './model/types.ts';
import { attributeOf } from './ooxml-nodes.ts';

/** The view characters a TOC field covers in one paragraph, and whether it starts or ends there. */
interface TocFieldRange {
  readonly start: number;
  readonly end: number;
  readonly startsHere: boolean;
  readonly endsHere: boolean;
}

/** A field character's type (`w:fldChar/@w:fldCharType`) on a marker, or undefined. */
function fieldCharType(item: OrderedItem): string | undefined {
  const node = item.node;
  if (item.kind !== 'zero' || node === null || node.kind === 'textValue') return undefined;
  return node.localName === 'fldChar' ? attributeOf(node, 'fldCharType') : undefined;
}

/** Fields open after a marker: one more at a begin, one fewer at an end. */
function fieldDepthAfter(depth: number, item: OrderedItem): number {
  const type = fieldCharType(item);
  if (type === 'begin') {
    return depth + 1;
  }

  if (type === 'end') {
    return Math.max(0, depth - 1);
  }

  return depth;
}

/** How many view characters each item of a paragraph draws. */
function charWidths(p: OrderedViewParagraph): Map<OrderedItem, number> {
  const widths = new Map<OrderedItem, number>();
  for (const char of p.chars) {
    widths.set(char.item, (widths.get(char.item) ?? 0) + 1);
  }

  return widths;
}

/**
 * The view characters of a paragraph in a table of contents control that its TOC field covers,
 * as the engine's read-only ranges (`tocParagraphRanges`) give them: from the field's begin to
 * its end, `startsHere` and `endsHere` where the paragraph holds them; null when the field
 * covers none of the paragraph (the control's heading, a paragraph after the field). A control
 * with no field is generated whole (`fieldlessTocRanges`).
 */
function tocFieldRange(ctx: Ctx, toc: OrderedItem, p: OrderedViewParagraph): TocFieldRange | null {
  const own = new Set(p.items);
  let depth = 0;
  let fields = 0;
  let range: TocFieldRange | null = null;
  let offset = 0;
  const widths = charWidths(p);
  for (let item = ctx.story.tree.next(toc); item; item = ctx.story.tree.next(item)) {
    if (closesContainer(item, toc, ctx.doc) || item === p.mark) break;
    const inside = own.has(item);
    if (inside && depth > 0 && range === null)
      range = { start: offset, end: p.chars.length, startsHere: false, endsHere: false };
    if (!isHidden(ctx.doc, item)) {
      const next = fieldDepthAfter(depth, item);
      if (next > depth) fields += 1;
      if (inside && depth === 0 && next > 0)
        range = { start: offset, end: p.chars.length, startsHere: true, endsHere: false };
      if (inside && depth > 0 && next === 0 && range)
        range = { ...range, end: offset, endsHere: true };
      depth = next;
    }

    if (inside) offset += widths.get(item) ?? 0;
  }

  const whole: TocFieldRange = {
    start: 0,
    end: p.chars.length,
    startsHere: false,
    endsHere: false,
  };
  if (range === null && depth > 0) return whole;
  if (fields === 0 && !fieldBeginsAfter(ctx, toc, p.mark)) return whole;
  return range;
}

/** Whether a field begins in the control after `from`. */
function fieldBeginsAfter(ctx: Ctx, toc: OrderedItem, from: OrderedItem): boolean {
  for (let item = ctx.story.tree.next(from); item; item = ctx.story.tree.next(item)) {
    if (closesContainer(item, toc, ctx.doc)) return false;
    if (fieldCharType(item) === 'begin' && !isHidden(ctx.doc, item)) return true;
  }

  return false;
}

/** Whether an edit of view characters `[start, end)` reaches the field's range (`contains`). */
function reachesRange(range: TocFieldRange, length: number, start: number, end: number): boolean {
  if (start !== end) return start < range.end && end > range.start;
  if (range.endsHere && !range.startsHere && range.end === 0) return false;
  if (range.startsHere && range.start > 0 && start <= range.start) return false;
  if (range.endsHere && range.end < length && start >= range.end) return false;
  return start >= range.start && start <= range.end;
}

/**
 * Whether an edit of view characters `[start, end)` (a caret when they are equal) in a paragraph
 * of the table of contents control `toc` reaches what its TOC field covers.
 */
export function reachesTocField(
  ctx: Ctx,
  toc: OrderedItem,
  p: OrderedViewParagraph,
  start: number,
  end: number
): boolean {
  const range = tocFieldRange(ctx, toc, p);
  return range !== null && reachesRange(range, p.chars.length, start, end);
}
