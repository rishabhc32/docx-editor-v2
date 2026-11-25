// Where an op's offsets land against the spans of the content controls in a paragraph.

import type { OffsetSpan } from './tree-op-segments.ts';

/**
 * Whether a control's own characters overlap the offsets an op addresses.
 *
 * THE TWO EDGES ARE DIFFERENT PLACES. The applier gives a boundary offset to the run that starts
 * there, so a point op that writes content at a control's leading edge lands in the control's
 * first run, and the same op at its trailing edge lands in whatever follows. A point that writes
 * nothing (a comment marker, a paragraph split) is beside the control at either edge.
 *
 * A non-empty range that so much as touches the inside is refused WHOLE: a partial edit that
 * clipped itself to the unlocked side would silently do something other than what was asked.
 */
export function intersects(span: OffsetSpan | null, range: OffsetSpan, writes: boolean): boolean {
  if (span === null) return false;
  if (range.end <= range.start) {
    return writes
      ? range.start >= span.start && range.start < span.end
      : range.start > span.start && range.start < span.end;
  }
  return range.start < span.end && range.end > span.start;
}

/** Whether an op addresses only characters this control holds: the "filling in a field" case. */
export function encloses(span: OffsetSpan | null, range: OffsetSpan, writes: boolean): boolean {
  if (span === null) return false;
  if (range.end <= range.start) return intersects(span, range, writes);
  return range.start >= span.start && range.end <= span.end;
}

/** The part of `range` inside `within` (a caret stays as it is). */
export function clipped(range: OffsetSpan, within: OffsetSpan | null): OffsetSpan {
  if (within === null || range.end <= range.start) {
    return range;
  }
  return { start: Math.max(range.start, within.start), end: Math.min(range.end, within.end) };
}

/**
 * Whether `range` reaches a place none of `spans` holds: a character outside them all, or a caret
 * none of them takes (the edges as {@link intersects} reads them).
 */
export function reachesOutside(
  spans: readonly (OffsetSpan | null)[],
  range: OffsetSpan,
  writes: boolean
): boolean {
  if (range.end <= range.start) {
    return !spans.some((span) => intersects(span, range, writes));
  }
  const sorted = spans
    .filter((span): span is OffsetSpan => span !== null)
    .sort((a, b) => a.start - b.start);
  let from = range.start;
  for (const span of sorted) {
    if (span.start > from) {
      return true;
    }
    from = Math.max(from, span.end);
    if (from >= range.end) {
      return false;
    }
  }
  return true;
}
