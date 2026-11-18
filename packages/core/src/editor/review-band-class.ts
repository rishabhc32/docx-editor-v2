// The class a review band draws in on the page: a comment's, a custom node's while its card
// is open, or a revision's in its own colour. Pure, so the surface asks it for each band it
// paints and nothing else.

import type { ReviewItem, ReviewRevisionKind } from '../store/store/review-items.ts';

/** Separates a decision's key from its per-range index. A NUL cannot occur in either. */
export const RANGE_SUFFIX = '\u0000range\u0000';

/** The class a band draws in, or null when this range should not be drawn at all. */
export function bandClassFor(
  key: string,
  activeKeys: ReadonlySet<string>,
  byKey: ReadonlyMap<string, ReviewItem>
): string | null {
  // A revision's key is suffixed per range when one decision covers several sites, so the
  // active test compares the DECISION, not the site. Split on NUL: an author name cannot hold
  // it, but it can hold `#`.
  const parts = key.split(RANGE_SUFFIX);
  const decision = parts[0]!;
  const isActive = activeKeys.has(decision);
  if (key.startsWith('comment-')) {
    const item = byKey.get(decision);
    const resolved = item?.kind === 'comment' && item.resolved;
    if (resolved && !isActive) return null;
    const classes = ['docx-comment-band'];
    if (isActive) {
      classes.push('docx-comment-band--active');
    }

    if (isActive && resolved) {
      classes.push('docx-comment-band--resolved-active');
    }

    return classes.join(' ');
  }

  // A custom node's band only while its card is open: the chip tint already marks the
  // node persistently, and the comment band is the right weight for "this is the one".
  if (key.startsWith('custom-')) {
    return isActive ? 'docx-comment-band docx-comment-band--active' : null;
  }

  const item = byKey.get(decision);
  if (!item || item.kind !== 'revision') return null;

  // In the revision's own colour (green for text arriving, red for text leaving), as the
  // decoration beneath it. A replacement's leading ranges are the struck half; the rest is
  // what took their place.
  const index = parts.length > 1 ? Number(parts[1]) : 0;
  const kindOf = (revisionKind: ReviewRevisionKind): 'delete' | 'insert' | 'other' => {
    if (revisionKind === 'delete' || revisionKind === 'moveFrom') return 'delete';
    if (revisionKind === 'insert' || revisionKind === 'moveTo') return 'insert';
    if (revisionKind === 'replace' && item.replacedRangeCount !== undefined) {
      return index < item.replacedRangeCount ? 'delete' : 'insert';
    }

    return 'other';
  };
  const kind = kindOf(item.revisionKind);

  // Every pending change has a faint band; the open one is strong.
  return `docx-revision-band docx-revision-band--${kind}${isActive ? ' docx-revision-band--active' : ''}`;
}
