// `ReviewItemQuery.wholeParagraphFormat`: a paragraph's own formatting change as its whole
// paragraph in the review queries.
//
// The store anchors `w:pPrChange` where the paragraph mark sits, as a collapsed range at the
// paragraph end (`review-site-locations.ts`). That point stays the default, because the caret
// rule depends on it. A host that draws its own review chrome can ask for the paragraph
// instead: its lines as bands and hit targets (`review-items-at.ts`), and a caret anywhere in
// it for `isActive` (`docx-editor.ts`).
//
// Only `w:pPrChange` widens. A change on the paragraph mark (`w:pPr/w:rPr/w:rPrChange`) has
// the same point range, but the item does not say which `w:rPrChange` sits on the mark, and
// one decision can join the mark's change with changes to runs over text.

import type { ReviewItemQuery } from '../contracts/editor.ts';
import {
  reviewItemKey,
  reviewItemRanges,
  type ReviewItem,
  type ReviewPosition,
  type ReviewRange,
} from '../store/store/review-items.ts';

/**
 * Whether `query` widens `range` of `item` to the paragraph it sits in. The range is then
 * collapsed at the paragraph end, so its offset is the paragraph's length.
 */
export function standsForParagraph(
  item: ReviewItem,
  range: ReviewRange,
  query: ReviewItemQuery | undefined
): boolean {
  return (
    query?.wholeParagraphFormat === true &&
    item.kind === 'revision' &&
    item.revisionKind === 'format' &&
    item.formattingKind === 'pPrChange' &&
    range.start.paragraphId === range.end.paragraphId &&
    range.start.offset === range.end.offset
  );
}

/**
 * `active` plus the key of each paragraph formatting change whose paragraph holds the caret,
 * when `query` asks for whole paragraphs. A change that the caret may not activate stays
 * out, as it does from the caret rule.
 */
export function withActiveParagraphChanges(
  active: ReadonlySet<string>,
  items: readonly ReviewItem[],
  query: ReviewItemQuery | undefined,
  caret: () => ReviewPosition | null,
  activatable: (item: ReviewItem) => boolean
): ReadonlySet<string> {
  if (query?.wholeParagraphFormat !== true) return active;
  const head = caret();
  if (head === null) return active;
  let keys: Set<string> | null = null;
  for (const item of items) {
    const holds = reviewItemRanges(item).some(
      (range) =>
        range.start.paragraphId === head.paragraphId && standsForParagraph(item, range, query)
    );
    if (holds && activatable(item)) (keys ??= new Set(active)).add(reviewItemKey(item));
  }
  return keys ?? active;
}
