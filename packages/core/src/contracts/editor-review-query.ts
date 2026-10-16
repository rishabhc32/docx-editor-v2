import type { ReviewRevisionKind } from './editor.ts';

/**
 * Narrows what `getReviewItems` returns.
 *
 * Filtering revision kinds is how a host hides structural cards it has no UI for,
 * `placement: false` skips the layout pass entirely when only metadata is wanted, and
 * `pairReplacements: false` lists a typed-over selection as its two halves.
 */
export interface ReviewItemQuery {
  /**
   * Revision kinds to leave out of the returned items, for example structural kinds that the
   * host has no card for. Default: none. Comments and custom items are never removed. This
   * filters only the returned data; to stop the caret from activating the same kinds, call
   * `setReviewActivationExclusions`.
   */
  readonly excludeRevisionKinds?: readonly ReviewRevisionKind[];
  /** When false, skip layout geometry; metadata is unchanged and anchors are null. Default true. */
  readonly placement?: boolean;
  /**
   * Whether to list a deletion and the insertion that directly follows it as one
   * `'replace'` item. Default true. Set it to false to list the deletion and the insertion as
   * separate items, for example to accept or reject each half on its own.
   *
   * A host that steps through the list with `setActiveReviewItem` stops once on
   * each replacement with the default, as Next Change does.
   *
   * Two revisions pair when they are next to each other among the queue's revisions, the
   * deletion comes first, both have the same author, nesting level and `readOnly` value, and
   * the deletion's last range ends at the same paragraph and offset where the insertion's
   * first range starts. The pair's `text` is the inserted words and `replacedText` the
   * deleted words. The two halves are not listed separately.
   *
   * Kinds in {@link excludeRevisionKinds} are removed before pairing, so an excluded deletion
   * or insertion never pairs. Excluding `'replace'` removes the pairs too.
   *
   * The pair's `key` is stable while both halves are pending, and every key-addressed verb
   * accepts it whatever query you read it with. Accepting or rejecting it resolves both
   * halves in one transaction and one undo step. Activating it opens the deletion; the pair
   * reports `isActive` while either half is active. A reply anchors over the inserted words,
   * and replies to either half are listed under the pair.
   */
  readonly pairReplacements?: boolean;
  /**
   * Whether a paragraph's own formatting change (`w:pPrChange`), such as a new alignment,
   * spacing, or indent, stands for its whole paragraph. Default false.
   *
   * The change covers no characters, so its range is collapsed at the paragraph end. By
   * default `getReviewItemRects` returns no bands for it, `getReviewItemsAt` never hits it,
   * and it is `isActive` only while the caret is at that end. When this is true:
   *
   * - `getReviewItemRects(key, query)` returns one band per line of the paragraph, on every
   *   painted page that holds a part of it. Each band spans the line's text column at glyph
   *   height, and the first band also covers the list number.
   * - `getReviewItemsAt` hits the change anywhere in those bands. It comes after every
   *   narrower item, so a change to a word in the paragraph, or a comment on it, comes first.
   * - `getReviewItems` reports the change `isActive` while the caret is anywhere in the
   *   paragraph, together with any narrower active item. A change that
   *   `setReviewActivationExclusions` excludes stays inactive.
   *
   * The item's ranges, `anchorY`, and `pageIndex` do not change. A formatting change on the
   * paragraph mark (`w:pPr/w:rPr/w:rPrChange`) keeps its collapsed range.
   * @public
   */
  readonly wholeParagraphFormat?: boolean;
}
