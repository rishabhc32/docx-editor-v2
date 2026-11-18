// THE FACE TYPED TEXT TAKES BESIDE STRUCK WORDS.
//
// Word gives typed text the formatting of the character before the caret, and at a
// paragraph's start the formatting of the character after it: the rule
// `authoredRunPropertiesAt` states for the toolbar. Struck words are still in the document:
// the caret rests beside them, the toolbar reads them, and so they count. Read only: nothing
// here writes.

import type { OoxmlNode } from '../package/ooxml-tree.ts';
import { isContentRevisionKind, MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import type { ParagraphOffsetIndex } from './tree-op-segments.ts';

/** The struck runs that touch an offset, at one level of the walk. */
export interface StruckRunsAround {
  /** The struck run the offset sits in or right after: the character before the caret. */
  readonly left: OoxmlNode | null;
  /** The struck run the offset sits in or right before: the character after the caret. */
  readonly right: OoxmlNode | null;
}

/**
 * The struck runs touching `offset` among `nodes`, one level of the tracked walk. The tracked
 * insertion walk (`tree-op-tracked.ts`) never descends into a `w:del`, so it holds no run to
 * copy a face from beside one; this gives it the answer the untracked lane gives
 * (`insertBesideDeletion` in `tree-op-apply.ts`).
 *
 * Only revision wrappers are entered. A hyperlink, a field or a content control the walk
 * declined to enter keeps its face to itself: typing beside a link must not come out in the
 * link's character style, which is the same line the untracked lane draws.
 */
export function struckRunsAround(
  nodes: readonly OoxmlNode[],
  offset: number,
  offsets: ParagraphOffsetIndex
): StruckRunsAround {
  let left: OoxmlNode | null = null;
  let right: OoxmlNode | null = null;
  const visit = (children: readonly OoxmlNode[], struck: boolean, depth: number): void => {
    // File-controlled nesting: past the bound the runs simply lend no face.
    if (depth >= MAX_INLINE_CONTAINER_DEPTH) return;
    for (const node of children) {
      if (node.kind === 'textValue') continue;
      if (node.kind === 'run') {
        const span = struck ? offsets.spanOf(node) : null;

        // A run that measures nothing (field chrome, an empty run) is no character.
        if (!span || span.end <= span.start) continue;
        if (span.start < offset && offset <= span.end) left = node;
        if (right === null && span.start <= offset && offset < span.end) right = node;
        continue;
      }

      if (!isContentRevisionKind(node.kind)) continue;
      const deleted = node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom';
      visit(node.children, struck || deleted, depth + 1);
    }
  };
  visit(nodes, false, 0);
  return { left, right };
}
