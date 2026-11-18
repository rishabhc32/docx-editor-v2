// The caret an edit's ops publish, followed through the ids a collaboration room assigns.

import type { TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import type { SemanticPosition } from '@docx-editor.dev/core/layout';

/** The caret one commit publishes, resolved after the commit settles. */
export interface CommittedCaret {
  /** Where the caret stands now, or null when the commit published none. */
  readonly position: () => SemanticPosition | null;
  readonly stop: () => void;
}

/**
 * Watch the next commit for its caret. `order` reads the story from the tree, not layout.
 *
 * A new table or row names its first cell by the id the store minted (`#new:2`), and a shared
 * room renames created nodes before the caret is placed. The rename keeps the reading order,
 * so the caret is kept as a position in that order. The store publishes the caret before the
 * room renames (`canonical-primitive-capture.ts`), so the position is read while the minted id
 * stands.
 */
export function watchCommittedCaret(
  session: Pick<TreeDocxSessionView, 'subscribe'>,
  order: () => readonly string[]
): CommittedCaret {
  let published: {
    readonly paragraphId: string;
    readonly offset: number;
    readonly index: number;
    readonly count: number;
  } | null = null;
  const stop = session.subscribe((change) => {
    if (!change.caret) return;
    const { paragraphId, start } = change.caret;
    const ids = order();
    published = { paragraphId, offset: start, index: ids.indexOf(paragraphId), count: ids.length };
  });
  return {
    stop,
    position: () => {
      if (!published) return null;
      const ids = order();
      if (ids.includes(published.paragraphId)) {
        return { paragraphId: published.paragraphId, offset: published.offset };
      }

      // A position names the same paragraph only while the story has the same paragraphs.
      const paragraphId = ids.length === published.count ? ids[published.index] : undefined;
      return paragraphId === undefined ? null : { paragraphId, offset: published.offset };
    },
  };
}
