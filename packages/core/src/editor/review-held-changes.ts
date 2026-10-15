// Why a review decision is refused while a change inside what it removes is still pending.
//
// Rejecting a proposed row, cell or table — or accepting a proposed deletion of one — takes the
// structure's content with it. The author's own pending changes in there go along
// (`containedOwnChangeSites`); another reviewer's stay theirs to answer, so the decision is
// refused until they are. A bare "the revision was refused" left the reviewer nothing to act on.

import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { ReviewRevisionItem } from '../store/store/review-items.ts';
import { holdsPendingChange } from '../store/store/revision-contained-changes.ts';

/** Keyed in `@docx-editor.dev/i18n`'s refusal table; the two must stay the same sentence. */
export const HELD_CHANGE_REASON =
  'another pending change sits inside this one; accept or reject that change first';

/** The specific refusal for a decision held up by a pending change inside it, or null. */
export function heldChangeRefusal(
  part: OoxmlPart | null,
  item: ReviewRevisionItem,
  action: 'accept' | 'reject'
): string | null {
  return part && holdsPendingChange(part, item, action) ? HELD_CHANGE_REASON : null;
}
