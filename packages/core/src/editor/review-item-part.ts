// The story part a review item's changes live in.
//
// Read from the part name every range carries, not from the site's id: a node a collaboration
// room created has a logical id (`lid:…`) that names no part, so its changes dropped out of
// Accept All, Reject All and author changes. A change no range places (a styles change) falls
// back to the id.

import type { ReviewItem } from '../contracts/editor.ts';
import { firstReviewRange, revisionSiteNodeIdsOf } from '../store/store/review-items.ts';
import type { OoxmlPart } from '../store/index.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';
import { partOfNodeId } from './surface-scope.ts';

/** The part holding `item`'s changes, read from the live package without opening a store. */
export function partOfReviewItem(
  session: PaginatedSurface['session'],
  item: ReviewItem
): OoxmlPart | null {
  const partName = firstReviewRange(item)?.partName;
  return (
    (partName === undefined ? undefined : session.currentPackage().parts.get(partName)) ??
    partOfNodeId(session, revisionSiteNodeIdsOf(item)[0])
  );
}
