// The organiser's undo and redo of a direct edit (task 6.7): the suggestions the edit made and
// accepted at once turn rejected (taken back) or accepted again (made again). Nothing moves:
// every item the edit labelled shows exactly as before or after it, and the cards of other
// people's suggestions around those items are seen again (their versions move up).

import { setStatus, touchNeighbours } from '../model/mutations.ts';
import { sameMaker, type OrderedMaker } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';

/**
 * Turn `maker`'s direct suggestion to `status`. Refused when it is not theirs, not a direct
 * edit's (a decided suggestion is not), or not in the other status now.
 */
export function revertDirect(
  doc: OrderedDocument,
  suggestion: string,
  maker: OrderedMaker,
  status: 'accepted' | 'rejected'
): void {
  const found = doc.suggestions.get(suggestion);
  if (!found) throw new OrderedRefusal('unknown-suggestion', suggestion);
  if (!sameMaker(found.attribution, maker)) {
    throw new OrderedRefusal('not-own-suggestion', suggestion);
  }

  if (found.direct !== true) throw new OrderedRefusal('not-direct', suggestion);
  const from = status === 'accepted' ? 'rejected' : 'accepted';
  if (found.status !== from)
    throw new OrderedRefusal('stale-revert', `${suggestion} is ${found.status}`);
  setStatus(doc, suggestion, status);
  touchNeighbours(doc, suggestion);
}
