// A suggestion's attribution changed without a decision (group 9.10, the `attribution` change
// kind). Only this function writes `OrderedSuggestion.attribution` after the suggestion is
// made, inside an undo record: the record takes the write back.

import { OrderedRefusal } from './refusal.ts';
import type { OrderedAttribution, OrderedDocument } from './types.ts';
import { recordUndo } from './undo-record.ts';

/**
 * Give a suggestion another attribution. Its version moves up, so a decision on its card that
 * an organiser saw before is stale.
 */
export function setAttribution(
  doc: OrderedDocument,
  id: string,
  attribution: OrderedAttribution
): void {
  if (doc.undo === null) throw new Error('An attribution is written inside an undo record.');
  const suggestion = doc.suggestions.get(id);
  if (!suggestion) throw new OrderedRefusal('unknown-suggestion', id);
  const before = suggestion.attribution;
  suggestion.attribution = attribution;
  suggestion.version += 1;
  recordUndo(doc, () => {
    suggestion.attribution = before;
    suggestion.version -= 1;
  });
}
