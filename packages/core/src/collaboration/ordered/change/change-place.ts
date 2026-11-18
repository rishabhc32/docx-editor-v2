// A change's place as the model's edits take it.

import {
  resolveOrderedPlace,
  type OrderedPlace,
  type OrderedResolvedPlace,
} from '../edits/places.ts';
import type { OrderedDocument } from '../model/types.ts';
import type { OrderedChangePlace } from './types.ts';

/** The model's place for a change's place: a paragraph's mark, or a unit's address. */
export function placeOfChange(place: OrderedChangePlace): OrderedPlace {
  return 'mark' in place ? place.mark : place;
}

/** The view paragraph and offset of a change's place in a story. */
export function resolveChangePlace(
  doc: OrderedDocument,
  storyName: string,
  place: OrderedChangePlace
): OrderedResolvedPlace {
  return resolveOrderedPlace(doc, storyName, placeOfChange(place));
}
