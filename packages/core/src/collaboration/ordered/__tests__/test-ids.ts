// An id minter for tests that name an edit by a short label: the label becomes the namespace of
// a change id, so two edits with the same label make the same ids, as two copies of one change.

import { orderedIdMinter, type OrderedIdMinter } from '../change/ids.ts';

/** The minter of the change `<label>:1`, the label cut to a valid namespace. */
export function testIds(label: string): OrderedIdMinter {
  const namespace = label.replace(/[^A-Za-z0-9_-]/g, '_').slice(0, 64) || 'x';
  return orderedIdMinter(`${namespace}:1`);
}
