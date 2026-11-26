// The first difference between two documents in every field a checkpoint stores (design D14): for
// the tests, and for the check a room's close runs on the file it wrote before it deletes the old.

import { sideDifference } from './snapshot-side.ts';
import { DOCUMENT_CODEC } from './snapshot-write.ts';
import type { OrderedDocument } from './types.ts';

/** The path of the first stored field or side state where `b` differs from `a`, or null. */
export function snapshotDifference(a: OrderedDocument, b: OrderedDocument): string | null {
  return DOCUMENT_CODEC.diff(a, b, 'document') ?? sideDifference(a, b);
}
