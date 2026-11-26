// The first difference between two documents in every field a checkpoint stores (design D3): for
// the tests, and for the check a room's close runs on the file it wrote before it deletes the old.

import { DOCUMENT_CODEC } from './snapshot-write.ts';
import { loadedSpace, type OrderedDocument } from './types.ts';

/**
 * The path of the first stored field where `b` differs from `a`, or null. `b` may be open under
 * another Word file number: a loaded space of `a`'s file compares as `b`'s file names it.
 */
export function snapshotDifference(a: OrderedDocument, b: OrderedDocument): string | null {
  return DOCUMENT_CODEC.diff(a, b, 'document', {
    space: (space) => {
      const loaded = loadedSpace(space);
      return loaded === null || loaded.file !== a.file
        ? space
        : `${loaded.kind}:${b.file}:${loaded.node}`;
    },
  });
}
