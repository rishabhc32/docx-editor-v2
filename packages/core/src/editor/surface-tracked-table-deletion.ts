// Suggesting mode for a table command that removes the whole table.
import type { TreeDocOp } from '../store/store/tree-op-types.ts';

/**
 * A table command's ops as the surface applies them, with a whole-table removal proposed.
 *
 * `deleteTable`, and Delete Rows or Delete Columns over the whole table, plan the table's
 * `deleteBlock`. `attributeTrackedOps` passes that op through untracked, because the
 * range-deletion lane removes covered tables with it too and its joins rely on the table being
 * gone. This lane alone attributes it, and the store proposes the deletion in the form Word
 * writes.
 */
export function createTrackedTableDeletion(deps: {
  readonly trackedAuthorOrNone: () => string | undefined;
  readonly trackedDate: () => string;
}): (ops: readonly TreeDocOp[]) => readonly TreeDocOp[] {
  return (ops) => {
    const author = deps.trackedAuthorOrNone();
    if (author === undefined) return ops;
    const revision = { author, date: deps.trackedDate() };
    return ops.map((op) => (op.op === 'deleteBlock' ? { ...op, revision } : op));
  };
}
