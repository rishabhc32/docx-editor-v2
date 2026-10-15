// Suggesting mode for a table command that removes the whole table.
import type { TreeDocOp } from '../store/store/tree-op-types.ts';

/**
 * A table command's ops as the surface applies them, with a whole-table removal proposed.
 *
 * `deleteTable`, and Delete Rows or Delete Columns over the whole table, plan the table's
 * `deleteBlock`. That op is not revision-capable — the range-deletion lane removes covered
 * tables with it too, and its joins rely on the table being gone — so `attributeTrackedOps`
 * passed it through, and suggesting mode removed the table with no card to reject. This lane
 * alone attributes it, and the store proposes the deletion in the form Word writes.
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
