// Why suggesting mode refuses a table change on a row another author's pending change holds.
//
// Word answers a row's property records as one decision, so the store refuses a tracked table
// change to a row that holds another author's pending property change
// (`tree-op-tracked-table-properties.ts`). Its `detail` says so, but a commit reports only the
// refusal code, which tells the user nothing to do. The surface asks the store again with the
// same ops, and only after a refusal, so a change that lands costs nothing.

import type { ExecResult } from '../contracts/types.ts';
import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import { OTHER_AUTHORS_ROW_CHANGE } from '../store/store/tree-op-tracked-table-properties.ts';
import { applyTreeOp } from '../store/store/tree-ops.ts';

/** Keyed in `@docx-editor.dev/i18n`'s refusal table; the two must stay the same sentence. */
export const HELD_ROW_REASON =
  "another author's formatting change is pending on a row this change affects; accept or reject that change first";

/**
 * The refusal for table ops the store refused over another author's pending row change, or
 * null for any other refusal. `refused` is the commit's refusal and `ops` are the ops it
 * applied, attributed as the surface attributed them.
 */
export function heldTableRowRefusal(
  refused: string | null,
  part: OoxmlPart | null,
  ops: readonly TreeDocOp[]
): Extract<ExecResult, { ok: false }> | null {
  if (refused !== 'unsupported-revision' || part === null) return null;
  let current = part;
  for (const op of ops) {
    const result = applyTreeOp(current, op);
    if (!result.ok) {
      return result.detail === OTHER_AUTHORS_ROW_CHANGE
        ? { ok: false, code: 'unsupported', reason: HELD_ROW_REASON }
        : null;
    }
    current = result.part;
  }
  return null;
}
