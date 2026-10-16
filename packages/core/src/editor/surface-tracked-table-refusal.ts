// Why suggesting mode refuses a table change on a row another author's pending change holds.
//
// Word answers a row's property records as one decision, so the store refuses a tracked table
// change to a row that holds another author's pending property change
// (`tree-op-tracked-table-properties.ts`). Its `detail` says so, but a commit reports only the
// refusal code, which tells the user nothing to do. The surface asks the store again with the
// same ops, and only after a refusal, so a change that lands costs nothing.
//
// A shared room refuses a table formatting suggestion while it is disconnected too. Two authors
// can change one row before either sees the other's record, and the author whose change lost
// takes it back when the other change arrives (`surface-table-row-race.ts`). A change made
// offline arrives too late for that, so the session refuses it until the room is back.

import type { EditorCollaborationSession } from '../collaboration/index.ts';
import type { ExecResult } from '../contracts/types.ts';
import type { OoxmlPart } from '../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import { recordsTableFormatting } from '../store/store/table-row-owner.ts';
import { OTHER_AUTHORS_ROW_CHANGE } from '../store/store/tree-op-tracked-table-properties.ts';
import { applyTreeOp } from '../store/store/tree-ops.ts';

/** Keyed in `@docx-editor.dev/i18n`'s refusal table; the two must stay the same sentence. */
export const HELD_ROW_REASON =
  "another author's formatting change is pending on a row this change affects; accept or reject that change first";

/** Keyed in `@docx-editor.dev/i18n`'s refusal table; the two must stay the same sentence. */
export const OFFLINE_TABLE_REASON = 'reconnect to suggest table formatting';

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

/**
 * The refusal for a table formatting suggestion a disconnected room refused, or null for any
 * other refusal. The session refuses it with the code it gives every write that waits for the
 * room (`collaboration-session-not-ready`).
 */
export function offlineTableRefusal(
  refused: string | null,
  collaboration: Pick<EditorCollaborationSession, 'status'> | null | undefined,
  ops: readonly TreeDocOp[]
): Extract<ExecResult, { ok: false }> | null {
  return refused === 'collaboration-session-not-ready' &&
    collaboration?.status() === 'disconnected' &&
    ops.some(recordsTableFormatting)
    ? { ok: false, code: 'unsupported', reason: OFFLINE_TABLE_REASON }
    : null;
}
