// TRACKED table edits — the table ops that carry a `revision`, validated and applied in one
// place, so suggesting mode proposes a table edit rather than applying it outright.
//
// A property op — cell fill, vertical alignment and borders, row height, column and right-edge
// resize, table properties — writes exactly as it does untracked, then records what it replaced
// (`withTablePropertyChanges`). A column insertion or deletion marks the column's cells
// (`tree-op-tracked-table-columns.ts`), and a table's `deleteBlock` marks every row deleted
// (`tree-op-tracked-table-deletion.ts`).
//
// A tracked op meets every rule of its untracked form first, so the two cannot drift.

import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { applyTableProperties } from './tree-op-table-authoring.ts';
import { applyTableCellPropertyOp, applyTableResizeOp } from './tree-op-tables.ts';
import {
  applyTrackedColumnEdit,
  validateTrackedColumnEdit,
} from './tree-op-tracked-table-columns.ts';
import {
  applyTrackedTableDeletion,
  proposesTableDeletion,
  validateTrackedTableDeletion,
} from './tree-op-tracked-table-deletion.ts';
import { withTablePropertyChanges } from './tree-op-tracked-table-properties.ts';
import type { TreeDocOp, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

/**
 * The table ops suggesting mode attributes. A table's `deleteBlock` is proposed too, but the
 * range-deletion lane removes covered tables with that op, so the surface attributes it apart.
 */
export const REVISION_CAPABLE_TABLE_OPS = [
  'insertTableColumn',
  'deleteTableColumn',
  'setTableCellFill',
  'setTableCellVerticalAlignment',
  'setTableCellBorders',
  'setTableRowHeight',
  'setTableColumnWidths',
  'setTableRightEdgeWidth',
  'setTableProperties',
] as const satisfies readonly TreeDocOp['op'][];

type TrackedTableOp = Extract<
  TreeDocOp,
  { op: (typeof REVISION_CAPABLE_TABLE_OPS)[number] | 'deleteBlock' }
>;

const TRACKED_TABLE_OPS: ReadonlySet<string> = new Set([
  ...REVISION_CAPABLE_TABLE_OPS,
  'deleteBlock',
]);

/**
 * Whether `op` is a table op that names a revision. An own `revision: undefined` names none,
 * so the op is untracked, but the untracked validators refuse a key they do not know: the key
 * comes off here too.
 */
function namesRevision(op: TreeDocOp): op is TrackedTableOp {
  return TRACKED_TABLE_OPS.has(op.op) && 'revision' in op;
}

/** Each member of `T` without its `revision`. */
type Untracked<T> = T extends unknown ? Omit<T, 'revision'> : never;

/** The op as an untracked write: the same op without its `revision`. */
function untrackedFormOf<T extends TrackedTableOp>(op: T): Untracked<T> {
  const { revision: _revision, ...untracked } = op;
  return untracked as Untracked<T>;
}

/**
 * Why a table op that names a revision is refused, or null when it may proceed; undefined for
 * any other op. `validateUntracked` validates the op's untracked form.
 */
export function validateTrackedTableEdit(
  part: OoxmlPart,
  op: TreeDocOp,
  validateUntracked: (op: TreeDocOp) => TreeOpRejection | null
): TreeOpRejection | null | undefined {
  if (!namesRevision(op)) return undefined;
  const untracked = validateUntracked(untrackedFormOf(op));
  if (untracked || op.revision === undefined) return untracked;
  if (invalidRevisionAttribution(op.revision)) return 'invalid-property-value';
  if (op.op === 'deleteBlock') return validateTrackedTableDeletion(part, op, op.revision);
  if (op.op === 'insertTableColumn' || op.op === 'deleteTableColumn') {
    return validateTrackedColumnEdit(part, op, op.revision);
  }
  return null;
}

/**
 * Apply a table op that names a revision, after `validateTrackedTableEdit` admitted it. Null
 * for any other op, and for one the store applies as it is: a `deleteBlock` that withdraws a
 * table the author proposed, or a column op that names no author.
 */
export function applyTrackedTableEdit(
  part: OoxmlPart,
  op: TreeDocOp,
  options?: EditOptions
): TreeOpResult | null {
  if (!namesRevision(op)) return null;
  switch (op.op) {
    case 'deleteBlock':
      return proposesTableDeletion(part, op) ? applyTrackedTableDeletion(part, op, options) : null;
    case 'insertTableColumn':
    case 'deleteTableColumn':
      return op.revision === undefined
        ? null
        : applyTrackedColumnEdit(part, op, op.revision, options);
    case 'setTableProperties':
      return withTablePropertyChanges(
        part,
        applyTableProperties(part, untrackedFormOf(op), options),
        op,
        options
      );
    case 'setTableCellBorders':
    case 'setTableCellFill':
    case 'setTableCellVerticalAlignment':
      return withTablePropertyChanges(
        part,
        applyTableCellPropertyOp(part, untrackedFormOf(op), options),
        op,
        options
      );
    default:
      return withTablePropertyChanges(
        part,
        applyTableResizeOp(part, untrackedFormOf(op), options),
        op,
        options
      );
  }
}
