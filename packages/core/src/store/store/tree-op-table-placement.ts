// Table placement for `setTableProperties`: horizontal alignment (`w:tblPr/w:jc`, §17.4.28)
// and preferred width (`w:tblPr/w:tblW`, §17.4.63).
//
// Each value either writes its leaf or, as `null`, removes it so the table style's value
// applies again. The values are in document terms: `left` is the leading edge, so a
// right-to-left table (`w:bidiVisual`) places `left` on the right. Callers that think in
// screen terms map before they get here.
//
// `validateTableProperties` admits the two keys and checks their values here, and
// `applyTableProperties` hands its table here to place. A tracked op meets the same check
// through its untracked form (`validateTrackedTableEdit`).

import type { OoxmlElement } from '../package/ooxml-tree.ts';
import { MAX_TABLE_WIDTH_TWIPS } from './table-constraints.ts';
import { removeTblPrChild } from './tree-op-table-properties.ts';
import { wmlChildNamed } from './tree-op-table-shared.ts';
import type { TreeDocOp, TreeOpRejection } from './tree-op-types.ts';

type Op = Extract<TreeDocOp, { op: 'setTableProperties' }>;

/** A `w:tblPr` leaf with the attributes to write, or null attributes to remove it. */
type PlacementLeaf = readonly [name: 'jc' | 'tblW', attributes: Record<string, string> | null];

const PLACEMENT_KEYS: readonly string[] = ['justification', 'preferredWidth'];
/** A percentage width counts fiftieths of a percent; 5000 is the full text column. */
const MAX_TABLE_WIDTH_FIFTIETHS = 5000;
const JUSTIFICATIONS: readonly unknown[] = ['left', 'center', 'right'];

function invalidPreferredWidth(width: unknown): boolean {
  if (width === undefined || width === null) return false;
  if (typeof width !== 'object' || Object.keys(width).sort().join() !== 'type,value') return true;
  const { type, value } = width as { readonly type: unknown; readonly value: unknown };
  if (typeof value !== 'number' || !Number.isInteger(value)) return true;
  if (type === 'auto') return value !== 0;
  const max =
    type === 'dxa' ? MAX_TABLE_WIDTH_TWIPS : type === 'pct' ? MAX_TABLE_WIDTH_FIFTIETHS : 0;
  return value < 1 || value > max;
}

/** The op's keys other than the placement ones, whose values `validatePlacement` checks. */
export function nonPlacementKeys(op: Op): readonly string[] {
  return Object.keys(op).filter((key) => !PLACEMENT_KEYS.includes(key));
}

/** Why the op's placement values are refused: malformed or out of range. */
export function validatePlacement(op: Op): TreeOpRejection | null {
  const { justification } = op;
  if (justification !== undefined && justification !== null) {
    if (!JUSTIFICATIONS.includes(justification)) return 'invalidArgs';
  }
  // `columnWidthsTwips` writes the table width as well, so the two cannot share one op.
  if (op.preferredWidth !== undefined && op.columnWidthsTwips !== undefined) return 'invalidArgs';
  return invalidPreferredWidth(op.preferredWidth) ? 'invalidArgs' : null;
}

/** The `w:tblPr` leaves the op writes or removes, in the order it names them. */
function tablePlacementLeaves(op: Op): readonly PlacementLeaf[] {
  const leaves: PlacementLeaf[] = [];
  if (op.justification !== undefined) {
    leaves.push(['jc', op.justification === null ? null : { val: op.justification }]);
  }
  const width = op.preferredWidth;
  if (width !== undefined) {
    leaves.push(['tblW', width === null ? null : { w: String(width.value), type: width.type }]);
  }
  return leaves;
}

function withoutTblPrChild(table: OoxmlElement, name: string): OoxmlElement {
  const props = wmlChildNamed(table, 'tblPr');
  if (!props) return table;
  const removed = removeTblPrChild(props, name);
  if (!removed.ok) return table;
  return {
    ...table,
    children: table.children.map((node) => (node.id === props.id ? removed.container : node)),
  } as OoxmlElement;
}

/**
 * The table with the op's placement leaves in place: `write` sets a leaf as the op sets its
 * others, and a null value removes the leaf.
 */
export function withPlacement(
  table: OoxmlElement,
  op: Op,
  write: (table: OoxmlElement, name: string, attributes: Record<string, string>) => OoxmlElement
): OoxmlElement {
  let placed = table;
  for (const [name, attributes] of tablePlacementLeaves(op)) {
    placed = attributes ? write(placed, name, attributes) : withoutTblPrChild(placed, name);
  }
  return placed;
}
