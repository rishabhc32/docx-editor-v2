import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlPart } from '../package/ooxml-tree.ts';
import type { ReviewRevisionItem } from './review-items.ts';
import { tableChildren } from './revision-table-children.ts';
import { buildCellGridSlots } from './tree-op-table-vmerge.ts';

// A proposed column is one decision over every cell of the column, `w:cellDel` or `w:cellIns`
// on each. Its card read as a cell change, which is not how a reviewer thinks of it.

type StructuralChange = NonNullable<ReviewRevisionItem['structuralChanges']>[number];

/** The table cell a `w:tcPr` marker sits on, and the nearest table that holds it. */
function markedCell(
  part: OoxmlPart,
  markerId: string
): { readonly cell: OoxmlElement; readonly table: OoxmlElement } | null {
  const properties = parentNodeOf(part, markerId);
  const cell = properties && parentNodeOf(part, properties.id);
  if (!cell || cell.kind !== 'tableCell') return null;
  let table = parentNodeOf(part, cell.id);
  while (table && table.kind !== 'table') table = parentNodeOf(part, table.id);
  return table ? { cell, table } : null;
}

/**
 * How many whole columns of one table a decision's cell insertions or deletions make up.
 *
 * Whole means that in every row of the table the decision's cells cover the same grid
 * columns, and no other cell covers them. The count is the columns a reader sees: grid
 * columns that no cell boundary separates count once. Absent for a decision that is not one
 * kind of cell change, that reaches into another table, or that leaves part of a column out.
 */
function structuralColumnsOf(
  part: OoxmlPart,
  changes: readonly StructuralChange[] | undefined,
  markerIds: readonly string[]
): number | undefined {
  if (changes?.length !== 1 || (changes[0] !== 'cellDelete' && changes[0] !== 'cellInsert'))
    return undefined;
  const cells = new Set<string>();
  let table: OoxmlElement | undefined;
  for (const id of markerIds) {
    const found = markedCell(part, id);
    if (!found || (table && found.table !== table)) return undefined;
    table = found.table;
    cells.add(found.cell.id);
  }
  if (!table) return undefined;
  // Grid intervals the decision covers, adjacent ones joined, the same in every row.
  let spans: [number, number][] | undefined;
  let matched = 0;
  const starts = new Set<number>();
  for (const row of tableChildren(table, 'tableRow')) {
    if (row.kind !== 'tableRow') continue;
    const rowCells = tableChildren(row, 'tableCell');
    const slots = buildCellGridSlots(row, rowCells);
    const covered: [number, number][] = [];
    for (const [index, cell] of rowCells.entries()) {
      const { startCol: at, span } = slots[index]!;
      starts.add(at);
      if (!cells.has(cell.id)) continue;
      matched += 1;
      const last = covered[covered.length - 1];
      if (last && last[1] === at) last[1] = at + span;
      else covered.push([at, at + span]);
    }
    if (covered.length === 0 || (spans && covered.join(';') !== spans.join(';'))) return undefined;
    spans = covered;
  }
  if (!spans || matched !== cells.size) return undefined;
  let columns = 0;
  for (const start of starts)
    if (spans.some(([from, to]) => start >= from && start < to)) columns += 1;
  return columns;
}

/** The `structuralColumns` field of a review item, present only when it has a count. */
export function structuralColumnsField(
  part: OoxmlPart,
  changes: readonly StructuralChange[] | undefined,
  markerIds: readonly string[]
): Pick<ReviewRevisionItem, 'structuralColumns'> {
  const columns = structuralColumnsOf(part, changes, markerIds);
  return columns === undefined ? {} : { structuralColumns: columns };
}
