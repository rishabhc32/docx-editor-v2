import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlPart, OoxmlTableRowNode } from '../package/ooxml-tree.ts';
import type { ReviewRevisionItem } from './review-items.ts';
import { tableChildren } from './revision-table-children.ts';
import { buildCellGridSlots } from './tree-op-table-vmerge.ts';

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
  while (table && table.kind !== 'table') {
    table = parentNodeOf(part, table.id);
  }

  return table ? { cell, table } : null;
}

/** The cells the markers sit on, and their one table; null when they reach into two tables. */
function markedCells(
  part: OoxmlPart,
  markerIds: readonly string[]
): { readonly table: OoxmlElement; readonly cells: ReadonlySet<string> } | null {
  const cells = new Set<string>();
  let table: OoxmlElement | undefined;
  for (const id of markerIds) {
    const found = markedCell(part, id);
    if (!found || (table && found.table !== table)) return null;
    table = found.table;
    cells.add(found.cell.id);
  }

  return table ? { table, cells } : null;
}

/**
 * A row's grid intervals that `cells` cover, adjacent ones joined; the start column of each of
 * its cells; and how many of its cells are in `cells`.
 */
function rowCoverage(
  row: OoxmlTableRowNode,
  cells: ReadonlySet<string>
): { readonly covered: [number, number][]; readonly starts: number[]; readonly matched: number } {
  const rowCells = tableChildren(row, 'tableCell');
  const slots = buildCellGridSlots(row, rowCells);
  const covered: [number, number][] = [];
  const starts: number[] = [];
  let matched = 0;
  for (const [index, cell] of rowCells.entries()) {
    const { startCol: at, span } = slots[index]!;
    starts.push(at);
    if (!cells.has(cell.id)) continue;
    matched += 1;
    const last = covered[covered.length - 1];
    if (last && last[1] === at) {
      last[1] = at + span;
    } else {
      covered.push([at, at + span]);
    }
  }

  return { covered, starts, matched };
}

/**
 * How many whole columns of one table a decision's cell insertions or deletions make up: a
 * proposed column is one decision over every cell of the column, `w:cellDel` or `w:cellIns` on
 * each.
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
  if (changes?.length !== 1 || (changes[0] !== 'cellDelete' && changes[0] !== 'cellInsert')) {
    return undefined;
  }

  const marked = markedCells(part, markerIds);
  if (!marked) return undefined;

  // Grid intervals the decision covers, the same in every row
  let spans: [number, number][] | undefined;
  let matched = 0;
  const starts = new Set<number>();
  for (const row of tableChildren(marked.table, 'tableRow')) {
    if (row.kind !== 'tableRow') continue;
    const coverage = rowCoverage(row, marked.cells);
    for (const start of coverage.starts) {
      starts.add(start);
    }

    matched += coverage.matched;
    const { covered } = coverage;
    if (covered.length === 0 || (spans && covered.join(';') !== spans.join(';'))) return undefined;
    spans = covered;
  }

  if (!spans || matched !== marked.cells.size) return undefined;
  let columns = 0;
  for (const start of starts) {
    if (spans.some(([from, to]) => start >= from && start < to)) columns += 1;
  }

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
