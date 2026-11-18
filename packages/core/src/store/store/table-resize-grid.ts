import type { OoxmlElement, OoxmlTableCellNode } from '../package/ooxml-tree.ts';
import { readTwipsMeasure } from '../units.ts';
import type { EditableTableTopology } from './tree-op-table-topology.ts';
import type { TreeOpRejection } from './tree-op-types.ts';
import { wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';

function spanOf(cell: OoxmlTableCellNode): number {
  const properties = wmlChildNamed(cell, 'tcPr');
  const span = properties && wmlChildNamed(properties, 'gridSpan');
  if (!span) return 1;
  const value = wmlAttributeValue(span, 'val');
  return value && /^\d{1,4}$/.test(value) ? Number(value) : 0;
}

/** The grid columns a row skips before or after its cells (`w:gridBefore`, `w:gridAfter`). */
function skipOf(row: OoxmlElement, name: 'gridBefore' | 'gridAfter'): number {
  const properties = wmlChildNamed(row, 'trPr');
  const skip = properties && wmlChildNamed(properties, name);
  if (!skip) return 0;
  const value = wmlAttributeValue(skip, 'val');
  return value && /^\d{1,4}$/.test(value) ? Number(value) : -1;
}

/**
 * Resizing changes grid widths without adding or removing merged cells. A merge keeps its cells:
 * a horizontal one (`w:gridSpan`) takes the sum of its columns, a vertical one (`w:vMerge`) is
 * one cell per row. A row may skip columns before or after its cells. The legacy `w:hMerge`
 * stays refused.
 */
export function validateTableResizeGrid(topology: EditableTableTopology): TreeOpRejection | null {
  if (!topology.grid || topology.gridColumns.length === 0) return 'unknown-grid-column';
  if (topology.rows.length === 0) return 'tree-invariant';
  for (const { row, cells } of topology.rows) {
    const before = skipOf(row, 'gridBefore');
    const after = skipOf(row, 'gridAfter');
    if (before < 0 || after < 0) return 'tree-invariant';
    let columns = before;
    for (const cell of cells) {
      const properties = wmlChildNamed(cell, 'tcPr');
      if (properties && wmlChildNamed(properties, 'hMerge')) return 'table-has-merge';
      const span = spanOf(cell);
      if (span < 1 || span > topology.gridColumns.length - columns) return 'tree-invariant';
      columns += span;
    }
    if (columns + after !== topology.gridColumns.length) return 'tree-invariant';
  }
  return null;
}

/** After validation, update each affected cell from the sum of its covered columns. */
export function* resizedGridCells(
  topology: EditableTableTopology,
  widths: ReadonlyMap<number, number>
): Generator<{ cell: OoxmlTableCellNode; width: number }> {
  for (const { row, cells } of topology.rows) {
    let column = skipOf(row, 'gridBefore');
    for (const cell of cells) {
      const end = column + spanOf(cell);
      let affected = false;
      let width = 0;
      for (; column < end; column++) {
        affected ||= widths.has(column);
        width +=
          widths.get(column) ??
          readTwipsMeasure(wmlAttributeValue(topology.gridColumns[column]!, 'w'))!;
      }
      if (affected) yield { cell, width };
    }
  }
}
