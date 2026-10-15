// The stored grid column a table cell starts at.
//
// Layout reports where a cell is DRAWN. A view that lays a proposed column out as accepted or
// rejected draws the surviving cells in a grid the document does not have, so the drawn column
// cannot name a stored `w:gridCol`: Insert Column beside a cell that took a hidden column's
// space landed on the wrong side of it. Commands address the store, so they read the column
// from the tree, by the same rules the store's table operations read it with.

import { findNode, flattenContentControls, type OoxmlPart } from '@docx-editor.dev/core/store';
import { buildCellGridSlots } from '../store/store/tree-op-table-vmerge.ts';

/** The authored grid column `cellId` starts at, or null when the table does not hold it. */
export function storedGridColumnOf(
  part: OoxmlPart,
  tableId: string,
  cellId: string
): number | null {
  const table = findNode(part, tableId);
  if (!table || table.kind !== 'table') return null;
  for (const row of flattenContentControls(table.children)) {
    if (row.kind !== 'tableRow') continue;
    const cells = flattenContentControls(row.children).filter((cell) => cell.kind === 'tableCell');
    const index = cells.findIndex((cell) => cell.id === cellId);
    if (index >= 0) return buildCellGridSlots(row, cells)[index]!.startCol;
  }
  return null;
}
