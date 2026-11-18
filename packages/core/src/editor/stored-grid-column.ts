// The stored grid column a table cell starts at.

import { findNode, flattenContentControls, type OoxmlPart } from '@docx-editor.dev/core/store';
import { buildCellGridSlots } from '../store/store/tree-op-table-vmerge.ts';

/**
 * The authored grid column `cellId` starts at, or null when the table does not hold it.
 *
 * Read from the tree, by the rules the store's table operations use, not from layout: a view
 * that lays a proposed column out as accepted or rejected draws the cells in a grid the
 * document does not have, and commands address the store.
 */
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
