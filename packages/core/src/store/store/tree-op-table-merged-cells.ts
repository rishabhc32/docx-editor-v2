// The whole merged cell around a cell of a vertical merge, and the order Delete Columns takes
// the columns of a table with merged cells.

import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { readEditableTableTopology } from './tree-op-table-topology.ts';
import { buildRowGridSlots } from './tree-op-table-vmerge.ts';
import { cellsCoveringColumn, tableHasColumnMergeMarkers } from './tree-op-table-merged-columns.ts';

/**
 * `cellIds` with every cell of each vertical merge one of them is in, in table order. Ids the
 * table does not hold, and a table the part does not hold, are left as they are, for the op's
 * own validation to refuse by name.
 *
 * Word draws a vertical merge (`w:vMerge restart`, then `continue` in the rows below on the same
 * grid interval) as one cell, and a cell op (fill, borders, vertical alignment) must name all of
 * its `w:tc` or it is refused (`invalidArgs`). So a caret in a merged cell, or a rectangle that
 * reaches part of it, means the whole merged cell. A row inserted inside a merge repeats
 * `w:vMerge`, so its cell belongs to the merged cell too.
 */
export function wholeMergedCellIds(
  part: OoxmlPart,
  tableId: string,
  cellIds: readonly string[]
): readonly string[] {
  const read = readEditableTableTopology(part.root, tableId);
  if (!read.ok || !read.topology.hasMerge) return cellIds;
  const groups: string[][] = [];
  const groupOf = new Map<string, string[]>();
  const open = new Map<string, string[]>();
  for (const { row, cells } of read.topology.rows) {
    const slots = buildRowGridSlots(row);
    for (const [index, cell] of cells.entries()) {
      const slot = slots[index];
      if (!slot) continue;
      const interval = `${slot.startCol}:${slot.span}`;
      let group = slot.vMergeKind === 'continue' ? open.get(interval) : undefined;
      if (!group) {
        group = [];
        groups.push(group);
      }

      if (slot.vMergeKind === 'restart') {
        open.set(interval, group);
      } else if (slot.vMergeKind === 'none') {
        open.delete(interval);
      }

      group.push(cell.id);
      groupOf.set(cell.id, group);
    }
  }

  const wanted = new Set<string>();
  for (const id of cellIds) {
    for (const member of groupOf.get(id) ?? [id]) {
      wanted.add(member);
    }
  }

  if (wanted.size === cellIds.length) return cellIds;
  const ordered = groups.flat().filter((id) => wanted.has(id));
  for (const id of cellIds) if (!groupOf.has(id)) ordered.push(id);
  return ordered;
}

/**
 * The grid columns a Delete Columns command deletes, in the order it deletes them. In a table
 * with merged cells each deletion builds the grid again (`tree-op-table-merged-columns.ts`) and
 * keeps the ids of the columns before the change, so the columns go from the last one. A column
 * whose covering cells an earlier deletion already takes is left out: Word deletes a merged
 * cell once. A table without merges keeps the order it was given. With no column selected, the
 * caret's column goes.
 */
export function columnDeletionOrder(
  part: OoxmlPart,
  tableId: string,
  selected: readonly string[],
  caretColumnId: string
): string[] {
  const gridColumnIds = selected.length > 0 ? selected : [caretColumnId];
  const read = readEditableTableTopology(part.root, tableId);
  if (!read.ok || !tableHasColumnMergeMarkers(read.topology)) return [...gridColumnIds];
  const { topology } = read;
  const index = (id: string) => topology.gridColumns.findIndex((column) => column.id === id);
  const taken = new Set<string>();
  const out: string[] = [];
  for (const id of [...gridColumnIds].sort((a, b) => index(b) - index(a))) {
    const covering = (cellsCoveringColumn(topology, id) ?? []).filter((cell) => cell !== null);
    if (covering.length > 0 && covering.every((cell) => taken.has(cell.id))) continue;
    for (const cell of covering) taken.add(cell.id);
    out.push(id);
  }

  return out;
}
