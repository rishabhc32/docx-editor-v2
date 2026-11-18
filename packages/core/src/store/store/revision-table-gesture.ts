// A tracked table resize resolves as one decision: it records the table's grid
// (`w:tblGridChange`) beside the row and cell records of its table, all under one revision
// (`withTablePropertyChanges`).

import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { tableOf } from './revision-table-implicit-height.ts';
import type { RevisionSite } from './tree-op-revisions.ts';
import { wmlAttributeValue } from './tree-op-table-shared.ts';
import { ROW_PROPERTY_CONTAINERS } from './tree-op-tracked-table-records.ts';

const ROW_RECORDS: ReadonlySet<string> = new Set(
  Object.values(ROW_PROPERTY_CONTAINERS).map(({ record }) => record)
);
const GESTURE_RECORDS: ReadonlySet<string> = new Set([...ROW_RECORDS, 'tblPrChange']);

/**
 * Whether a selected row record of the grid's table shares the grid record's `@w:id`. A grid
 * Word wrote is shared by every row decision of its table, so `deferredTableGridSites` keeps a
 * selected grid record standing while another of them survives; a grid selected with a row
 * record of its own revision is that gesture's, and the two resolve together.
 */
export function selectedWithItsGesture(
  part: OoxmlPart,
  grid: RevisionSite,
  sites: readonly RevisionSite[],
  selected: ReadonlySet<string>
): boolean {
  const table = grid.parent && parentNodeOf(part, grid.parent.id);
  if (!table) return false;
  const id = wmlAttributeValue(grid.node, 'id');
  return sites.some(
    (site) =>
      ROW_RECORDS.has(site.node.localName) &&
      selected.has(site.node.id) &&
      wmlAttributeValue(site.node, 'id') === id &&
      tableOf(part, site.node)?.id === table.id
  );
}

/**
 * The grid records a decision of `matched` takes along: a table's `w:tblGridChange` whose
 * gesture records (the table's, its rows' and its cells' under the grid's `@w:id`) are all
 * matched. Decided alone after them, the grid would be unused history
 * (`unboundTableHistories`), which a reject keeps, so the result would hang on the order. A
 * decision of the whole gesture by its author's address so takes the grid along: the grid
 * record carries no author to be addressed with it.
 */
export function gestureGridSites(
  part: OoxmlPart,
  sites: readonly RevisionSite[],
  matched: readonly RevisionSite[]
): RevisionSite[] {
  const selected = new Set(matched.map((site) => site.node.id));
  const grids: RevisionSite[] = [];
  for (const grid of sites) {
    if (grid.node.localName !== 'tblGridChange' || selected.has(grid.node.id)) {
      continue;
    }

    const table = grid.parent && parentNodeOf(part, grid.parent.id);
    if (!table) {
      continue;
    }

    const id = wmlAttributeValue(grid.node, 'id');
    const records = sites.filter(
      (site) =>
        GESTURE_RECORDS.has(site.node.localName) &&
        wmlAttributeValue(site.node, 'id') === id &&
        tableOf(part, site.node)?.id === table.id
    );
    if (records.length > 0 && records.every((site) => selected.has(site.node.id))) {
      grids.push(grid);
    }
  }

  return grids;
}
