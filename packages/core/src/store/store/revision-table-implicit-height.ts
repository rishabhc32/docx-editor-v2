import { implicitTableRowAlignments } from './revision-table-implicit-alignment.ts';
import { ooxmlTreesEqual } from '../package/ooxml-serialize.ts';
import { parentNodeOf } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlPart } from '../package/ooxml-tree.ts';
import { preservedRevisionProperty } from './revision-property-records.ts';
import type { RevisionSite } from './tree-op-revisions.ts';
function children(node: OoxmlElement, name: string): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  for (const child of node.children)
    if (
      child.kind !== 'textValue' &&
      child.namespaceUri === WML_NAMESPACE_URI &&
      child.localName === name
    )
      found.push(child);
  return found;
}
const attr = (node: OoxmlElement, name: string) =>
  node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)?.value;

/**
 * The one change record `name` of `container`, when it holds `container` as it is. A cell's
 * own structural marks (`w:cellIns`, a new column's) stay out of its record, so out of this.
 */
export function unchangedRecord(container: OoxmlElement, name: string): OoxmlElement | undefined {
  const records = children(container, name);
  const record = records[0];
  if (records.length !== 1 || !record) return undefined;
  const previous = children(record, container.localName);
  if (previous.length !== 1) return undefined;
  const current = {
    ...container,
    children: container.children.filter(
      (n) => n.id !== record.id && !preservedRevisionProperty(name, n)
    ),
  } as OoxmlElement;
  return ooxmlTreesEqual(current, previous[0]!) ? record : undefined;
}

/** Whether two records have one author and one date. */
export const sameActor = (left: OoxmlElement, right: OoxmlElement): boolean =>
  attr(left, 'author') === attr(right, 'author') && attr(left, 'date') === attr(right, 'date');

/**
 * The table's anchor for its rows' heights: an unchanged `w:tblPrChange`, beside a grid record.
 * The grid record holds the grid unchanged as Word writes it; a later column changes the grid
 * and keeps the record's original, and the heights still lean on the table's record.
 */
export function tableHeightAnchor(table: OoxmlElement): OoxmlElement | undefined {
  const tblPr = children(table, 'tblPr');
  const grid = children(table, 'tblGrid');
  if (tblPr.length !== 1 || grid.length !== 1 || children(grid[0]!, 'tblGridChange').length !== 1) {
    return undefined;
  }
  return unchangedRecord(tblPr[0]!, 'tblPrChange');
}

/** A row height Word added with no row record. */
export interface WordHeightRow {
  readonly row: OoxmlElement;
  /** The row's `w:trHeight`, its only row property. */
  readonly height: OoxmlElement;
  /** The unchanged `w:tcPrChange` of each cell: one author and date, whose height it is. */
  readonly cells: readonly OoxmlElement[];
  /** The row's own `w:tblPrExChange`, or the table's `w:tblPrChange` (`tableHeightAnchor`). */
  readonly anchor: OoxmlElement;
  readonly rowAnchor: boolean;
}

/** The row's `w:trHeight` when it is the row's only property, plain, as Word writes it. */
function plainHeight(row: OoxmlElement): OoxmlElement | undefined {
  const props = children(row, 'trPr');
  if (props.length !== 1) return undefined;
  const height = children(props[0]!, 'trHeight');
  // Other row properties are independent; never infer their old state. The row's own markers
  // (`w:ins`, `w:del`: a pending new or struck row) stay out of a row record, so out of this
  const own = props[0]!.children.filter((n) => !preservedRevisionProperty('trPrChange', n));
  if (height.length !== 1 || own.length !== 1 || height[0]!.children.length) return undefined;
  if (
    !/^\d+$/.test(attr(height[0]!, 'val') ?? '') ||
    height[0]!.attributes.some(
      (a) => a.namespaceUri !== WML_NAMESPACE_URI || !['val', 'hRule'].includes(a.localName)
    )
  )
    return undefined;
  const others = row.children.some(
    (n) =>
      n.kind !== 'textValue' &&
      (n.namespaceUri !== WML_NAMESPACE_URI || !['trPr', 'tblPrEx', 'tc'].includes(n.localName))
  );
  return others ? undefined : height[0];
}

/**
 * The rows of `table` that hold a height Word added with no row record, in Word's form. On a row
 * with no `w:trPr`, Word writes a new height with no `w:trPrChange`: it records an unchanged
 * `w:tcPrChange` on each cell of the row, and ANCHORS the change on the table's unchanged
 * `w:tblPrChange` and `w:tblGridChange` (a table holds one `w:tblPrChange`, so a second author's
 * height in the table shares the first one's), or on the row's own unchanged `w:tblPrExChange`
 * (Word's form once the table's record is gone). The cells' records say whose height it is.
 * Word for Mac, 2026-10-30: `__tests__/fixtures/word-created-row-height.ts`,
 * `__tests__/fixtures/word-row-height-authors.ts`.
 */
export function wordHeightRows(table: OoxmlElement): WordHeightRow[] {
  const tableAnchor = tableHeightAnchor(table);
  const found: WordHeightRow[] = [];
  for (const row of children(table, 'tr')) {
    const height = plainHeight(row);
    if (!height) continue;
    const exceptions = children(row, 'tblPrEx');
    const rowAnchor =
      exceptions.length === 1 ? unchangedRecord(exceptions[0]!, 'tblPrExChange') : undefined;
    if (exceptions.length > 1 || (exceptions.length === 1 && !rowAnchor)) continue;
    const anchor = rowAnchor ?? tableAnchor;
    if (!anchor) continue;
    // Word's cell records hold the cells unchanged. The engine's and a room's record of one
    // change carry its anchor's own `w:id`, and may hold the same change's other cell
    // properties too (a fill pressed in the same suggestion)
    const id = attr(anchor, 'id');
    const cells = children(row, 'tc').map((cell) => {
      const properties = children(cell, 'tcPr');
      if (properties.length !== 1) return undefined;
      const unchanged = unchangedRecord(properties[0]!, 'tcPrChange');
      if (unchanged) return unchanged;
      const records = children(properties[0]!, 'tcPrChange');
      const own = records.length === 1 ? records[0] : undefined;
      return own && id !== undefined && attr(own, 'id') === id && sameActor(own, anchor)
        ? own
        : undefined;
    });
    const first = cells[0];
    if (!first || !cells.every((cell) => cell !== undefined && sameActor(cell, first))) continue;
    if (rowAnchor && !sameActor(rowAnchor, first)) continue;
    found.push({
      row,
      height,
      cells: cells as OoxmlElement[],
      anchor,
      rowAnchor: rowAnchor !== undefined,
    });
  }
  return found;
}

/** Whether the height's anchor is its author's own record (the row's, or the table's). */
export const anchorsOwnHeight = (row: WordHeightRow): boolean =>
  row.rowAnchor || sameActor(row.anchor, row.cells[0]!);

/** The table that holds `node`. */
export function tableOf(part: OoxmlPart, node: OoxmlElement | null): OoxmlElement | undefined {
  let ancestor = node && parentNodeOf(part, node.id);
  while (ancestor && ancestor.kind !== 'table') {
    ancestor = parentNodeOf(part, ancestor.id);
  }
  return ancestor?.kind === 'table' ? ancestor : undefined;
}

/**
 * The heights a reject of `matched` takes away (by `w:trHeight` id, to its row's id): those
 * whose every cell record is matched. The anchor need not be: a table record that another
 * height still leans on stays, or moves onto that row (`revision-table-row-anchor.ts`).
 */
export function implicitTableRowProperties(
  part: OoxmlPart,
  matched: readonly RevisionSite[]
): ReadonlyMap<string, string> {
  const selected = new Set(matched.filter((s) => !s.refused).map((s) => s.node.id));
  const result = new Map(implicitTableRowAlignments(part, matched).rows);
  const tables = new Map<string, OoxmlElement>();
  for (const site of matched) {
    if (site.node.localName !== 'tcPrChange' || !selected.has(site.node.id)) continue;
    const table = tableOf(part, site.parent);
    if (table) tables.set(table.id, table);
  }
  for (const table of tables.values()) {
    for (const height of wordHeightRows(table)) {
      if (!height.cells.every((cell) => selected.has(cell.id))) continue;
      result.set(height.height.id, height.row.id);
    }
  }
  return result;
}
