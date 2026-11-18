// Column insertion and deletion in a table with merged cells (`w:gridSpan`, `w:vMerge`), as
// Word does them. Word for Mac 16, scripted on 2026-10-29 (the samples are listed in the ordered
// collaboration design, D17).
//
// A table with no merge keeps the column ops of `tree-op-tables.ts`, which put one cell in every
// row at one grid position: both rules give the same table there. `w:hMerge` (the legacy
// horizontal merge) is refused, as before.

import {
  applyEdits,
  createNodeIdAllocator,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import { usedParaIds } from '../package/para-id.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
  type OoxmlTableCellNode,
  type OoxmlTableRowNode,
} from '../package/ooxml-tree.ts';
import {
  wmlFreshNamespaceContextAt,
  type WmlFreshNamespaceContext,
} from '../package/wml-namespace.ts';
import { paragraphIdsWithin } from './tree-op-blocks.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import {
  patchTcPrChild,
  patchTrPrChild,
  removeTcPrChild,
  removeTrPrChild,
} from './tree-op-table-properties.ts';
import { isWmlGridCol, wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';
import { readEditableTableTopology, type EditableTableTopology } from './tree-op-table-topology.ts';
import { buildFreshColumnCell, freshWmlElement } from './tree-op-tables.ts';
import type { TreeDocOp, TreeOpEffect, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type InsertColumnOp = Extract<TreeDocOp, { op: 'insertTableColumn' }>;
type DeleteColumnOp = Extract<TreeDocOp, { op: 'deleteTableColumn' }>;

const INTEGRAL = /^\d{1,9}$/;

/** A non-negative whole number attribute of a WML element, or null. */
export function wholeAttribute(node: OoxmlElement | undefined, name: string): number | null {
  const raw = node && wmlAttributeValue(node, name);
  return raw !== undefined && INTEGRAL.test(raw) ? Number(raw) : null;
}

function cellProperty(cell: OoxmlTableCellNode, name: string): OoxmlElement | undefined {
  const properties = wmlChildNamed(cell, 'tcPr');
  return properties && wmlChildNamed(properties, name);
}

/** A cell's span over grid columns (`w:gridSpan`), 1 when absent; 0 when unreadable. */
function spanOf(cell: OoxmlTableCellNode): number {
  const span = cellProperty(cell, 'gridSpan');
  if (!span) return 1;
  return wholeAttribute(span, 'val') ?? 0;
}

/** The grid columns a row leaves empty before or after its cells. */
function rowSkip(row: OoxmlTableRowNode, name: 'gridBefore' | 'gridAfter'): number {
  const properties = wmlChildNamed(row, 'trPr');
  const skip = properties && wmlChildNamed(properties, name);
  return skip ? (wholeAttribute(skip, 'val') ?? 0) : 0;
}

/** Whether the table holds a merge: horizontal (`w:gridSpan`, `w:hMerge`) or vertical. */
export function tableHasColumnMergeMarkers(topology: EditableTableTopology): boolean {
  return topology.rows.some(({ cells }) =>
    cells.some((cell) => ['gridSpan', 'hMerge', 'vMerge'].some((n) => cellProperty(cell, n)))
  );
}

/** Why a table without merges cannot take the regular column ops, or null. */
export function validateRegularColumnTable(
  topology: EditableTableTopology
): TreeOpRejection | null {
  if (tableHasColumnMergeMarkers(topology)) return 'table-has-merge';
  if (topology.rows.length === 0) return 'tree-invariant';
  const cellCount = topology.rows[0]!.cells.length;
  if (cellCount === 0) return 'tree-invariant';
  for (const { row, cells } of topology.rows) {
    if (cells.length !== cellCount) return 'tree-invariant';
    if (rowSkip(row, 'gridBefore') > 0 || rowSkip(row, 'gridAfter') > 0) return 'tree-invariant';
  }

  if (topology.grid && topology.gridColumns.length !== cellCount) return 'tree-invariant';
  return null;
}

/** A row's cells placed on the grid: the columns skipped first, and each cell's span. */
interface RowLayout {
  readonly row: OoxmlTableRowNode;
  readonly before: number;
  readonly cells: readonly { readonly cell: OoxmlTableCellNode; readonly start: number }[];
}

interface Layout {
  /** Each grid column's width in twips. */
  readonly widths: readonly number[];
  readonly rows: readonly RowLayout[];
}

/** The table on its grid, or why it cannot be read as one Word would edit. */
function layoutOf(topology: EditableTableTopology): Layout | TreeOpRejection {
  if (!topology.grid || topology.gridColumns.length === 0) return 'unknown-grid-column';
  if (topology.rows.length === 0) return 'tree-invariant';
  const widths: number[] = [];
  for (const column of topology.gridColumns) {
    const width = wholeAttribute(column, 'w');

    // A column with no width would vanish from the grid that is built again
    if (width === null || width === 0) return 'tree-invariant';
    widths.push(width);
  }

  const rows: RowLayout[] = [];
  for (const { row, cells } of topology.rows) {
    if (cells.some((cell) => cellProperty(cell, 'hMerge'))) return 'table-has-merge';
    const before = rowSkip(row, 'gridBefore');
    let at = before;
    const placed: { cell: OoxmlTableCellNode; start: number }[] = [];
    for (const cell of cells) {
      const span = spanOf(cell);
      if (span < 1) return 'tree-invariant';
      placed.push({ cell, start: at });
      at += span;
    }

    if (cells.length === 0 || at + rowSkip(row, 'gridAfter') > widths.length) {
      return 'tree-invariant';
    }

    rows.push({ row, before, cells: placed });
  }

  return { widths, rows };
}

/** The width in twips of grid columns `[from, from + count)`. */
const widthOf = (widths: readonly number[], from: number, count: number): number =>
  widths.slice(from, from + count).reduce((sum, width) => sum + width, 0);

/** The planned cells of a row as it stands: each cell with the width its span covers. */
function plannedCells(
  layout: Layout,
  row: RowLayout
): { cell: OoxmlTableCellNode | null; width: number }[] {
  return row.cells.map(({ cell, start }) => ({
    cell,
    width: widthOf(layout.widths, start, spanOf(cell)),
  }));
}

/** The index of the row's cell that covers a grid column, or -1. */
function coveringIndex(row: RowLayout, column: number): number {
  return row.cells.findIndex(({ cell, start }) => start <= column && column < start + spanOf(cell));
}

/** A row as cells in order, each with its width in twips; a new cell has no node yet. */
interface PlannedRow {
  readonly row: OoxmlTableRowNode;
  readonly beforeWidth: number;
  readonly cells: readonly {
    readonly cell: OoxmlTableCellNode | null;
    readonly width: number;
  }[];
}

/** The grid built again from the rows: column widths, and each row's placement on it. */
interface Rebuilt {
  readonly widths: readonly number[];
  readonly rows: readonly {
    readonly before: number;
    readonly spans: readonly number[];
    readonly after: number;
  }[];
}

function rebuiltGrid(rows: readonly PlannedRow[]): Rebuilt {
  const edges = new Set<number>([0]);
  const ends: number[] = [];
  for (const { beforeWidth, cells } of rows) {
    let x = beforeWidth;
    edges.add(x);
    for (const { width } of cells) {
      x += width;
      edges.add(x);
    }

    ends.push(x);
  }

  const sorted = [...edges].sort((a, b) => a - b);
  const index = new Map(sorted.map((x, k) => [x, k]));
  const last = sorted.length - 1;
  return {
    widths: sorted.slice(1).map((x, k) => x - sorted[k]!),
    rows: rows.map(({ beforeWidth, cells }, r) => {
      let x = beforeWidth;
      const spans = cells.map(({ width }) => {
        const from = index.get(x)!;
        x += width;
        return index.get(x)! - from;
      });
      return { before: index.get(beforeWidth)!, spans, after: last - index.get(ends[r]!)! };
    }),
  };
}

export function wmlAttribute(
  name: string,
  value: string,
  wml: WmlFreshNamespaceContext
): OoxmlAttribute {
  return {
    kind: 'genericExtension',
    namespaceUri: WML_NAMESPACE_URI,
    localName: name,
    prefix: wml.attributePrefix,
    value,
  };
}

/** The `w:w` and `w:type` attributes of a width (`CT_TblWidth`). */
function widthAttributes(
  type: string,
  value: number,
  wml: WmlFreshNamespaceContext
): OoxmlAttribute[] {
  return [wmlAttribute('w', String(value), wml), wmlAttribute('type', type, wml)];
}

/** A cell's `w:tcW` of `type`. */
export function widthElement(
  type: string,
  value: number,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlElement {
  return freshWmlElement('tcW', nextId, wml, widthAttributes(type, value, wml));
}

/** The cell with `properties` as its `w:tcPr`: in place of the one it has, or first (`CT_Tc`). */
function cellWithProperties(
  cell: OoxmlTableCellNode,
  properties: OoxmlElement
): OoxmlTableCellNode {
  const existing = wmlChildNamed(cell, 'tcPr');
  const children = existing
    ? cell.children.map((child) => (child === existing ? properties : child))
    : [properties, ...cell.children];
  return Object.freeze({ ...cell, children }) as OoxmlTableCellNode;
}

/** The cell with one `w:tcPr` child set in schema order, or removed. */
export function withCellProperty(
  cell: OoxmlTableCellNode,
  property: OoxmlElement | { readonly remove: string },
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  const existing = wmlChildNamed(cell, 'tcPr');
  if (!existing && 'remove' in property) return cell;
  const container = existing ?? freshWmlElement('tcPr', nextId, wml, []);

  // Both refuse only an expanded-name mismatch, and every property here is minted in the WML
  // namespace, so the refusal branch is unreachable, as in `withTrackedRowMarker`.
  const patched =
    'remove' in property
      ? removeTcPrChild(container, property.remove)
      : patchTcPrChild(container, property);
  if (!patched.ok || patched.container === existing) return cell;
  return cellWithProperties(cell, patched.container);
}

/** A cell with its `w:gridSpan` set to `span` (removed at 1) and, given, its `w:tcW`. */
function placedCell(
  cell: OoxmlTableCellNode,
  span: number,
  width: number | null,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  if (span === spanOf(cell) && width === null) return cell;
  let properties = wmlChildNamed(cell, 'tcPr') ?? freshWmlElement('tcPr', nextId, wml, []);
  const patch = (result: { ok: boolean; container?: OoxmlElement }) => {
    if (result.ok && result.container) properties = result.container;
  };
  patch(
    span > 1
      ? patchTcPrChild(
          properties,
          freshWmlElement('gridSpan', nextId, wml, [wmlAttribute('val', String(span), wml)])
        )
      : removeTcPrChild(properties, 'gridSpan')
  );
  if (width !== null) {
    patch(patchTcPrChild(properties, widthElement('dxa', width, nextId, wml)));
  }

  return cellWithProperties(cell, properties);
}

/** A row with its skipped columns (`w:gridBefore`, `w:gridAfter` and their widths) and cells. */
function placedRow(
  row: OoxmlTableRowNode,
  skips: { readonly before: number; readonly after: number },
  widths: { readonly before: number; readonly after: number },
  cells: readonly OoxmlTableCellNode[],
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableRowNode {
  const existing = wmlChildNamed(row, 'trPr');
  let properties = existing ?? freshWmlElement('trPr', nextId, wml, []);
  const set = (name: string, attributes: OoxmlAttribute[] | null) => {
    const result = attributes
      ? patchTrPrChild(properties, freshWmlElement(name, nextId, wml, attributes))
      : removeTrPrChild(properties, name);
    if (result.ok && result.container) properties = result.container;
  };
  set('gridBefore', skips.before > 0 ? [wmlAttribute('val', String(skips.before), wml)] : null);
  set('gridAfter', skips.after > 0 ? [wmlAttribute('val', String(skips.after), wml)] : null);
  set('wBefore', skips.before > 0 ? widthAttributes('dxa', widths.before, wml) : null);
  set('wAfter', skips.after > 0 ? widthAttributes('dxa', widths.after, wml) : null);
  const children: OoxmlNode[] = [];
  let cellsPlaced = false;
  for (const child of row.children) {
    if (child === existing) {
      children.push(properties);
    } else if (child.kind !== 'tableCell') {
      children.push(child);
    } else if (!cellsPlaced) {
      children.push(...cells);
      cellsPlaced = true;
    }
  }

  if (!cellsPlaced) children.push(...cells);

  // New row properties lead the row, after `w:tblPrEx`
  if (!existing && properties.children.length > 0) {
    const at = children.findIndex((c) => c.kind !== 'textValue' && c.localName === 'tblPrEx');
    children.splice(at + 1, 0, properties);
  }

  return Object.freeze({ ...row, children }) as OoxmlTableRowNode;
}

/** The table with its grid built again and its rows replaced. */
function tableWith(
  topology: EditableTableTopology,
  rebuilt: Rebuilt,
  rows: readonly (OoxmlTableRowNode | null)[],
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlElement {
  const grid = topology.grid!;

  // A column at the same place with the same width keeps its node (and its id): the columns
  // left of the change, which a later op of the same command may still name
  const kept = new Map<string, OoxmlElement>();
  let x = 0;
  for (const column of topology.gridColumns) {
    const width = wholeAttribute(column, 'w') ?? 0;
    kept.set(`${x}:${width}`, column);
    x += width;
  }

  let at = 0;
  const columns = rebuilt.widths.map((width) => {
    const same = kept.get(`${at}:${width}`);
    at += width;
    return same ?? freshWmlElement('gridCol', nextId, wml, [wmlAttribute('w', String(width), wml)]);
  });
  const newGrid = Object.freeze({
    ...grid,
    children: [...columns, ...grid.children.filter((child) => !isWmlGridCol(child))],
  }) as OoxmlElement;
  const byId = new Map(topology.rows.map(({ row }, k) => [row.id, rows[k]]));
  const children: OoxmlNode[] = [];
  for (const child of topology.table.children) {
    if (child === grid) {
      children.push(newGrid);
    } else if (byId.has(child.id)) {
      const row = byId.get(child.id);
      if (row) children.push(row);
    } else {
      children.push(child);
    }
  }

  return Object.freeze({ ...topology.table, children }) as OoxmlElement;
}

/**
 * The rows planned, built on a new grid: each cell's span, each row's skipped columns. A planned
 * cell with no node of its own takes the node `nodeOf` gives it.
 */
function rowsOnGrid(
  planned: readonly PlannedRow[],
  nodeOf: ReadonlyMap<PlannedRow['cells'][number], OoxmlTableCellNode>,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): { readonly rebuilt: Rebuilt; readonly rows: OoxmlTableRowNode[] } {
  const rebuilt = rebuiltGrid(planned);
  const rows = planned.map((plan, r) => {
    const placement = rebuilt.rows[r]!;
    const cells = plan.cells.map((entry, k) => {
      const span = placement.spans[k]!;
      const made = nodeOf.get(entry);
      return made
        ? placedCell(made, span, entry.width, nextId, wml)
        : placedCell(entry.cell!, span, null, nextId, wml);
    });
    const end = plan.cells.reduce((x, { width }) => x + width, plan.beforeWidth);
    const total = rebuilt.widths.reduce((sum, width) => sum + width, 0);
    return placedRow(
      plan.row,
      { before: placement.before, after: placement.after },
      { before: plan.beforeWidth, after: total - end },
      cells,
      nextId,
      wml
    );
  });
  return { rebuilt, rows };
}

/** Why a column op on a table with merged cells is refused, or null. */
export function validateMergedColumnOp(
  topology: EditableTableTopology,
  op: InsertColumnOp | DeleteColumnOp
): TreeOpRejection | null {
  const layout = layoutOf(topology);
  if (typeof layout === 'string') return layout;
  if (!('gridColumnId' in op) || typeof op.gridColumnId !== 'string') return 'invalidArgs';
  const column = topology.gridColumns.findIndex((c) => c.id === op.gridColumnId);
  if (column === -1) return 'unknown-grid-column';
  if (op.op === 'insertTableColumn') {
    if (op.where !== 'left' && op.where !== 'right') return 'invalidArgs';

    // Word reads the column from the first row
    return coveringIndex(layout.rows[0]!, column) === -1 ? 'unknown-grid-column' : null;
  }

  // A table keeps a row: deleting the column from every row is deleting the table
  const survives = layout.rows.some(
    (row) => row.cells.length > 1 || coveringIndex(row, column) === -1
  );
  return survives ? null : 'block-required';
}

/**
 * Insert a column as Word does in a table with merged cells. Word takes the FIRST ROW's cell
 * that covers the chosen grid column and adds one cell beside that index in every row (right:
 * the index after it; left: the index itself). So a row whose cell spans the column gets its
 * new cell after the spanning cell, not inside it. The new cell has the column's width and no
 * merge of its own. The grid is built again from the rows, and a row that ends short gets
 * `w:gridAfter`.
 */
export function applyMergedColumnInsertion(
  part: OoxmlPart,
  op: InsertColumnOp,
  topology: EditableTableTopology,
  options?: EditOptions
): TreeOpResult {
  const rejection = validateMergedColumnOp(topology, op);
  if (rejection) return { ok: false, reason: rejection };
  const layout = layoutOf(topology) as Layout;
  const column = topology.gridColumns.findIndex(
    (c) => 'gridColumnId' in op && c.id === op.gridColumnId
  );
  const first = coveringIndex(layout.rows[0]!, column);
  const at = op.where === 'right' ? first + 1 : first;
  const width = layout.widths[column]!;
  const wml = wmlFreshNamespaceContextAt(part, topology.table);
  const nextId = createNodeIdAllocator(part);
  const used = new Set(usedParaIds(part.root));
  const fresh = new Map<PlannedRow['cells'][number], OoxmlTableCellNode>();
  const planned: PlannedRow[] = layout.rows.map((row, r) => {
    const cells = plannedCells(layout, row);

    // The new cell copies the properties of the cell it stands beside, its merges left out
    const index = Math.min(at, cells.length);
    const source = row.cells[Math.max(0, Math.min(index - 1, row.cells.length - 1))]!.cell;
    const entry = { cell: null, width };
    fresh.set(
      entry,
      buildFreshColumnCell(
        source,
        part,
        topology.table,
        nextId,
        `${topology.table.id}:merged-col:${column}:r${r}`,
        used,
        wml
      )
    );
    cells.splice(index, 0, entry);
    return { row: row.row, beforeWidth: widthOf(layout.widths, 0, row.before), cells };
  });
  const { rebuilt, rows } = rowsOnGrid(planned, fresh, nextId, wml);
  const table = tableWith(topology, rebuilt, rows, nextId, wml);
  const madeCells = [...fresh.values()];
  const paragraphs = madeCells.flatMap((cell) => paragraphIdsWithin(cell));
  const effect: TreeOpEffect = {
    dirty: [topology.table.id],
    created: [...madeCells.map((cell) => cell.id), ...paragraphs],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
    ...(paragraphs[0] ? { caret: { paragraphId: paragraphs[0] } } : {}),
  };
  return withTable(part, table, effect, options);
}

/** `part` with `table` in place of the table of its id. */
function withTable(
  part: OoxmlPart,
  table: OoxmlElement,
  effect: TreeOpEffect,
  options: EditOptions | undefined
): TreeOpResult {
  return fromEdit(
    applyEdits(part, [(current) => replaceNode(current, table.id, table, options)], options),
    effect
  );
}

/**
 * Delete a column as Word does in a table with merged cells. In each row the cell that covers
 * the grid column goes: a spanning cell goes whole, with its text. A row with no cell left goes
 * too. The cells after it move left, and a row that ends short gets `w:gridAfter`.
 */
export function applyMergedColumnDeletion(
  part: OoxmlPart,
  op: DeleteColumnOp,
  topology: EditableTableTopology,
  options?: EditOptions
): TreeOpResult {
  const rejection = validateMergedColumnOp(topology, op);
  if (rejection) return { ok: false, reason: rejection };
  const layout = layoutOf(topology) as Layout;
  const column = topology.gridColumns.findIndex((c) => c.id === op.gridColumnId);
  const removedCells: OoxmlTableCellNode[] = [];
  const removedRows: OoxmlTableRowNode[] = [];
  const planned: PlannedRow[] = [];
  for (const row of layout.rows) {
    const covering = coveringIndex(row, column);
    const cells = plannedCells(layout, row).filter((_, k) => k !== covering);
    if (covering !== -1) removedCells.push(row.cells[covering]!.cell);
    if (cells.length === 0) {
      removedRows.push(row.row);
      continue;
    }

    // A column the row skips before its cells is no longer there either
    const skipped = covering === -1 && column < row.before ? layout.widths[column]! : 0;
    planned.push({
      row: row.row,
      beforeWidth: widthOf(layout.widths, 0, row.before) - skipped,
      cells,
    });
  }

  const wml = wmlFreshNamespaceContextAt(part, topology.table);
  const nextId = createNodeIdAllocator(part);
  const { rebuilt, rows } = rowsOnGrid(planned, new Map(), nextId, wml);
  const kept = new Map(planned.map((plan, k) => [plan.row.id, rows[k]!]));
  const table = tableWith(
    topology,
    rebuilt,
    topology.rows.map(({ row }) => kept.get(row.id) ?? null),
    nextId,
    wml
  );
  const firstRow = rows[0];
  const caret = firstRow?.children
    .find((child) => child.kind === 'tableCell')
    ?.children.find((child) => child.kind === 'paragraph')?.id;
  const effect: TreeOpEffect = {
    dirty: [topology.table.id],
    created: [],
    deleted: [
      op.gridColumnId,
      ...removedRows.map((row) => row.id),
      ...removedCells.flatMap((cell) => [cell.id, ...paragraphIdsWithin(cell)]),
    ],
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
    ...(caret ? { caret: { paragraphId: caret } } : {}),
  };
  return withTable(part, table, effect, options);
}

/** Per row, the cell that covers a grid column, or null for a row that skips it. */
export function cellsCoveringColumn(
  topology: EditableTableTopology,
  gridColumnId: string
): (OoxmlTableCellNode | null)[] | null {
  const layout = layoutOf(topology);
  const column = topology.gridColumns.findIndex((c) => c.id === gridColumnId);
  if (typeof layout === 'string' || column === -1) return null;
  return layout.rows.map((row) => row.cells[coveringIndex(row, column)]?.cell ?? null);
}

/**
 * The row of a proposed column with its planned cells: the new cell, marked, holding half of
 * its lender's width, and the lender keeping the rest. Both cells go into `nodeOf`. Refused,
 * not approximated: a row with no new cell or no lender, and a lender whose width is not whole
 * twips or that carries a property record.
 */
function lentRow(
  layout: Layout,
  row: RowLayout,
  created: ReadonlySet<string>,
  marker: (nextId: () => string, wml: WmlFreshNamespaceContext) => OoxmlElement,
  nodeOf: Map<PlannedRow['cells'][number], OoxmlTableCellNode>,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): PlannedRow | TreeOpRejection {
  const cells = plannedCells(layout, row);
  const at = row.cells.findIndex(({ cell }) => created.has(cell.id));
  const lenderAt = at > 0 ? at - 1 : 1;
  const lender = cells[lenderAt];
  const lenderWidth = lender?.cell && cellProperty(lender.cell, 'tcW');
  if (
    at === -1 ||
    !lender?.cell ||
    cellProperty(lender.cell, 'tcPrChange') ||
    (lenderWidth && wmlAttributeValue(lenderWidth, 'type') !== 'dxa')
  ) {
    return 'unsupported-revision';
  }

  const half = Math.floor(lender.width / 2);
  const kept = { cell: null, width: lender.width - half };
  cells[lenderAt] = kept;
  nodeOf.set(kept, lender.cell);
  const made = cells[at]!.cell!;
  const entry = { cell: null, width: half };
  cells[at] = entry;
  nodeOf.set(entry, withCellProperty(made, marker(nextId, wml), nextId, wml));
  return { row: row.row, beforeWidth: widthOf(layout.widths, 0, row.before), cells };
}

/**
 * A proposed column in a table with merged cells: Word's cells (`applyMergedColumnInsertion`),
 * each marked `w:cellIns` and holding width its row's neighbour lends it, half of that cell. The
 * neighbour is the cell a reject gives the space back to (`rebuildRevisionTable`: the cell
 * before it, or at the row's start the cell after it), so a reject restores the table exactly
 * and an accept keeps it as wide as before. Refused, not approximated: a grid history, and the
 * rows `lentRow` refuses.
 */
export function applyTrackedMergedColumnInsertion(
  part: OoxmlPart,
  op: InsertColumnOp,
  topology: EditableTableTopology,
  marker: (nextId: () => string, wml: WmlFreshNamespaceContext) => OoxmlElement,
  options?: EditOptions
): TreeOpResult {
  if (topology.grid && wmlChildNamed(topology.grid, 'tblGridChange')) {
    return { ok: false, reason: 'unsupported-revision' };
  }

  const inserted = applyMergedColumnInsertion(part, op, topology, options);
  if (!inserted.ok) return inserted;
  const read = readEditableTableTopology(inserted.part.root, topology.table.id);
  if (!read.ok) return { ok: false, reason: 'tree-invariant' };
  const after = read.topology;
  const layout = layoutOf(after);
  if (typeof layout === 'string') return { ok: false, reason: layout };
  const created = new Set(inserted.effect.created);
  const planned: PlannedRow[] = [];

  // The new cells and their lenders, which take the widths planned for them
  const marked = new Map<PlannedRow['cells'][number], OoxmlTableCellNode>();
  const wml = wmlFreshNamespaceContextAt(inserted.part, after.table);
  const nextId = createNodeIdAllocator(inserted.part);
  for (const row of layout.rows) {
    const lent = lentRow(layout, row, created, marker, marked, nextId, wml);
    if (typeof lent === 'string') return { ok: false, reason: lent };
    planned.push(lent);
  }

  const { rebuilt, rows } = rowsOnGrid(planned, marked, nextId, wml);
  const table = tableWith(after, rebuilt, rows, nextId, wml);
  return withTable(inserted.part, table, inserted.effect, options);
}
