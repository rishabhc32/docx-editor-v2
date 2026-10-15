// TRACKED column insertion and deletion — `w:cellIns` and `w:cellDel` (§17.13.5.4, §17.13.5.1)
// on every cell of one grid column, the cell twins of a tracked row.
//
// THE NEIGHBOUR HOLDS THE WIDTH. Word resolves a cell revision by giving the cell's grid space
// to the cell before it in the row, or to the first cell after it at the row's start
// (`rebuildRevisionTable`): accepting a deleted cell, or rejecting an inserted one, widens that
// neighbour. So a tracked column edit keeps the table's width, where the untracked one grows or
// shrinks it.
//
//   - Deleting marks the column's cells and nothing else. Reject drops the markers. Accept
//     removes the cells, and the neighbour takes their width.
//   - Inserting takes the new column's width FROM that neighbour — the cell that would take it
//     back — by halving it, in `w:tblGrid` and in each row's `w:tcW`. Reject then restores the
//     table exactly, and Accept keeps the new column in a table of the same width.
//
// Cell text is not wrapped in `w:del`: the marker decides the cell, the rule a tracked row
// deletion follows. Deleting a column this author proposed inserting WITHDRAWS it — the same
// edit as rejecting it, which hands the width back — rather than proposing to delete a proposal.
//
// Refused with `unsupported-revision`, never approximated: another author's pending cell
// insertion or a cell merge in the column, a row another author proposed inserting or deleting,
// another author's pending change inside a column being withdrawn, a neighbour with a
// cell-property record (its recorded width would stop matching), a table with no `w:tblGrid`,
// and a neighbour width Word could not hand back exactly. Merged cells are already refused by
// the untracked column ops (`table-has-merge`).

import {
  applyEdits,
  createNodeIdAllocator,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlPart,
  type OoxmlTableCellNode,
} from '../package/ooxml-tree.ts';
import {
  wmlFreshNamespaceContextAt,
  type WmlFreshNamespaceContext,
} from '../package/wml-namespace.ts';
import { CELL_REVISION_NAMES, revisionAttribute } from './revision-table-plan.ts';
import { paragraphIdsWithin } from './tree-op-blocks.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { resolveRevisions } from './tree-op-revisions.ts';
import { patchTcPrChild, removeTcPrChild } from './tree-op-table-properties.ts';
import { isWmlGridCol, wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';
import { readEditableTableTopology, type EditableTableTopology } from './tree-op-table-topology.ts';
import { holdsOtherAuthorsRevision } from './tree-op-tracked-table-deletion.ts';
import {
  applyInsertTableColumn,
  freshWmlElement,
  gridColWithWidth,
  trackedRevisionAttributes,
} from './tree-op-tables.ts';
import type {
  RevisionAttributionInput,
  TreeDocOp,
  TreeOpEffect,
  TreeOpRejection,
  TreeOpResult,
} from './tree-op-types.ts';

type TableColumnOp = Extract<TreeDocOp, { op: 'insertTableColumn' | 'deleteTableColumn' }>;
type InsertColumnOp = Extract<TableColumnOp, { op: 'insertTableColumn' }>;
type DeleteColumnOp = Extract<TableColumnOp, { op: 'deleteTableColumn' }>;

/** The cell's pending structural revision: `w:cellIns`, `w:cellDel` or `w:cellMerge`. */
function cellRevisionOf(cell: OoxmlElement): OoxmlElement | undefined {
  const properties = wmlChildNamed(cell, 'tcPr');
  if (!properties) return undefined;
  for (const name of CELL_REVISION_NAMES) {
    const marker = wmlChildNamed(properties, name);
    if (marker) return marker;
  }
  return undefined;
}

/** Whether someone other than `author` proposed inserting or deleting the row. */
function isOtherAuthorsRow(row: OoxmlElement, author: string): boolean {
  const properties = wmlChildNamed(row, 'trPr');
  return ['ins', 'del'].some((name) => {
    const marker = properties && wmlChildNamed(properties, name);
    return marker !== undefined && revisionAttribute(marker, 'author') !== author;
  });
}

function wmlAttribute(
  localName: string,
  value: string,
  wml: WmlFreshNamespaceContext
): OoxmlAttribute {
  return {
    kind: 'genericExtension',
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: wml.attributePrefix,
    value,
  };
}

/** The cell with one `w:tcPr` child set in schema order, or removed. */
function withCellProperty(
  cell: OoxmlTableCellNode,
  property: OoxmlElement | { readonly remove: string },
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  const existing = wmlChildNamed(cell, 'tcPr');
  if (!existing && 'remove' in property) return cell;
  const container = existing ?? freshWmlElement('tcPr', nextId, wml, []);
  // Both refuse only an expanded-name mismatch, and every property here is minted in the WML
  // namespace, so the refusal branch is unreachable — as in `withTrackedRowMarker`.
  const patched =
    'remove' in property
      ? removeTcPrChild(container, property.remove)
      : patchTcPrChild(container, property);
  if (!patched.ok || patched.container === existing) return cell;
  const children = existing
    ? cell.children.map((child) => (child === existing ? patched.container : child))
    : [patched.container, ...cell.children];
  return Object.freeze({ ...cell, children }) as OoxmlTableCellNode;
}

function withCellMarker(
  cell: OoxmlTableCellNode,
  kind: 'cellIns' | 'cellDel',
  revisionId: string,
  revision: RevisionAttributionInput,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  const attributes = trackedRevisionAttributes(revisionId, revision, wml);
  return withCellProperty(cell, freshWmlElement(kind, nextId, wml, attributes), nextId, wml);
}

/** The first paragraph of the first-row cell beside the column, where the caret goes. */
function caretBeside(topology: EditableTableTopology, columnIndex: number): string | undefined {
  const cells = topology.rows[0]?.cells ?? [];
  const beside = cells[columnIndex < cells.length - 1 ? columnIndex + 1 : columnIndex - 1];
  return beside?.children.find((child) => child.kind === 'paragraph')?.id;
}

// ── Deletion ──────────────────────────────────────────────────────────────────────────────

interface ColumnDeletionPlan {
  readonly topology: EditableTableTopology;
  readonly columnIndex: number;
  /** Cells that gain `w:cellDel`. */
  readonly marked: readonly OoxmlTableCellNode[];
  /** Cells this author proposed, with their `w:cellIns`: withdrawn by rejecting it. */
  readonly withdrawn: readonly {
    readonly cell: OoxmlTableCellNode;
    readonly marker: OoxmlElement;
  }[];
}

function planColumnDeletion(
  part: OoxmlPart,
  op: DeleteColumnOp,
  author: string
): ColumnDeletionPlan | TreeOpRejection {
  const result = readEditableTableTopology(part.root, op.tableId);
  if (!result.ok) return 'unknown-table';
  const { topology } = result;
  const columnIndex = topology.gridColumns.findIndex((column) => column.id === op.gridColumnId);
  if (columnIndex === -1) return 'unknown-grid-column';
  const marked: OoxmlTableCellNode[] = [];
  const withdrawn: { cell: OoxmlTableCellNode; marker: OoxmlElement }[] = [];
  for (const { row, cells } of topology.rows) {
    // A row another author proposed decides its cells too; a mark in it would hold that decision.
    if (isOtherAuthorsRow(row, author)) return 'unsupported-revision';
    const cell = cells[columnIndex]!;
    const marker = cellRevisionOf(cell);
    if (!marker) marked.push(cell);
    // Already proposed for deletion: that decision stands and needs no second one.
    else if (marker.localName === 'cellDel') continue;
    else if (marker.localName === 'cellIns' && revisionAttribute(marker, 'author') === author) {
      // Withdrawing removes the cell, so it must not take someone else's pending change along.
      if (holdsOtherAuthorsRevision(cell, author)) return 'unsupported-revision';
      withdrawn.push({ cell, marker });
    } else return 'unsupported-revision';
  }
  // A table has to keep a column. Pending deletions cannot keep it alive for another column
  // suggestion — the rule a tracked `deleteTableRow` applies to the last row.
  const survives = topology.gridColumns.some(
    (_, index) =>
      index !== columnIndex &&
      topology.rows.some(({ cells }) => cellRevisionOf(cells[index]!)?.localName !== 'cellDel')
  );
  if (!survives) return 'block-required';
  return { topology, columnIndex, marked, withdrawn };
}

function applyTrackedColumnDeletion(
  part: OoxmlPart,
  op: DeleteColumnOp,
  revision: RevisionAttributionInput,
  options?: EditOptions
): TreeOpResult {
  const plan = planColumnDeletion(part, op, revision.author);
  if (typeof plan === 'string') return { ok: false, reason: plan };
  const { topology, columnIndex } = plan;
  const caretParagraphId = caretBeside(topology, columnIndex);
  if (!caretParagraphId) return { ok: false, reason: 'tree-invariant' };
  const nextId = createNodeIdAllocator(part);
  const wml = wmlFreshNamespaceContextAt(part, topology.table);
  // The transaction's id when it lends one: Delete Columns writes an op per column, one decision.
  const revisionId = options?.revisionIds?.() ?? nextRevisionId(part)();
  const marked = plan.marked.map((cell) =>
    withCellMarker(cell, 'cellDel', revisionId, revision, nextId, wml)
  );
  const effect: TreeOpEffect = {
    dirty: [topology.table.id, ...marked.map((cell) => cell.id)],
    created: [],
    deleted: plan.withdrawn.flatMap(({ cell }) => [cell.id, ...paragraphIdsWithin(cell)]),
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
    caret: { paragraphId: caretParagraphId },
  };
  const edits = marked.map(
    (cell) => (current: OoxmlPart) => replaceNode(current, cell.id, cell, options)
  );
  const proposed = fromEdit(applyEdits(part, edits, options), effect);
  if (!proposed.ok || plan.withdrawn.length === 0) return proposed;
  const withdrawal = resolveRevisions(proposed.part, 'reject', undefined, {
    ...options,
    siteNodeIds: plan.withdrawn.map(({ marker }) => marker.id),
  });
  if (!withdrawal.ok || !withdrawal.part) {
    return { ok: false, reason: withdrawal.reason ?? 'unsupported-revision' };
  }
  return { ok: true, part: withdrawal.part, effect };
}

// ── Insertion ─────────────────────────────────────────────────────────────────────────────

/** A cell's `w:tcW` as the neighbour lends it: split in two, or copied as it is. */
type LentWidth =
  | { readonly split: true; readonly type: 'dxa' | 'pct'; readonly value: number }
  | { readonly split: false; readonly width: OoxmlElement | undefined };

const INTEGRAL_WIDTH = /^\d{1,9}$/;

/**
 * How a neighbour cell lends width, or null when Word could not hand it back exactly.
 *
 * Word adds a removed cell's `w:tcW` to its neighbour's only for `dxa` and `pct` widths of one
 * type. An `auto` or `nil` width is lent as it is; a typeless or non-integral one is refused.
 */
function lentWidthOf(cell: OoxmlElement): LentWidth | null {
  const properties = wmlChildNamed(cell, 'tcPr');
  const width = properties && wmlChildNamed(properties, 'tcW');
  if (!width) return { split: false, width: undefined };
  const type = wmlAttributeValue(width, 'type');
  if (type === 'auto' || type === 'nil') return { split: false, width };
  const raw = wmlAttributeValue(width, 'w') ?? '';
  if ((type !== 'dxa' && type !== 'pct') || !INTEGRAL_WIDTH.test(raw)) return null;
  return { split: true, type, value: Number(raw) };
}

function gridColumnsOf(grid: OoxmlElement): OoxmlElement[] {
  const columns: OoxmlElement[] = [];
  for (const child of grid.children) if (isWmlGridCol(child)) columns.push(child);
  return columns;
}

/** The integral width of one `w:gridCol`, or null. */
function gridWidthAt(grid: OoxmlElement, index: number): number | null {
  const column = gridColumnsOf(grid)[index];
  const raw = column && wmlAttributeValue(column, 'w');
  return raw !== undefined && INTEGRAL_WIDTH.test(raw) ? Number(raw) : null;
}

interface ColumnInsertionPlan {
  /** Grid index of the new column. */
  readonly insertIndex: number;
  /** Grid index, BEFORE the insertion, of the column that lends its width. */
  readonly lenderIndex: number;
  /** Per row, how the lending cell's `w:tcW` splits. */
  readonly lent: readonly LentWidth[];
}

function planColumnInsertion(
  part: OoxmlPart,
  op: InsertColumnOp
): ColumnInsertionPlan | TreeOpRejection {
  const result = readEditableTableTopology(part.root, op.tableId);
  if (!result.ok) return 'unknown-table';
  const { topology } = result;
  // The width is taken from the grid, so a table without one cannot lend it.
  if (!topology.grid || !('gridColumnId' in op)) return 'unsupported-revision';
  const referenceIndex = topology.gridColumns.findIndex((column) => column.id === op.gridColumnId);
  if (referenceIndex === -1) return 'unknown-grid-column';
  const insertIndex = op.where === 'left' ? referenceIndex : referenceIndex + 1;
  const lenderIndex = Math.max(insertIndex - 1, 0);
  if (gridWidthAt(topology.grid, lenderIndex) === null) return 'unsupported-revision';
  // A grid history holds its own copy of the grid, which has to split the same way.
  const history = wmlChildNamed(topology.grid, 'tblGridChange');
  if (history) {
    const recorded = wmlChildNamed(history, 'tblGrid');
    if (
      !recorded ||
      gridColumnsOf(recorded).length !== topology.gridColumns.length ||
      gridWidthAt(recorded, lenderIndex) === null
    )
      return 'unsupported-revision';
  }
  const lent: LentWidth[] = [];
  for (const { cells } of topology.rows) {
    const lender = cells[lenderIndex]!;
    const properties = wmlChildNamed(lender, 'tcPr');
    // A recorded `w:tcW` would stop matching the width the cell lends.
    if (properties && wmlChildNamed(properties, 'tcPrChange')) return 'unsupported-revision';
    const width = lentWidthOf(lender);
    if (!width) return 'unsupported-revision';
    lent.push(width);
  }
  return { insertIndex, lenderIndex, lent };
}

/** The grid with the column at `index` set to `width`, every other entry kept. */
function withGridWidth(
  grid: OoxmlElement,
  index: number,
  width: number,
  wml: WmlFreshNamespaceContext
): OoxmlElement {
  const column = gridColumnsOf(grid)[index]!;
  const children = grid.children.map((child) =>
    child === column ? gridColWithWidth(column, width, wml) : child
  );
  return Object.freeze({ ...grid, children }) as OoxmlElement;
}

/**
 * `w:tblGrid` with the lender halved and the new column holding the other half.
 *
 * `grid` is the grid AFTER the untracked insertion. Its `w:tblGridChange` still holds the grid
 * from before it, so the record gains the same column, split from the same lender.
 */
function lentGrid(
  grid: OoxmlElement,
  plan: ColumnInsertionPlan,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlElement {
  const { insertIndex, lenderIndex } = plan;
  const lenderAfter = insertIndex > 0 ? lenderIndex : 1;
  const width = gridWidthAt(grid, lenderAfter)!;
  const half = Math.floor(width / 2);
  const live = withGridWidth(
    withGridWidth(grid, lenderAfter, width - half, wml),
    insertIndex,
    half,
    wml
  );
  const history = wmlChildNamed(live, 'tblGridChange');
  const recorded = history && wmlChildNamed(history, 'tblGrid');
  if (!history || !recorded) return live;
  const recordedWidth = gridWidthAt(recorded, lenderIndex)!;
  const recordedHalf = Math.floor(recordedWidth / 2);
  const halved = withGridWidth(recorded, lenderIndex, recordedWidth - recordedHalf, wml);
  const columns = gridColumnsOf(halved);
  const before = columns[insertIndex];
  const children = [...halved.children];
  children.splice(
    before ? children.indexOf(before) : children.indexOf(columns[columns.length - 1]!) + 1,
    0,
    freshWmlElement('gridCol', nextId, wml, [wmlAttribute('w', String(recordedHalf), wml)])
  );
  const record = Object.freeze({ ...halved, children }) as OoxmlElement;
  const updated = Object.freeze({
    ...history,
    children: history.children.map((child) => (child === recorded ? record : child)),
  }) as OoxmlElement;
  return Object.freeze({
    ...live,
    children: live.children.map((child) => (child === history ? updated : child)),
  }) as OoxmlElement;
}

function widthElement(
  type: string,
  value: number,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlElement {
  return freshWmlElement('tcW', nextId, wml, [
    wmlAttribute('w', String(value), wml),
    wmlAttribute('type', type, wml),
  ]);
}

function applyTrackedColumnInsertion(
  part: OoxmlPart,
  op: InsertColumnOp,
  revision: RevisionAttributionInput,
  options?: EditOptions
): TreeOpResult {
  const plan = planColumnInsertion(part, op);
  if (typeof plan === 'string') return { ok: false, reason: plan };
  // ONE id for every new cell, as a tracked `insertTable` writes: one gesture is one decision.
  const revisionId = options?.revisionIds?.() ?? nextRevisionId(part)();
  // The untracked op builds the column — cells, paragraphs, grid entry — exactly as editing
  // mode does. This lends it its width and proposes it.
  const inserted = applyInsertTableColumn(part, op, options);
  if (!inserted.ok) return inserted;
  const after = readEditableTableTopology(inserted.part.root, op.tableId);
  const grid = after.ok ? after.topology.grid : undefined;
  if (!after.ok || !grid) return { ok: false, reason: 'tree-invariant' };
  const { topology } = after;
  const { insertIndex, lenderIndex } = plan;
  const lenderAfter = insertIndex > 0 ? lenderIndex : 1;
  const nextId = createNodeIdAllocator(inserted.part);
  const wml = wmlFreshNamespaceContextAt(inserted.part, topology.table);
  const replaced: OoxmlElement[] = [lentGrid(grid, plan, nextId, wml)];
  for (const [rowIndex, { cells }] of topology.rows.entries()) {
    const lent = plan.lent[rowIndex]!;
    let lender = cells[lenderAfter]!;
    let cell = cells[insertIndex]!;
    if (lent.split) {
      const half = Math.floor(lent.value / 2);
      const kept = widthElement(lent.type, lent.value - half, nextId, wml);
      lender = withCellProperty(lender, kept, nextId, wml);
      cell = withCellProperty(cell, widthElement(lent.type, half, nextId, wml), nextId, wml);
    } else {
      // The new cell takes the lender's own `auto` or `nil` width, or none when it has none.
      const width = lent.width && ({ ...lent.width, id: nextId() } as OoxmlElement);
      cell = withCellProperty(cell, width ?? { remove: 'tcW' }, nextId, wml);
    }
    cell = withCellMarker(cell, 'cellIns', revisionId, revision, nextId, wml);
    replaced.push(lender, cell);
  }
  const edits = replaced.map(
    (node) => (current: OoxmlPart) => replaceNode(current, node.id, node, options)
  );
  const lenders = topology.rows.map(({ cells }) => cells[lenderAfter]!.id);
  const effect: TreeOpEffect = {
    ...inserted.effect,
    dirty: [...new Set([...inserted.effect.dirty, grid.id, ...lenders])],
  };
  return fromEdit(applyEdits(inserted.part, edits, options), effect);
}

// ── Dispatch ──────────────────────────────────────────────────────────────────────────────

/**
 * Why a proposed column op is refused, or null when it fits the tracked forms above. Runs after
 * the op's untracked form passed.
 */
export function validateTrackedColumnEdit(
  part: OoxmlPart,
  op: TableColumnOp,
  revision: RevisionAttributionInput
): TreeOpRejection | null {
  const plan =
    op.op === 'deleteTableColumn'
      ? planColumnDeletion(part, op, revision.author)
      : planColumnInsertion(part, op);
  return typeof plan === 'string' ? plan : null;
}

/** Propose a column op. Runs after `validateTrackedColumnEdit` admitted it. */
export function applyTrackedColumnEdit(
  part: OoxmlPart,
  op: TableColumnOp,
  revision: RevisionAttributionInput,
  options?: EditOptions
): TreeOpResult {
  return op.op === 'deleteTableColumn'
    ? applyTrackedColumnDeletion(part, op, revision, options)
    : applyTrackedColumnInsertion(part, op, revision, options);
}
