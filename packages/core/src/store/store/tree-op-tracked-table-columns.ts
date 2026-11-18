// TRACKED column insertion and deletion: `w:cellIns` and `w:cellDel` (§17.13.5.4, §17.13.5.1)
// on every cell of one grid column, the cell twins of a tracked row.
//
// THE NEIGHBOUR HOLDS THE WIDTH. Word resolves a cell revision by giving the cell's grid space
// to the cell before it in the row, or to the first cell after it at the row's start
// (`rebuildRevisionTable`): accepting a deleted cell, or rejecting an inserted one, widens that
// neighbour. So a tracked column edit keeps the table's width, where the untracked one grows or
// shrinks it. A table with merged cells takes Word's merged forms
// (`applyTrackedMergedColumnInsertion`, and the covering cells in `planColumnDeletion`).

import {
  applyEdits,
  createNodeIdAllocator,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlPart, OoxmlTableCellNode } from '../package/ooxml-tree.ts';
import {
  wmlFreshNamespaceContextAt,
  type WmlFreshNamespaceContext,
} from '../package/wml-namespace.ts';
import { CELL_REVISION_NAMES, revisionAttribute } from './revision-table-plan.ts';
import { paragraphIdsWithin } from './tree-op-blocks.ts';
import { cloneWithNewIds, fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { patchTcPrChild } from './tree-op-table-properties.ts';
import {
  isWmlElement,
  isWmlGridCol,
  wmlAttributeValue,
  wmlChildNamed,
} from './tree-op-table-shared.ts';
import { readEditableTableTopology, type EditableTableTopology } from './tree-op-table-topology.ts';
import {
  applyTrackedMergedColumnInsertion,
  cellsCoveringColumn,
  tableHasColumnMergeMarkers,
  wholeAttribute,
  widthElement,
  withCellProperty,
  wmlAttribute,
} from './tree-op-table-merged-columns.ts';
import {
  afterWithdrawal,
  holdsOtherAuthorsRevision,
  rowMarker,
} from './tree-op-tracked-table-deletion.ts';
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
  return (['ins', 'del'] as const).some((name) => {
    const marker = rowMarker(row, name);
    return marker !== undefined && revisionAttribute(marker, 'author') !== author;
  });
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

/**
 * The first paragraph of the first-row cell beside the column, where the caret goes; with merged
 * cells, where that cell is part of the deletion (`taken`), the first cell outside it.
 */
function caretBeside(
  topology: EditableTableTopology,
  columnIndex: number,
  taken: ReadonlySet<string> = new Set()
): string | undefined {
  const cells = topology.rows[0]?.cells ?? [];
  const beside = cells[columnIndex < cells.length - 1 ? columnIndex + 1 : columnIndex - 1];
  const outside = [beside, ...topology.rows.flatMap((entry) => entry.cells)].find(
    (cell) => cell !== undefined && !taken.has(cell.id)
  );
  return outside?.children.find((child) => child.kind === 'paragraph')?.id;
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

/** What a column deletion does with one cell of the column. */
type CellDeletion = 'mark' | 'keep' | 'withdraw' | 'refuse';

/**
 * A cell with no revision is marked. Cell text is not wrapped in `w:del`: the marker decides
 * the cell, the rule a tracked row deletion follows. A cell this author proposed inserting is
 * WITHDRAWN: the same edit as rejecting it, which hands the width back.
 */
function cellDeletionOf(
  cell: OoxmlTableCellNode,
  marker: OoxmlElement | undefined,
  author: string
): CellDeletion {
  if (!marker) {
    return 'mark';
  }

  // Already proposed for deletion: that decision stands and needs no second one.
  if (marker.localName === 'cellDel') {
    return 'keep';
  }

  if (marker.localName === 'cellIns' && revisionAttribute(marker, 'author') === author) {
    // Withdrawing removes the cell, so it must not take someone else's pending change along.
    return holdsOtherAuthorsRevision(cell, author) ? 'refuse' : 'withdraw';
  }

  return 'refuse';
}

/**
 * Whether the table keeps a column. Pending deletions cannot keep it alive for another column
 * suggestion: the rule a tracked `deleteTableRow` applies to the last row.
 */
function keepsAColumn(
  topology: EditableTableTopology,
  columnIndex: number,
  covering: readonly (OoxmlTableCellNode | null)[] | null
): boolean {
  const deleted = (cell: OoxmlTableCellNode) => cellRevisionOf(cell)?.localName === 'cellDel';
  if (covering) {
    return topology.rows.some(({ cells }, rowIndex) =>
      cells.some((cell) => cell !== covering[rowIndex] && !deleted(cell))
    );
  }

  return topology.gridColumns.some(
    (_, index) =>
      index !== columnIndex && topology.rows.some(({ cells }) => !deleted(cells[index]!))
  );
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

  // With merged cells, the cell of each row that covers the column, as Word deletes it
  const merged = tableHasColumnMergeMarkers(topology);
  const covering = merged ? cellsCoveringColumn(topology, op.gridColumnId) : null;
  if (merged && !covering) return 'tree-invariant';
  for (const [rowIndex, { row, cells }] of topology.rows.entries()) {
    // A row another author proposed decides its cells too; a mark in it would hold that decision.
    if (isOtherAuthorsRow(row, author)) return 'unsupported-revision';
    const cell = covering ? covering[rowIndex] : cells[columnIndex]!;
    if (!cell) continue;
    const marker = cellRevisionOf(cell);
    const deletion = cellDeletionOf(cell, marker, author);
    if (deletion === 'refuse') {
      return 'unsupported-revision';
    }

    if (deletion === 'mark') {
      marked.push(cell);
    } else if (deletion === 'withdraw' && marker) {
      withdrawn.push({ cell, marker });
    }
  }

  if (!keepsAColumn(topology, columnIndex, covering)) return 'block-required';
  return { topology, columnIndex, marked, withdrawn };
}

/**
 * Deleting marks the column's cells and nothing else. Reject drops the markers. Accept removes
 * the cells, and the neighbour takes their width.
 */
function applyTrackedColumnDeletion(
  part: OoxmlPart,
  op: DeleteColumnOp,
  revision: RevisionAttributionInput,
  options?: EditOptions
): TreeOpResult {
  const plan = planColumnDeletion(part, op, revision.author);
  if (typeof plan === 'string') return { ok: false, reason: plan };
  const { topology, columnIndex } = plan;
  const taken = new Set(
    [...plan.marked, ...plan.withdrawn.map(({ cell }) => cell)].map((c) => c.id)
  );
  const caretParagraphId = caretBeside(topology, columnIndex, taken);
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
  const markers = plan.withdrawn.map(({ marker }) => marker.id);
  return afterWithdrawal(proposed, markers, effect, options);
}

// ── Insertion ─────────────────────────────────────────────────────────────────────────────

/** A cell's `w:tcW` as the neighbour lends it: split in two, or copied as it is. */
type LentWidth =
  | { readonly split: true; readonly type: 'dxa' | 'pct'; readonly value: number }
  | { readonly split: false; readonly width: OoxmlElement | undefined };

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
  const value = wholeAttribute(width, 'w');
  if ((type !== 'dxa' && type !== 'pct') || value === null) return null;
  return { split: true, type, value };
}

function gridColumnsOf(grid: OoxmlElement): OoxmlElement[] {
  const columns: OoxmlElement[] = [];
  for (const child of grid.children) {
    if (isWmlGridCol(child)) {
      columns.push(child);
    }
  }

  return columns;
}

/** The integral width of one `w:gridCol`, or null. */
function gridWidthAt(grid: OoxmlElement, index: number): number | null {
  return wholeAttribute(gridColumnsOf(grid)[index], 'w');
}

interface ColumnInsertionPlan {
  /** Grid index of the new column. */
  readonly insertIndex: number;
  /** Grid index, BEFORE the insertion, of the column that lends its width. */
  readonly lenderIndex: number;
  /** Grid index, AFTER the insertion, of the column that lends its width. */
  readonly lenderAfter: number;
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
    ) {
      return 'unsupported-revision';
    }
  }

  const lent: LentWidth[] = [];
  for (const { row, cells } of topology.rows) {
    const lender = cells[lenderIndex]!;
    const width = lentWidthOf(lender);
    if (!width) return 'unsupported-revision';

    // A cell-property record lends the same width, or a reject of it would stop matching. Only
    // a row formatting record over every cell: the new cell gets it too, so Word's reject of it
    // does not reset that cell (any other record in the row would)
    const recorded = recordedWidthElement(lender) !== undefined;
    if (
      recorded &&
      (!rowRecordOf(row, cells) || (width.split && !recordedWidthOf(lender, width)))
    ) {
      return 'unsupported-revision';
    }

    lent.push(width);
  }

  const lenderAfter = insertIndex > 0 ? lenderIndex : 1;
  return { insertIndex, lenderIndex, lenderAfter, lent };
}

/** The cell's `w:tcPrChange` and the `w:tcPr` it records, when it has one. */
function recordedWidthElement(
  cell: OoxmlElement
): { readonly history: OoxmlElement; readonly recorded: OoxmlElement } | undefined {
  const properties = wmlChildNamed(cell, 'tcPr');
  const history = properties && wmlChildNamed(properties, 'tcPrChange');
  const recorded = history && wmlChildNamed(history, 'tcPr');
  return history && recorded ? { history, recorded } : undefined;
}

/** A cell's property record, and the width it holds. */
interface RecordedWidth {
  readonly history: OoxmlElement;
  readonly recorded: OoxmlElement;
  readonly value: number;
}

/**
 * The record of a cell and the width it holds, when it can lend what the shown width lends:
 * the same type, and more than half the shown width. Undefined with no record; null when it
 * cannot.
 *
 * Word writes no record for a lent width, and its reject of a cell-property record gives the
 * cell the recorded width back. So the record lends the same half: a reject of it then keeps
 * the new column whole, as the record of a change made after the column holds the width the
 * cell shows (a row height recorded on every cell, Word's form).
 */
function recordedWidthOf(
  cell: OoxmlElement,
  shown: Extract<LentWidth, { split: true }>
): RecordedWidth | undefined | null {
  const record = recordedWidthElement(cell);
  if (!record) return undefined;
  const width = wmlChildNamed(record.recorded, 'tcW');
  const value = wholeAttribute(width, 'w');
  if (!width || wmlAttributeValue(width, 'type') !== shown.type || value === null) {
    return null;
  }

  return value > Math.floor(shown.value / 2) ? { ...record, value } : null;
}

/** The lender with `half` lent from its record's width too (planned by `recordedWidthOf`). */
function withRecordLending(
  cell: OoxmlTableCellNode,
  shown: Extract<LentWidth, { split: true }>,
  half: number,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  const record = recordedWidthOf(cell, shown);
  const properties = wmlChildNamed(cell, 'tcPr');
  if (!record || !properties) return cell;
  const patched = patchTcPrChild(
    record.recorded,
    widthElement(shown.type, record.value - half, nextId, wml)
  );
  if (!patched.ok) return cell;
  const replace = (parent: OoxmlElement, from: OoxmlElement, to: OoxmlElement) =>
    Object.freeze({
      ...parent,
      children: parent.children.map((child) => (child === from ? to : child)),
    }) as OoxmlElement;
  const history = replace(record.history, record.recorded, patched.container);
  return replace(
    cell,
    properties,
    replace(properties, record.history, history)
  ) as OoxmlTableCellNode;
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
 * `w:tblGrid` with the lender halved and the new column holding the other half. Inserting takes
 * the new column's width FROM the neighbour that would take it back, in `w:tblGrid` and in each
 * row's `w:tcW`. Reject then restores the table exactly, and Accept keeps the new column in a
 * table of the same width.
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
  const { insertIndex, lenderIndex, lenderAfter } = plan;
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

/** The author and date of a record, as one key to compare records by. */
function stampOf(history: OoxmlElement | undefined): string | undefined {
  return (
    history &&
    JSON.stringify([wmlAttributeValue(history, 'author'), wmlAttributeValue(history, 'date')])
  );
}

/**
 * Word's plain height: on a row with no `w:trPr`, Word writes the height with no row record and
 * one author's `w:tcPrChange` on each cell, each with an id of its own
 * (`revision-table-implicit-height.ts`). The first cell's record stands for them.
 */
function plainHeightRecord(
  properties: OoxmlElement | undefined,
  histories: readonly (OoxmlElement | undefined)[]
): OoxmlElement | undefined {
  const plain = properties?.children.filter((child) => child.kind !== 'textValue') ?? [];
  const first = histories[0];
  if (
    plain.length !== 1 ||
    !isWmlElement(plain[0]!, 'trHeight') ||
    !first ||
    wmlAttributeValue(first, 'author') === undefined
  ) {
    return undefined;
  }

  return histories.every((history) => stampOf(history) === stampOf(first)) ? first : undefined;
}

/**
 * The engine's row record: one `w:trPrChange`, and a `w:tcPrChange` with its id on each cell.
 */
function rowChangeRecord(
  record: OoxmlElement,
  histories: readonly (OoxmlElement | undefined)[]
): OoxmlElement | undefined {
  const id = wmlAttributeValue(record, 'id');
  if (id === undefined) return undefined;
  const covered = histories.every(
    (history) => history !== undefined && wmlAttributeValue(history, 'id') === id
  );
  return covered ? record : undefined;
}

/** The row's formatting record when it covers every cell of the row (`added` left out). */
function rowRecordOf(
  row: OoxmlElement,
  cells: readonly OoxmlElement[],
  added?: OoxmlElement
): OoxmlElement | undefined {
  const properties = wmlChildNamed(row, 'trPr');
  const record = properties && wmlChildNamed(properties, 'trPrChange');
  const histories = cells
    .filter((cell) => cell !== added)
    .map((cell) => recordedWidthElement(cell)?.history);
  return record ? rowChangeRecord(record, histories) : plainHeightRecord(properties, histories);
}

/**
 * The new cell with the row's formatting record, holding the cell's properties as they show,
 * as that record holds every other cell's. Word's reject of a row's formatting gives a cell
 * with no record default formatting (`revision-unrecorded-cell-format.test.ts`), which would
 * take the new cell's lent width away while its column is still pending. So one file comes
 * out whichever came first, the row's formatting or the column.
 */
function withRowRecord(
  cell: OoxmlTableCellNode,
  record: OoxmlElement,
  nextId: () => string,
  wml: WmlFreshNamespaceContext
): OoxmlTableCellNode {
  const properties = wmlChildNamed(cell, 'tcPr');
  const shown = (properties?.children ?? []).map((child) => cloneWithNewIds(child, nextId));
  const history = freshWmlElement(
    'tcPrChange',
    nextId,
    wml,
    record.attributes.map((attribute) => ({ ...attribute })),
    [freshWmlElement('tcPr', nextId, wml, [], shown)]
  );
  return withCellProperty(cell, history, nextId, wml);
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

  // The untracked op builds the column (cells, paragraphs, grid entry) exactly as editing mode
  // does. This lends it its width and proposes it.
  const inserted = applyInsertTableColumn(part, op, options);
  if (!inserted.ok) return inserted;
  const after = readEditableTableTopology(inserted.part.root, op.tableId);
  const grid = after.ok ? after.topology.grid : undefined;
  if (!after.ok || !grid) return { ok: false, reason: 'tree-invariant' };
  const { topology } = after;
  const { insertIndex, lenderAfter } = plan;
  const nextId = createNodeIdAllocator(inserted.part);
  const wml = wmlFreshNamespaceContextAt(inserted.part, topology.table);
  const replaced: OoxmlElement[] = [lentGrid(grid, plan, nextId, wml)];
  for (const [rowIndex, { row, cells }] of topology.rows.entries()) {
    const lent = plan.lent[rowIndex]!;
    let lender = cells[lenderAfter]!;
    let cell = cells[insertIndex]!;
    if (lent.split) {
      const half = Math.floor(lent.value / 2);
      const kept = widthElement(lent.type, lent.value - half, nextId, wml);
      lender = withRecordLending(
        withCellProperty(lender, kept, nextId, wml),
        lent,
        half,
        nextId,
        wml
      );
      cell = withCellProperty(cell, widthElement(lent.type, half, nextId, wml), nextId, wml);
    } else {
      // The new cell takes the lender's own `auto` or `nil` width, or none when it has none.
      const width = lent.width && ({ ...lent.width, id: nextId() } as OoxmlElement);
      cell = withCellProperty(cell, width ?? { remove: 'tcW' }, nextId, wml);
    }

    const record = rowRecordOf(row, cells, cells[insertIndex]!);
    if (record) cell = withRowRecord(cell, record, nextId, wml);
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
 * the op's untracked form passed. Refused with `unsupported-revision`, never approximated.
 */
export function validateTrackedColumnEdit(
  part: OoxmlPart,
  op: TableColumnOp,
  revision: RevisionAttributionInput
): TreeOpRejection | null {
  if (op.op === 'insertTableColumn') {
    const merged = mergedTopology(part, op);
    if (merged) {
      const result = applyTrackedMergedColumnInsertion(
        part,
        op,
        merged,
        cellInsertion('0', revision)
      );
      return result.ok ? null : result.reason;
    }
  }

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
  if (op.op === 'deleteTableColumn') return applyTrackedColumnDeletion(part, op, revision, options);
  const merged = mergedTopology(part, op);
  if (!merged) return applyTrackedColumnInsertion(part, op, revision, options);

  // ONE id for every new cell, as for a table without merges
  const revisionId = options?.revisionIds?.() ?? nextRevisionId(part)();
  return applyTrackedMergedColumnInsertion(
    part,
    op,
    merged,
    cellInsertion(revisionId, revision),
    options
  );
}

/** The table's topology when it holds merged cells (Word's column rules for them), else null. */
function mergedTopology(part: OoxmlPart, op: InsertColumnOp): EditableTableTopology | null {
  const result = readEditableTableTopology(part.root, op.tableId);
  return result.ok && tableHasColumnMergeMarkers(result.topology) ? result.topology : null;
}

/** A `w:cellIns` for the proposal, minted where the cell is built. */
const cellInsertion =
  (revisionId: string, revision: RevisionAttributionInput) =>
  (nextId: () => string, wml: WmlFreshNamespaceContext): OoxmlElement =>
    freshWmlElement('cellIns', nextId, wml, trackedRevisionAttributes(revisionId, revision, wml));
