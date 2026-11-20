// Table, row and cell property presses (tasks 3.1 and 9.7): the engine's own op on a scratch
// part that holds the table as the view draws it.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import type { RevisionAttributionInput, TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { wholeMergedCellIds } from '../../../store/store/tree-op-table-merged-cells.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { inserter, oneMaker, statusOf } from '../model/document.ts';
import { relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedElementRecord, OrderedItem, OrderedPropertyTarget } from '../model/types.ts';
import { childNamed } from '../ooxml-nodes.ts';
import { refuseLockedTable } from '../placement.ts';
import { pendingOwn, suggestionFor, type Ctx } from './context.ts';
import {
  blockOf,
  cellColumns,
  cellsOf,
  elementRecord,
  gridColumnsOf,
  openOf,
  rowCells,
  rowsOf,
  changedKeys,
  same,
  tableCellSpaces,
  withRecord,
} from './table-nodes.ts';
import type { OrderedTableOp } from './tables.ts';
import { takenOver } from './text.ts';
import { engineScratch, revisionOf, shellId, viewNode, viewPart } from './templates.ts';

/**
 * The grid columns a resize names, read from the cell it names. Dragging the edge between two
 * columns moves the cell's right edge (the grid column it ends at and the next one keep their
 * sum); dragging the table's right edge names a cell that ends at the last grid column. A cell
 * keeps its address while columns come and go, so the change still names the edge its sender
 * dragged; a resize whose widths no longer fit (another person added or removed a column there
 * first) is refused by the engine, by name.
 */
function resized(
  table: OoxmlElement,
  cellId: string,
  op: Extract<OrderedTableOp, { op: 'setTableColumnWidths' | 'setTableRightEdgeWidth' }>
): { readonly left: string; readonly right: string | undefined } {
  const at = cellColumns(table, cellId);
  const grid = gridColumnsOf(table);
  const last = at === null ? undefined : grid[at.start + at.span - 1];
  if (at === null || !last)
    throw new OrderedRefusal('engine-refused', `${op.op}: unknown-grid-column`);
  return { left: last.id, right: grid[at.start + at.span]?.id };
}

/** The engine's op for a press at the caret's table, row or cells. */
function tableTreeOp(
  table: OoxmlElement,
  rowId: string,
  cellIds: readonly string[],
  op: OrderedTableOp,
  revision: RevisionAttributionInput | undefined
): TreeDocOp {
  const tableId = table.id;
  const tracked = revision === undefined ? {} : { revision };
  switch (op.op) {
    case 'setTableProperties':
      return { ...op, tableId, ...tracked };
    case 'setTableRowHeight':
      return { op: op.op, tableId, rowId, heightTwips: op.heightTwips, ...tracked };
    case 'setTableCellFill':
      return { op: op.op, tableId, cellIds, color: op.color, ...tracked };
    case 'setTableCellVerticalAlignment':
      return { op: op.op, tableId, cellIds, alignment: op.alignment, ...tracked };
    case 'setTableCellBorders':
      return 'spec' in op
        ? { op: op.op, tableId, cellIds, scope: op.scope, spec: op.spec, ...tracked }
        : { op: op.op, tableId, cellIds, scope: 'none', target: op.target, ...tracked };
    case 'setTableColumnWidths': {
      const { left, right } = resized(table, cellIds[0] ?? '', op);
      if (right === undefined)
        throw new OrderedRefusal('engine-refused', `${op.op}: no grid column after the cell`);
      return {
        op: op.op,
        tableId,
        leftGridColumnId: left,
        rightGridColumnId: right,
        leftWidthTwips: op.leftWidthTwips,
        rightWidthTwips: op.rightWidthTwips,
        ...tracked,
      };
    }
    case 'setTableRightEdgeWidth':
      return {
        op: op.op,
        tableId,
        gridColumnId: resized(table, cellIds[0] ?? '', op).left,
        columnWidthTwips: op.columnWidthTwips,
        tableWidthTwips: op.tableWidthTwips,
        ...tracked,
      };
  }
}

/** An op on cells, which the engine applies to whole merged cells only. */
const isCellOp = (op: OrderedTableOp): boolean =>
  op.op === 'setTableCellFill' ||
  op.op === 'setTableCellBorders' ||
  op.op === 'setTableCellVerticalAlignment';

/** A resize: it names exactly one cell. */
const isResize = (op: OrderedTableOp): boolean =>
  op.op === 'setTableColumnWidths' || op.op === 'setTableRightEdgeWidth';

/** Whether properties hold a change record the engine wrote (`w:tcPrChange` and the like). */
const holdsChangeRecord = (element: OoxmlElement | undefined): element is OoxmlElement =>
  element !== undefined &&
  element.children.some((c) => c.kind !== 'textValue' && c.localName.endsWith('Change'));

/**
 * Whether a container's properties take a record: tracked, when the engine wrote a change record
 * in them, changed or not (a row's cells are one decision in Word); direct, when they changed.
 */
function keeps(
  tracked: boolean,
  was: OoxmlElement | undefined,
  now: OoxmlElement | undefined
): now is OoxmlElement {
  return tracked ? holdsChangeRecord(now) : now !== undefined && !same(was, now);
}

/**
 * How a record of the op sets its container's properties. A tracked row or cell record sets
 * only what the op changed: the rest (a column's lent width, pending under another card) stays
 * that change's own, whatever is decided first. So does a table's record that changed nothing
 * (a row height's anchor, Word's form).
 */
function recordFlags(
  tracked: boolean,
  target: OrderedPropertyTarget,
  element: OoxmlElement,
  was: OoxmlElement | undefined
): Pick<OrderedElementRecord, 'recordUnchanged' | 'merge'> {
  const table = target === 'tblPr' || target === 'tblGrid';
  const merge = tracked && (!table || (was !== undefined && same(was, element)));
  const keys = merge && !table ? changedKeys(was, element) : [];
  return { ...(tracked ? { recordUnchanged: true } : {}), ...(merge ? { merge: keys } : {}) };
}

/**
 * What the op set in a container of its author's own new row or table: a record of that
 * insertion, shown with no change record (`untracked`), as the engine writes none there and its
 * reject takes the container away whole. It sets only what changed, as `ownFormat` does on words.
 */
function recordOwn(
  ctx: Ctx,
  own: string,
  open: OrderedItem,
  target: OrderedPropertyTarget,
  was: OoxmlElement | undefined,
  now: OoxmlElement | undefined,
  tableCells: readonly string[]
): void {
  if (!now || same(was, now)) {
    return;
  }

  const flags = { untracked: true, merge: changedKeys(was, now) };
  withRecord(ctx, open, elementRecord(ctx, own, target, now, tableCells, flags));
}

/**
 * Record what the engine's op changed, as records of the suggestion (`sid` makes it at its first
 * record): a suggestion's under the edit's own author, as Word's form records them
 * (`w:tblPrChange`, `w:tblGridChange`, `w:trPrChange`, `w:tcPrChange`); a direct press's as the
 * editor alone keeps what changed. Each record holds the cells the table showed, for a later
 * decision on a column (`compose/table-columns.ts`). A row its author proposed, and a table whose
 * rows all are, take records of that insertion (`recordOwn`), as the engine records nothing in
 * them. Returns the suggestion that holds the press, if any.
 */
function recordTableChanges(
  ctx: Ctx,
  before: OoxmlElement,
  after: OoxmlElement,
  sid: () => string,
  tracked: boolean
): string | null {
  const tableCells = tableCellSpaces(ctx, after);
  let made: string | null = null;
  const record = (target: OrderedPropertyTarget, element: OoxmlElement, was?: OoxmlElement) =>
    elementRecord(
      ctx,
      (made ??= sid()),
      target,
      element,
      tableCells,
      recordFlags(tracked, target, element, was)
    );

  // The pending insertion of a container, when it is the author's own (`insideOwnRow`)
  const ownOf = (open: OrderedItem) => pendingOwn(ctx, inserter(ctx.doc, open))?.id ?? null;
  const tableOpen = openOf(ctx, before);
  if (!tableOpen) throw new OrderedRefusal('malformed-story', 'a table with no open item');
  const rows = rowsOf(after).flatMap((row) => {
    const open = openOf(ctx, row);
    return open ? [{ row, open, own: ownOf(open) }] : [];
  });
  const ownTable = rows.every((r) => r.own !== null) ? ownOf(tableOpen) : null;
  let own = ownTable;
  for (const name of ['tblPr', 'tblGrid'] as const) {
    const now = childNamed(after, name);
    const was = childNamed(before, name);
    if (ownTable !== null) {
      recordOwn(ctx, ownTable, tableOpen, name, was, now, tableCells);
    } else if (keeps(tracked, was, now)) {
      withRecord(ctx, tableOpen, record(name, now, was));
    }
  }

  const oldRows = new Map(rowsOf(before).map((row) => [row.id, row]));
  const oldCells = cellsOf(before);
  for (const { row, open: rowOpen, own: ownRow } of rows) {
    const was = oldRows.get(row.id);
    const trPr = childNamed(row, 'trPr');
    const oldTrPr = was && childNamed(was, 'trPr');
    const cells = rowCells(row).flatMap((cell) => {
      const open = openOf(ctx, cell);
      const old = oldCells.get(cell.id);
      return open
        ? [{ open, tcPr: childNamed(cell, 'tcPr'), was: old && childNamed(old, 'tcPr') }]
        : [];
    });
    if (ownRow !== null) {
      own = ownRow;
      recordOwn(ctx, ownRow, rowOpen, 'trPr', oldTrPr, trPr, tableCells);
      for (const cell of cells) {
        recordOwn(ctx, ownRow, cell.open, 'tcPr', cell.was, cell.tcPr, tableCells);
      }

      continue;
    }

    if (keeps(tracked, oldTrPr, trPr)) withRecord(ctx, rowOpen, record('trPr', trPr, oldTrPr));
    // Word's height on a row with no properties: no row record, shown while it is pending; its
    // anchor is the table's record, or the row's own (`compose/row-heights.ts`)
    else if (tracked && trPr && !same(oldTrPr, trPr))
      withRecord(ctx, rowOpen, { ...record('trPr', trPr, oldTrPr), untracked: true });
    for (const cell of cells) {
      if (keeps(tracked, cell.was, cell.tcPr)) {
        withRecord(ctx, cell.open, record('tcPr', cell.tcPr, cell.was));
      }
    }

    if (tracked && made !== null) {
      heightTakenOver(ctx, row, made);
    }
  }

  return made ?? own;
}

/**
 * Policy 2 for a row height with no row record: its file form is its row's cell records (Word's
 * one `w:tcPrChange` on each cell, which a reopened file reads the height with). So once this
 * suggestion took over every cell record of the same author's height on the row, it takes the
 * height too, as the editor alone keeps one record for both (random session seed 19). Another
 * author's height stays theirs, on its row's own record (D17, rule 4). A retract gives it back
 * with the cell records.
 */
function heightTakenOver(ctx: Ctx, row: OoxmlElement, sid: string): void {
  const rowOpen = openOf(ctx, row);
  if (!rowOpen) return;
  const cells = rowCells(row).flatMap((cell) => openOf(ctx, cell) ?? []);
  const cellRecords = cells.flatMap((cell) =>
    cell.props.filter((r) => r.kind === 'element' && r.target === 'tcPr')
  );
  const props = rowOpen.props.map((p) => {
    const heightSuggestion = p.suggestion;
    if (
      p.kind !== 'element' ||
      !oneMaker(ctx.doc, heightSuggestion, sid) ||
      p.target !== 'trPr' ||
      p.untracked !== true ||
      heightSuggestion === sid ||
      statusOf(ctx.doc, heightSuggestion) !== 'pending' ||
      cellRecords.some((r) => r.suggestion === heightSuggestion) ||
      !cellRecords.some((r) => r.suggestion === sid && r.replaced === heightSuggestion)
    ) {
      return p;
    }

    return takenOver(ctx, p, sid);
  });
  if (props.some((p, k) => p !== rowOpen.props[k])) relabel(ctx.doc, ctx.story, rowOpen, { props });
}

/**
 * A table, row or cell property press (one press is one record), as the engine records it. A
 * cell op covers `cells` (open items of cells of the caret's table) when given, else the
 * caret's cell; a resize names exactly one cell. Null when the press changes no value: the
 * engine records nothing, so a pending record of the same value keeps its author. A suggestion
 * runs the op tracked; an organiser's direct press runs it untracked, as the editor alone does.
 * Over another person's pending record a direct press shows at once, and a reject of that record
 * takes it away too (Word's one record, as the engine's reject does).
 */
export function tableProperty(
  ctx: Ctx,
  p: OrderedViewParagraph,
  op: OrderedTableOp,
  cells?: readonly OrderedItem[]
): string | null {
  refuseLockedTable(ctx, p);
  const table = blockOf(p, 'tbl');
  const row = blockOf(p, 'tr');
  const targets = cells && cells.length > 0 ? cells : [blockOf(p, 'tc')];
  if (isResize(op) && targets.length !== 1)
    throw new OrderedRefusal('invalid-edit', 'a resize names one cell');
  const view = viewPart(ctx);
  const tableNode = viewNode(view, shellId(table));
  const own = cellsOf(tableNode);
  const named = targets.map((cell) => {
    const id = cell.kind === 'open' && cell.token === 'tc' ? shellId(cell) : null;
    if (id === null || !own.has(id))
      throw new OrderedRefusal('not-in-table', "a cell outside the caret's table");
    return id;
  });

  // A cell of a vertical merge names the merged cell: the engine wants all its cells
  const cellIds = isCellOp(op) ? wholeMergedCellIds(view, tableNode.id, named) : named;

  // Tracked under the edit's own author, so the engine knows the author's own records
  const revision = ctx.direct ? undefined : revisionOf(ctx);
  const treeOp = tableTreeOp(tableNode, shellId(row), cellIds, op, revision);
  const after = viewNode(engineScratch(view, [tableNode], treeOp), tableNode.id);
  if (after === tableNode) return null;
  return recordTableChanges(ctx, tableNode, after, () => suggestionFor(ctx).id, !ctx.direct);
}
