// Columns: insert and delete (tasks 3.1 and 9.7).

import type { OoxmlElement, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { build, revisionAttributes } from '../../../store/store/tree-op-tracked.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { insertedPending, statusOf } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem, OrderedPropertyTarget, OrderedSuggestion } from '../model/types.ts';
import { childNamed, isWml, withChildren } from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import { refuseLockedTable } from '../locked-content.ts';
import { retractSuggestion } from './retract.ts';
import { holdWithin, nestedIn, suggestionFor, takeBack, withNewIds, type Ctx } from './context.ts';
import {
  blockOf,
  cellsOf,
  elementRecord,
  gridColumnOfCell,
  openOf,
  rowCells,
  rowsOf,
  same,
  tableCellSpaces,
  withRecord,
} from './table-nodes.ts';
import {
  engineScratch,
  revisionOf,
  shellId,
  templateItems,
  TEMPLATE_REVISION,
  viewNode,
  viewPart,
} from './templates.ts';
import { takeBackNotesOf } from './note-body.ts';
import { recordedPending, struckAlready } from './text.ts';

/** The cell with a template `w:cellIns`, so it loads as the edit's suggestion's own cell. */
function asInsertedCell(ctx: Ctx, cell: OoxmlElement): OoxmlElement {
  const marker = build(
    ctx.ids.node(),
    'generic',
    'cellIns',
    revisionAttributes('0', TEMPLATE_REVISION),
    []
  );
  const tcPr = childNamed(cell, 'tcPr');
  if (!tcPr)
    return withChildren(cell, [
      build(ctx.ids.node(), 'generic', 'tcPr', [], [marker]),
      ...cell.children,
    ]);
  return withChildren(
    cell,
    cell.children.map((c) => (c === tcPr ? withChildren(tcPr, [...tcPr.children, marker]) : c))
  );
}

/** Whether a cell's properties hold a `w:cellDel`. */
function hasCellDel(cell: OoxmlElement): boolean {
  const tcPr = childNamed(cell, 'tcPr');
  return tcPr?.children.some((m) => isWml(m, 'cellDel')) === true;
}

/** The cell without a `w:tcPrChange` in its properties. */
function withoutRecord(cell: OoxmlElement): OoxmlElement {
  const tcPr = childNamed(cell, 'tcPr');
  const record = tcPr && childNamed(tcPr, 'tcPrChange');
  if (!tcPr || !record) return cell;
  const properties = withChildren(
    tcPr,
    tcPr.children.filter((c) => c !== record)
  );
  return withChildren(
    cell,
    cell.children.map((c) => (c === tcPr ? properties : c))
  );
}

/**
 * The new cells the engine's op added, as items of the suggestion, each next to the cell it
 * stands beside in its row (after the cell before it, or before the cell after it). Returns
 * their open items by the engine's node id.
 */
function placeNewCells(
  ctx: Ctx,
  before: OoxmlElement,
  after: OoxmlElement,
  sid: string
): Map<string, OrderedItem> {
  const added = new Map<string, OrderedItem>();
  const old = cellsOf(before);
  for (const row of rowsOf(after)) {
    const cells = rowCells(row);
    cells.forEach((cell, k) => {
      if (old.has(cell.id)) return;
      const previous = cells[k - 1];
      const following = cells[k + 1];
      const afterPrevious = previous !== undefined && old.has(previous.id);
      let beside: OrderedItem | null = null;
      if (afterPrevious) {
        beside = openOf(ctx, previous);
      } else if (following && old.has(following.id)) {
        beside = openOf(ctx, following);
      }

      if (!beside) throw new OrderedRefusal('engine-refused', 'a new cell with no old neighbour');
      const at = afterPrevious ? closeOf(ctx.story, beside) : ctx.story.tree.prev(beside);

      // The row's formatting record the engine gave the cell is that change's own record
      // (`recordRowFormatting`), not one of this suggestion
      const items = templateItems(
        ctx,
        ctx.direct ? asInsertedCell(ctx, cell) : withoutRecord(cell),
        sid
      ).map((item) => nestedIn(item, beside.within));
      insertItems(ctx.doc, ctx.story, at, items);
      const first = items[0];
      if (first) added.set(cell.id, first);
    });
  }

  return added;
}

/**
 * The widths the op changed (the grid, the cells it took width from, and the new cells' own),
 * as records of the suggestion that hold the cells the table shows after it (`tableCells`). A
 * new cell's record says which cells showed when its width was lent, as the others' do. A
 * suggestion's records show with no change record while pending, as Word writes none, so a
 * decision on the column gives the width back at compose by the engine's own rule
 * (`compose/table-columns.ts`).
 */
function recordWidths(
  ctx: Ctx,
  before: OoxmlElement,
  after: OoxmlElement,
  sid: string,
  tableCells: readonly string[],
  added: ReadonlyMap<string, OrderedItem> = new Map()
): void {
  const flags = ctx.direct ? {} : { untracked: true };
  const record = (target: OrderedPropertyTarget, element: OoxmlElement) =>
    elementRecord(ctx, sid, target, element, tableCells, flags);

  // A suggestion lends from the original of the pending records of the grid and of a cell's width
  // too (the engine's `lentGrid`, `withRecordLending`): the record names them (`over`), so their
  // reject keeps the column lent. The grid's keeps that original as the engine split it
  // (`w:tblGridChange`): only it says where the column went
  const lent = (open: OrderedItem, target: 'tblGrid' | 'tcPr', element: OoxmlElement) => {
    const given = record(target, element);
    const over = ctx.direct ? [] : recordedPending(ctx, open, target);
    if (over.length === 0 || given.element === null) {
      return given;
    }

    const original = childNamed(element, 'tblGridChange');
    const children = [...given.element.children, ...(original ? [withNewIds(ctx, original)] : [])];
    return { ...given, over, element: withChildren(given.element, children) };
  };
  const tableOpen = openOf(ctx, before);
  if (!tableOpen) throw new OrderedRefusal('malformed-story', 'a table with no open item');
  for (const name of ['tblPr', 'tblGrid'] as const) {
    const was = childNamed(before, name);
    const now = childNamed(after, name);
    if (!now || same(was, now)) {
      continue;
    }

    withRecord(ctx, tableOpen, name === 'tblGrid' ? lent(tableOpen, name, now) : record(name, now));
  }

  const old = cellsOf(before);
  for (const [id, cell] of cellsOf(after)) {
    const was = old.get(id);
    const now = childNamed(cell, 'tcPr');
    if (!now || (was && same(childNamed(was, 'tcPr'), now))) continue;
    const open = was ? openOf(ctx, cell) : added.get(id);
    if (open) {
      withRecord(ctx, open, lent(open, 'tcPr', now));
    }
  }

  // A row's columns skipped before or after its cells, which Word's column rules for merged
  // cells change (`w:gridBefore`, `w:gridAfter` and their widths)
  const oldRows = new Map(rowsOf(before).map((row) => [row.id, row]));
  for (const row of rowsOf(after)) {
    const was = oldRows.get(row.id);
    const now = childNamed(row, 'trPr');
    const then = was ? childNamed(was, 'trPr') : undefined;
    if (!was || same(then, now)) continue;
    const open = openOf(ctx, row);
    const properties = now ?? (then ? withChildren(then, []) : undefined);
    if (open && properties) withRecord(ctx, open, record('trPr', properties));
  }
}

/** Strike a cell and everything in it with the suggestion (a `w:cellDel`). */
function strikeCell(ctx: Ctx, open: OrderedItem, sid: string): void {
  const close = closeOf(ctx.story, open);
  for (const edge of [open, close]) {
    relabel(ctx.doc, ctx.story, edge, { dels: [...edge.dels, sid] });
  }

  for (const item of itemsInside(ctx.story, open)) {
    holdWithin(ctx, ctx.story, item, sid);
  }
}

/** Take back a cell of the author's own pending suggestion, and everything in it. */
function withdrawCell(ctx: Ctx, open: OrderedItem): void {
  const items = [open, ...itemsInside(ctx.story, open), closeOf(ctx.story, open)];
  for (const item of items) takeBack(ctx, ctx.story, item);

  // A new note whose own reference the cell held goes with it
  takeBackNotesOf(ctx, items);
}

/**
 * A new column beside a grid column. A suggestion runs the engine's tracked op, the form a
 * reopened Word file has: new cells with `w:cellIns` and the neighbour's width lent to them with
 * no change record. A direct column runs the untracked op, as the editor alone does with
 * tracking off: it takes its reference column's width and the table grows.
 */
function insertColumn(
  ctx: Ctx,
  view: OoxmlPart,
  table: OoxmlElement,
  cellId: string,
  where: 'left' | 'right'
): string {
  const gridColumnId = gridColumnOfCell(table, cellId);

  // A table with no grid: the engine builds one from the cells (refused when proposed)
  const target = gridColumnId === null ? { referenceCellId: cellId } : { gridColumnId };
  const op: TreeDocOp = {
    op: 'insertTableColumn',
    tableId: table.id,
    where,
    ...target,
    ...(ctx.direct ? {} : { revision: revisionOf(ctx) }),
  };
  const after = viewNode(engineScratch(view, [table], op), table.id);
  const suggestion = suggestionFor(ctx);
  const added = placeNewCells(ctx, table, after, suggestion.id);
  const tableCells = [
    ...tableCellSpaces(ctx, after),
    ...[...added.values()].map((open) => open.space),
  ];
  recordWidths(ctx, table, after, suggestion.id, tableCells, added);
  if (!ctx.direct) recordRowFormatting(ctx, after, added, tableCells);
  return suggestion.id;
}

/**
 * The pending change a row's formatting record names (the latest one it records), if any: a
 * row record, or Word's height with no row record (`compose/row-heights.ts`).
 */
function rowRecordSuggestion(ctx: Ctx, row: OrderedItem): string | undefined {
  const recorded = row.props.filter(
    (p) =>
      p.kind === 'element' &&
      p.target === 'trPr' &&
      (p.untracked !== true || (p.element !== null && childNamed(p.element, 'trHeight'))) &&
      p.acceptedOnly !== true &&
      statusOf(ctx.doc, p.suggestion) === 'pending'
  );
  return recorded.sort((a, b) => a.seq - b.seq).at(-1)?.suggestion;
}

/**
 * A new cell in a row whose formatting a pending change records on every cell: the engine gives
 * it that record too, holding the cell's properties as they show (`withRowRecord`), so a reject
 * of the row's formatting keeps the cell's lent width. That change records it here.
 */
function recordRowFormatting(
  ctx: Ctx,
  after: OoxmlElement,
  added: ReadonlyMap<string, OrderedItem>,
  tableCells: readonly string[]
): void {
  for (const row of rowsOf(after)) {
    const rowOpen = openOf(ctx, row);
    const suggestion = rowOpen && rowRecordSuggestion(ctx, rowOpen);
    if (!suggestion) continue;
    for (const cell of rowCells(row)) {
      const open = added.get(cell.id);
      const tcPr = childNamed(cell, 'tcPr');
      if (!open || !tcPr || !childNamed(tcPr, 'tcPrChange')) continue;

      // It changes nothing (`merge: []`): the cell's properties stay the column's own
      const record = elementRecord(ctx, suggestion, 'tcPr', tcPr, tableCells, {
        recordUnchanged: true,
        merge: [],
      });
      relabel(ctx.doc, ctx.story, open, { props: [...open.props, record] });
    }
  }
}

/** Delete a grid column, as the engine's tracked or untracked op deletes it. */
function deleteColumn(
  ctx: Ctx,
  view: OoxmlPart,
  table: OoxmlElement,
  cellId: string
): string | null {
  const gridColumnId = gridColumnOfCell(table, cellId);
  if (gridColumnId === null)
    throw new OrderedRefusal('engine-refused', 'deleteTableColumn: unknown-grid-column');
  const op: TreeDocOp = {
    op: 'deleteTableColumn',
    tableId: table.id,
    gridColumnId,
    ...(ctx.direct ? {} : { revision: revisionOf(ctx) }),
  };
  const after = viewNode(engineScratch(view, [table], op), table.id);
  return ctx.direct
    ? deleteColumnDirect(ctx, table, after)
    : deleteColumnTracked(ctx, table, after);
}

/**
 * A direct deletion: the untracked op removed the column's cells, so the table shrinks. They
 * are struck with one settled deletion, which also records the widths the op changed.
 */
function deleteColumnDirect(ctx: Ctx, table: OoxmlElement, after: OoxmlElement): string | null {
  const now = cellsOf(after);
  let suggestion: OrderedSuggestion | null = null;
  for (const [id, cell] of cellsOf(table)) {
    const open = openOf(ctx, cell);
    if (open && !now.has(id)) {
      suggestion ??= suggestionFor(ctx);
      strikeCell(ctx, open, suggestion.id);
    }
  }

  if (suggestion !== null) {
    recordWidths(ctx, table, after, suggestion.id, tableCellSpaces(ctx, after));
  }

  return suggestion?.id ?? null;
}

/**
 * A suggested deletion: `w:cellDel` on the column's cells, and no widths (compose removes the
 * cells once it is accepted). The author's own pending cells are taken back, as the engine
 * withdraws them.
 */
function deleteColumnTracked(ctx: Ctx, table: OoxmlElement, after: OoxmlElement): string | null {
  const now = cellsOf(after);
  let suggestion: OrderedSuggestion | null = null;
  let withdrawn: string | null = null;
  for (const [id, cell] of cellsOf(table)) {
    const open = openOf(ctx, cell);
    if (!open) {
      continue;
    }

    const kept = now.get(id);
    if (!kept) {
      // The engine withdrew the author's own pending cell: taken back
      if (open.by !== null && insertedPending(ctx.doc, open)) withdrawn ??= open.by;
      withdrawCell(ctx, open);
      continue;
    }

    // Policy 1: a cell another pending deletion struck keeps that one deletion
    if (hasCellDel(kept) && !hasCellDel(cell) && !struckAlready(ctx, open.dels)) {
      suggestion ??= suggestionFor(ctx);
      strikeCell(ctx, open, suggestion.id);
    }
  }

  // The author's own new column withdrawn whole: its lent widths and grid go too, as the
  // engine's withdrawal gives them back, so no card stays where a reopened file has none
  if (withdrawn !== null && withdrawnWhole(ctx, withdrawn)) {
    retractSuggestion(ctx.doc, withdrawn, ctx.attribution);
  }

  return suggestion?.id ?? withdrawn;
}

/** Whether every cell a pending column suggestion inserted is taken back. */
function withdrawnWhole(ctx: Ctx, suggestion: string): boolean {
  for (const item of ctx.doc.labelled.get(suggestion) ?? [])
    if (item.by === suggestion && !item.retracted) return false;
  return true;
}

/**
 * Insert a column beside the caret's cell (`left` or `right` of the grid column it starts at),
 * or delete that grid column. Returns the suggestion the edit made, or null when a deletion
 * struck nothing new (every cell was already struck: policy 1).
 */
export function columnOp(
  ctx: Ctx,
  p: OrderedViewParagraph,
  kind: 'insert' | 'delete',
  where: 'left' | 'right' = 'right'
): string | null {
  refuseLockedTable(ctx, p);
  const table = blockOf(p, 'tbl');
  const cell = blockOf(p, 'tc');
  const view = viewPart(ctx);
  const tableNode = viewNode(view, shellId(table));
  return kind === 'insert'
    ? insertColumn(ctx, view, tableNode, shellId(cell), where)
    : deleteColumn(ctx, view, tableNode, shellId(cell));
}
