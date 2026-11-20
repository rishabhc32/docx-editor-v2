// Rows, columns, new tables and table properties. Each one is the engine's own op on a scratch
// part that holds the table as the view draws it; what the op changed becomes items and labels.
// A row's or a cell's suggestion also labels what is inside it (`within`), so a rejected new row
// and an accepted struck row go with everything in them.

import { parentNodeOf } from '../../../store/package/ooxml-edit.ts';
import type {
  TableBorderEdgeTarget,
  TableBorderSpecInput,
  TreeDocColorValue,
} from '../../../store/store/tree-op-types.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { insertedPending } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import type { OrderedItem } from '../model/types.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import { isWml } from '../ooxml-nodes.ts';
import {
  closeOf,
  itemsInside,
  refuseLockedBlocks,
  refuseLockedTable,
  refuseLockedWithin,
} from '../placement.ts';
import { holdWithin, nestedIn, pendingOwn, suggestionFor, takeBack, type Ctx } from './context.ts';
import { takeBackNotesOf } from './note-body.ts';
import { othersPending } from './retract.ts';
import { blockOf } from './table-nodes.ts';
import { struckAlready } from './text.ts';
import {
  engineScratch,
  scratchBlocks,
  shellId,
  templateItems,
  TEMPLATE_REVISION,
  viewNode,
  viewPart,
} from './templates.ts';

/**
 * A table, row or cell property press at the caret, as the engine's op without its table,
 * targets and revision: the caret's table, row or cell is the target.
 *
 * @public
 */
export type OrderedTableOp =
  | {
      readonly op: 'setTableProperties';
      readonly styleId?: string;
      readonly headerRowCount?: number;
      readonly columnWidthsTwips?: readonly number[];
      readonly justification?: 'left' | 'center' | 'right' | null;
      readonly preferredWidth?: {
        readonly type: 'auto' | 'dxa' | 'pct';
        readonly value: number;
      } | null;
    }
  | { readonly op: 'setTableRowHeight'; readonly heightTwips: number }
  | { readonly op: 'setTableCellFill'; readonly color: TreeDocColorValue | null }
  | {
      readonly op: 'setTableCellVerticalAlignment';
      readonly alignment: 'top' | 'center' | 'bottom';
    }
  | {
      readonly op: 'setTableCellBorders';
      readonly scope: 'none';
      readonly target: TableBorderEdgeTarget;
    }
  | {
      readonly op: 'setTableCellBorders';
      readonly scope: TableBorderEdgeTarget;
      readonly spec: TableBorderSpecInput;
    }
  | {
      /**
       * Drag the edge between two columns: the right edge of the cell the change names. The
       * grid column the cell ends at takes `leftWidthTwips`, the next one `rightWidthTwips`;
       * their sum stays.
       */
      readonly op: 'setTableColumnWidths';
      readonly leftWidthTwips: number;
      readonly rightWidthTwips: number;
    }
  | {
      /**
       * Drag the table's right edge: the last grid column, which the cell the change names ends
       * at, takes `columnWidthTwips`, and the table `tableWidthTwips`.
       */
      readonly op: 'setTableRightEdgeWidth';
      readonly columnWidthTwips: number;
      readonly tableWidthTwips: number;
    };

// ---- rows ----

/** Insert a row next to the caret's row, as the engine's command builds it. */
export function insertRow(ctx: Ctx, p: OrderedViewParagraph, where: 'above' | 'below'): string {
  refuseLockedTable(ctx, p);
  const row = blockOf(p, 'tr');
  const table = blockOf(p, 'tbl');
  const view = viewPart(ctx);
  const tableNode = viewNode(view, shellId(table));
  const before = new Set(tableNode.children.map((c) => c.id));
  const scratch = engineScratch(view, [tableNode], {
    op: 'insertTableRow',
    tableId: tableNode.id,
    rowId: shellId(row),
    where,
    revision: TEMPLATE_REVISION,
  });
  const inserted = viewNode(scratch, tableNode.id).children.find(
    (c) => (c.kind === 'tableRow' || isWml(c, 'tr')) && !before.has(c.id)
  );
  if (!inserted || inserted.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', 'the engine inserted no row');
  const suggestion = suggestionFor(ctx);
  const items = templateItems(ctx, inserted, suggestion.id);
  const after = where === 'below' ? closeOf(ctx.story, row) : ctx.story.tree.prev(row);
  insertItems(ctx.doc, ctx.story, after, items);
  return suggestion.id;
}

/**
 * Strike the caret's row: its row items, and everything in it through `within`, as a file's
 * struck row loads. Your own pending row is taken back; refused when another author changed
 * something in it.
 */
export function deleteRow(
  ctx: Ctx,
  p: OrderedViewParagraph,
  options: { readonly evenStruck?: boolean } = {}
): string | null {
  // The engine's row deletion rearranges the table and removes the row: a lock around the
  // table or on a control in it keeps it, and so does one in the row that keeps its control
  refuseLockedTable(ctx, p);
  const table = blockOf(p, 'tbl');
  refuseLockedWithin(ctx, itemsInside(ctx.story, table, ctx.doc), false);
  const row = blockOf(p, 'tr');
  refuseLockedWithin(ctx, itemsInside(ctx.story, row, ctx.doc), true);
  const inside = itemsInside(ctx.story, row);
  const close = closeOf(ctx.story, row);
  const owner = pendingOwn(ctx, row.by);
  if (owner && insertedPending(ctx.doc, row)) {
    if (inside.some((item) => othersPending(ctx.doc, item, owner.id)))
      throw new OrderedRefusal('others-in-row', "a row that holds another author's change");
    for (const item of [row, ...inside, close]) takeBack(ctx, ctx.story, item);
    takeBackNotesOf(ctx, inside);
    return owner.id;
  }

  if (!options.evenStruck && struckAlready(ctx, row.dels)) return null;

  // A cell holds one `w:cellIns`, `w:cellDel` or `w:cellMerge`: as the engine's tracked row
  // deletion, a suggestion refuses a row whose cell shows one of its own (a pending new or struck
  // column)
  const cells = ownCellEdges(inside);
  if (
    !ctx.direct &&
    cells.some((cell) => insertedPending(ctx.doc, cell) || struckAlready(ctx, cell.dels))
  ) {
    throw engineRefusal('deleteTableRow', 'unsupported-revision');
  }

  // The row's own cells are struck too (`w:cellDel`, as the engine marks them); everything in
  // the row goes with it through `within`
  const suggestion = suggestionFor(ctx);
  for (const item of [row, close, ...cells]) {
    relabel(ctx.doc, ctx.story, item, { dels: [...item.dels, suggestion.id] });
  }

  for (const item of inside) {
    holdWithin(ctx, ctx.story, item, suggestion.id);
  }

  return suggestion.id;
}

/** The edges of a row's own cells among the items inside it, not those of a table in a cell. */
function ownCellEdges(inside: readonly OrderedItem[]): OrderedItem[] {
  const edges: OrderedItem[] = [];
  let depth = 0;
  for (const item of inside) {
    if (item.token === 'tbl') {
      depth += item.kind === 'open' ? 1 : -1;
    } else if (item.token === 'tc' && depth === 0) {
      edges.push(item);
    }
  }

  return edges;
}

// ---- new tables ----

/**
 * Insert > Table before the caret's paragraph, as the engine builds it. Right after a table, the
 * engine first adds a paragraph with an inserted mark (two tables with nothing between them are
 * one table in Word), so the scratch part holds that table too.
 */
export function insertTable(
  ctx: Ctx,
  p: OrderedViewParagraph,
  rows: number,
  cols: number,
  columnWidth = 2000
): string {
  if (!Number.isSafeInteger(rows) || !Number.isSafeInteger(cols) || rows < 1 || cols < 1)
    throw new OrderedRefusal('invalid-edit', 'a table needs a row and a column');
  if (!Number.isSafeInteger(columnWidth) || columnWidth < 1)
    throw new OrderedRefusal('invalid-edit', 'a column needs a width');

  // The table goes before the paragraph, into every block container it is in
  refuseLockedBlocks(ctx, p.blocks);
  const first = p.items[0] ?? p.mark;
  const before = ctx.story.tree.prev(first);
  const view = viewPart(ctx);
  const paragraph = viewNode(view, p.id);
  const siblings = parentNodeOf(view, paragraph.id)?.children ?? [];
  const previous = siblings[siblings.findIndex((c) => c.id === paragraph.id) - 1];
  const table = previous && (previous.kind === 'table' || isWml(previous, 'tbl')) ? previous : null;
  const scratch = engineScratch(view, table ? [table, paragraph] : [paragraph], {
    op: 'insertTable',
    beforeParagraphId: paragraph.id,
    rows,
    cols,
    columnWidthTwips: columnWidth,
    revision: TEMPLATE_REVISION,
  });
  const added = scratchBlocks(scratch).filter((c) => c.id !== paragraph.id && c.id !== table?.id);
  const made = added[added.length - 1];
  const between = added.slice(0, -1);
  if (
    !made ||
    made.kind !== 'table' ||
    between.length > (table ? 1 : 0) ||
    between.some((c) => c.kind !== 'paragraph')
  )
    throw new OrderedRefusal('engine-refused', 'the engine added more than a table');
  const suggestion = suggestionFor(ctx);
  const items = added.flatMap((node) =>
    node.kind === 'textValue'
      ? []
      : templateItems(ctx, node, suggestion.id).map((item) => nestedIn(item, p.mark.within))
  );
  insertItems(ctx.doc, ctx.story, before, items);
  return suggestion.id;
}

// ---- columns and table properties ----

export { columnOp } from './table-columns.ts';
export { tableProperty } from './table-properties.ts';
