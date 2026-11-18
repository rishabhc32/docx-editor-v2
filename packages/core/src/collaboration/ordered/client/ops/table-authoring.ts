// `authorTable` (task 9.7): the editing API's table authoring (the `Word.Table` members), as one
// op.

import { findNode } from '../../../../store/package/ooxml-edit.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import {
  planInsertTable,
  planTableMutation,
  type AutomationTableMutation,
} from '../../../../store/store/table-authoring-plan.ts';
import { paragraphTextOf } from '../../../../store/store/tree-op-apply.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { readTwipsMeasure } from '../../../../store/units.ts';
import { gridColumnsOf } from '../../edits/table-nodes.ts';
import type { OrderedTableOp } from '../../edits/tables.ts';
import { isHidden } from '../../model/document.ts';
import type { OrderedItem, OrderedStory } from '../../model/types.ts';
import { attributeOf } from '../../ooxml-nodes.ts';
import { closeOf } from '../../placement.ts';
import {
  addressOf,
  paragraphOfMark,
  placeAt,
  viewText,
  type ListParagraph,
} from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { containerOf } from './containers.ts';
import { splitStep } from './paragraphs.ts';
import { chunks } from './text.ts';
import {
  translator,
  type OrderedChangeBody,
  type OrderedOpTranslator,
  type OrderedTranslateBatch,
} from './types.ts';

type AuthorTable = Extract<TreeDocOp, { readonly op: 'authorTable' }>;
type Insertion = Extract<AuthorTable['action'], { readonly kind: 'insert' }>;

/**
 * A whole table's deletion is a block deletion (`deleteBlock`): group 9.10's op, refused by its
 * name.
 */
const BLOCK_DELETION = () => new OrderedClientRefusal('ordered-unsupported-op', 'deleteBlock');

/** The editor's part of a story: the package the client last published. */
function shownPart(batch: OrderedTranslateBatch, story: string): OoxmlPart {
  const part = batch.shown()?.parts.get(story);
  if (!part) throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no part shown');
  return part;
}

/**
 * The engine's plan of the op on the editor's part (`planInsertTable`, `planTableMutation`),
 * checked: the client refuses, by the engine's own reason, what the editor alone refuses (a
 * tracked deletion of a whole table, a table with merged cells, a cell that holds more than one
 * paragraph). A plan that deletes the whole table is a block deletion.
 */
function plannedByEngine(part: OoxmlPart, op: AuthorTable): void {
  const reads = { part, root: part.root, rawText: (id: string) => paragraphTextOf(part, id) };
  const { action } = op;
  const plan =
    action.kind === 'insert'
      ? planInsertTable(
          reads,
          action.paragraphId,
          action.offset,
          action.rowCount,
          action.columnCount,
          action.values,
          op.revision
        )
      : planTableMutation(reads, action.tableId, action.mutation, op.revision);
  if (!plan.ok) {
    throw new OrderedClientRefusal(
      'ordered-refused',
      'engine-refused',
      `authorTable: ${plan.reason}`
    );
  }

  if (plan.ops.some((planned) => planned.op === 'deleteBlock')) {
    throw BLOCK_DELETION();
  }
}

/** A row of a table as the list shows it: its open item and its cells' open items. */
interface ShownRow {
  readonly open: OrderedItem;
  readonly cells: OrderedItem[];
}

/** A table's own rows and their cells as the list shows them (a nested table's are not). */
function shownRows(
  batch: OrderedTranslateBatch,
  story: OrderedStory,
  table: OrderedItem
): ShownRow[] {
  const rows: ShownRow[] = [];
  const close = closeOf(story, table, batch.doc);
  let depth = 0;
  for (let item = story.tree.next(table); item && item !== close; item = story.tree.next(item)) {
    if (item.token === 'tbl' && item.kind === 'open') {
      depth += 1;
    }

    if (item.token === 'tbl' && item.kind === 'close') {
      depth -= 1;
    }

    if (depth > 0 || item.kind !== 'open' || isHidden(batch.doc, item)) {
      continue;
    }

    if (item.token === 'tr') {
      rows.push({ open: item, cells: [] });
    } else if (item.token === 'tc') {
      rows[rows.length - 1]?.cells.push(item);
    }
  }

  return rows;
}

/** The one paragraph of a cell, as the list shows it. */
function cellParagraph(
  batch: OrderedTranslateBatch,
  story: OrderedStory,
  cell: OrderedItem | undefined
): ListParagraph {
  if (cell) {
    const close = closeOf(story, cell, batch.doc);
    for (let item = story.tree.next(cell); item && item !== close; item = story.tree.next(item))
      if (item.kind === 'mark' && !isHidden(batch.doc, item)) {
        const found = paragraphOfMark(batch.doc, story, item);
        if (found) return found;
      }
  }

  throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a cell with no paragraph');
}

/**
 * Write a value as the engine does: typed before the paragraph's old text, which then goes.
 * `read` reads the paragraph again once the typing is applied.
 */
function* written(read: () => ListParagraph, text: string): Generator<OrderedChangeBody> {
  const before = read();
  const old = viewText(before.view);
  if (old === text) return;
  const part = before.story.name;

  // Each piece goes before the same place, so the pieces keep their order
  const place = placeAt(before, 0);
  for (const piece of text.length > 0 ? chunks(text) : [])
    yield { kind: 'insert', part, place, text: piece };
  if (old.length === 0) return;
  const after = read();
  const from = placeAt(after, text.length);
  yield { kind: 'delete', part, from, to: placeAt(after, text.length + old.length) };
}

/**
 * Write the values into the table's cells from row `row` and column `column` on. A row, a column
 * or a table the op adds is found by its place once the changes before it are applied, since
 * every copy makes it with ids of its own.
 */
function* values(
  batch: OrderedTranslateBatch,
  story: OrderedStory,
  table: () => OrderedItem,
  given: readonly (readonly string[])[] | undefined,
  row = 0,
  column = 0
): Generator<OrderedChangeBody> {
  if (!given) return;
  const shown = batch.shown()?.parts.get(story.name);

  // Values change text inside the cells only: the rows and cells stay as read here
  const rows = shownRows(batch, story, table());
  for (const [r, line] of given.entries())
    for (const [c, text] of line.entries()) {
      const read = () => {
        const p = cellParagraph(batch, story, rows[row + r]?.cells[column + c]);

        // A paragraph the editor shows has its text checked once, as every op's
        return shown && findNode(shown, p.view.id) ? batch.paragraph(p.view.id) : p;
      };
      yield* written(read, text);
    }
}

/** A table, row or cell press over list containers: one `table-property` change. */
export const press = (story: OrderedStory, on: readonly OrderedItem[], op: OrderedTableOp) =>
  ({ kind: 'table-property', part: story.name, on: on.map(addressOf), op }) as const;

/** The widths of the table's grid as the editor shows it, in twips. */
function gridWidths(part: OoxmlPart, tableId: string): number[] {
  const table = findNode(part, tableId);
  if (!table || table.kind === 'textValue') return [];
  return gridColumnsOf(table).map((column) => readTwipsMeasure(attributeOf(column, 'w')) ?? 0);
}

/** An existing table the op changes: its story, and its open item and rows read again. */
interface ChangedTable {
  readonly batch: OrderedTranslateBatch;
  readonly story: OrderedStory;
  readonly tableId: string;
  readonly table: () => OrderedItem;
  readonly rows: () => ShownRow[];
}

/** The mutations whose kinds are among `K` (some share one shape: rows and columns). */
type Mutation<K extends AutomationTableMutation['kind']> = Extract<
  AutomationTableMutation,
  { readonly kind: K }
>;

/** The engine's vertical alignment values, as the table op names them. */
const VERTICAL_ALIGNMENT = { Top: 'top', Center: 'center', Bottom: 'bottom' } as const;

/** A change to an existing table, in the plan's order. */
function* mutated(
  batch: OrderedTranslateBatch,
  tableId: string,
  mutation: AutomationTableMutation
): Generator<OrderedChangeBody> {
  const { story } = containerOf(batch.doc, tableId, ['tbl']);
  const table = () => containerOf(batch.doc, tableId, ['tbl']).open;
  const changed: ChangedTable = {
    batch,
    story,
    tableId,
    table,
    rows: () => shownRows(batch, story, table()),
  };
  switch (mutation.kind) {
    case 'delete':
      // A block deletion: refused by name before the step (`plannedByEngine`)
      return;
    case 'properties':
      yield press(story, [table()], {
        op: 'setTableProperties',
        ...(mutation.styleId === undefined ? {} : { styleId: mutation.styleId }),
        ...(mutation.headerRowCount === undefined
          ? {}
          : { headerRowCount: mutation.headerRowCount }),
      });
      return;
    case 'values':
      yield* values(batch, story, table, mutation.values);
      return;
    case 'cell':
      yield* cellChanges(changed, mutation);
      return;
    case 'deleteRows':
      yield* rowsDeleted(changed, mutation);
      return;
    case 'deleteColumns':
      yield* columnsDeleted(changed, mutation);
      return;
    case 'insertRows':
    case 'addRows':
      yield* rowsAdded(changed, mutation);
      return;
    case 'addColumns':
      yield* columnsAdded(changed, mutation);
      return;
  }
}

/** One cell's value, fill, vertical alignment and column width. */
function* cellChanges(
  { batch, story, tableId, table, rows }: ChangedTable,
  mutation: Mutation<'cell'>
): Generator<OrderedChangeBody> {
  const cell = containerOf(batch.doc, mutation.cellId, ['tc']).open;
  const index = rows()
    .find((row) => row.cells.includes(cell))
    ?.cells.indexOf(cell);
  if (index === undefined || index < 0) {
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a cell of another table');
  }

  if (mutation.value !== undefined) {
    yield* written(
      () => batch.paragraph(cellParagraph(batch, story, cell).view.id),
      mutation.value
    );
  }

  if (mutation.shadingColor !== undefined) {
    const color = mutation.shadingColor.replace(/^#/, '');
    yield press(story, [cell], {
      op: 'setTableCellFill',
      color: color === 'auto' ? null : { kind: 'hex', value: color },
    });
  }

  if (mutation.verticalAlignment !== undefined) {
    yield press(story, [cell], {
      op: 'setTableCellVerticalAlignment',
      alignment: VERTICAL_ALIGNMENT[mutation.verticalAlignment],
    });
  }

  if (mutation.columnWidth !== undefined) {
    const widths = gridWidths(shownPart(batch, story.name), tableId);
    widths[index] = Math.round(mutation.columnWidth * 20);
    yield press(story, [table()], { op: 'setTableProperties', columnWidthsTwips: widths });
  }
}

/** Rows deleted from the last, as the engine deletes them (all of them is a block deletion). */
function* rowsDeleted(
  { story, rows }: ChangedTable,
  mutation: Mutation<'deleteRows' | 'deleteColumns'>
): Generator<OrderedChangeBody> {
  const deleted = rows()
    .slice(mutation.index, mutation.index + mutation.count)
    .reverse();
  for (const row of deleted) {
    yield { kind: 'delete-row', part: story.name, row: addressOf(row.open) };
  }
}

/** Columns deleted from the last, each by the first row's cell. */
function* columnsDeleted(
  { story, rows }: ChangedTable,
  mutation: Mutation<'deleteRows' | 'deleteColumns'>
): Generator<OrderedChangeBody> {
  for (let k = mutation.index + mutation.count - 1; k >= mutation.index; k -= 1) {
    const cell = rows()[0]?.cells[k];
    if (!cell) {
      throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no column');
    }

    yield { kind: 'delete-column', part: story.name, cell: addressOf(cell) };
  }
}

/** The index of the row new rows stand beside: the row named, or the table's first or last. */
function sourceRowIndex(
  batch: OrderedTranslateBatch,
  all: readonly ShownRow[],
  mutation: Mutation<'insertRows' | 'addRows' | 'addColumns'>
): number {
  if (mutation.kind === 'insertRows') {
    const named = containerOf(batch.doc, mutation.rowId, ['tr']).open;
    return all.findIndex((row) => row.open === named);
  }

  return mutation.location === 'start' ? 0 : all.length - 1;
}

/** New rows beside one row, then their values. */
function* rowsAdded(
  { batch, story, table, rows }: ChangedTable,
  mutation: Mutation<'insertRows' | 'addRows' | 'addColumns'>
): Generator<OrderedChangeBody> {
  const all = rows();
  const at = sourceRowIndex(batch, all, mutation);
  const source = all[at];
  if (!source) {
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no row');
  }

  const above = mutation.location === 'start' || mutation.location === 'before';

  // Every new row stands beside the same row, as the engine adds them
  for (let k = 0; k < mutation.count; k += 1) {
    yield {
      kind: 'insert-row',
      part: story.name,
      row: addressOf(source.open),
      where: above ? 'above' : 'below',
    };
  }

  yield* values(batch, story, table, mutation.values, at + (above ? 0 : 1));
}

/** New columns at the table's start or end, then their values. */
function* columnsAdded(
  { batch, story, tableId, table, rows }: ChangedTable,
  mutation: Mutation<'addRows' | 'addColumns'>
): Generator<OrderedChangeBody> {
  // The table's columns before any is added: the grid's, or the first row's cells
  const widths = gridWidths(shownPart(batch, story.name), tableId);
  const columns = widths.length > 0 ? widths.length : (rows()[0]?.cells.length ?? 0);
  const first = mutation.location === 'start';

  // Each new column at the table's edge, beside the column the last one made
  for (let k = 0; k < mutation.count; k += 1) {
    const cells = rows()[0]?.cells ?? [];
    const edge = first ? cells[0] : cells[cells.length - 1];
    if (!edge) {
      throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no column');
    }

    yield {
      kind: 'insert-column',
      part: story.name,
      cell: addressOf(edge),
      where: first ? 'left' : 'right',
    };
  }

  yield* values(batch, story, table, mutation.values, 0, first ? 0 : columns);
}

/** A table's open items in a story, by item space. */
function tablesOf(story: OrderedStory): Set<string> {
  const out = new Set<string>();
  for (const item of story.tree.items())
    if (item.kind === 'open' && item.token === 'tbl') out.add(item.space);
  return out;
}

/** A new table before a paragraph's offset, split there first, then its values. */
function* inserted(batch: OrderedTranslateBatch, action: Insertion): Generator<OrderedChangeBody> {
  if (action.offset > 0) yield* splitStep(batch, action.paragraphId, action.offset)();

  // The list keeps the paragraph's own mark on the tail: the table goes at its start
  const tail = batch.paragraph(action.paragraphId);
  const { story } = tail;
  const before = tablesOf(story);
  yield {
    kind: 'insert-table',
    part: story.name,
    place: placeAt(tail, 0),
    rows: action.rowCount,
    columns: action.columnCount,
    // The engine's width (`planInsertTable` in `table-authoring-plan.ts`): the page's text
    // width over the columns. Keep the two the same
    columnWidth: Math.max(120, Math.floor(9360 / action.columnCount)),
  };
  if (!action.values) return;
  let made: OrderedItem | undefined;
  for (const item of story.tree.items())
    if (item.kind === 'open' && item.token === 'tbl' && !before.has(item.space)) made = item;
  const table = made;
  if (!table) throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no new table');
  yield* values(batch, story, () => table, action.values);
}

/** `authorTable`: the engine's plan, checked at once, as the changes of its parts. */
export const authorTable: OrderedOpTranslator = translator('authorTable', (op, batch) => {
  const { action } = op;
  const story =
    action.kind === 'insert'
      ? batch.paragraph(action.paragraphId).story.name
      : containerOf(batch.doc, action.tableId, ['tbl']).story.name;
  plannedByEngine(shownPart(batch, story), op);
  return () =>
    action.kind === 'insert'
      ? inserted(batch, action)
      : mutated(batch, action.tableId, action.mutation);
});
