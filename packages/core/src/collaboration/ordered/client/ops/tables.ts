// Table ops: table, row and cell properties as the engine's own table op (design D6), rows and
// a new table (task 6.1); columns, column widths, the right edge and table design (group 9.7).

import { findNode } from '../../../../store/package/ooxml-edit.ts';
import type { OoxmlElement } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { cellColumns, gridColumnsOf, rowCells, rowsOf } from '../../edits/table-nodes.ts';
import type { OrderedTableOp } from '../../edits/tables.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { addressOf, placeAt } from '../paragraphs.ts';
import { authorTable, press } from './table-authoring.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { containerOf, type ListContainer } from './containers.ts';
import {
  PARTY_SUGGESTS,
  translator,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedOpTranslator,
  type OrderedTranslateBatch,
} from './types.ts';

/** The cells of one table, as one selection of their open items. */
function cellsOf(
  doc: OrderedDocument,
  tableId: string,
  cellIds: readonly string[]
): ListContainer[] {
  const table = containerOf(doc, tableId, ['tbl']);
  return cellIds.map((id) => {
    const cell = containerOf(doc, id, ['tc']);
    if (cell.story !== table.story)
      throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a cell of another story');
    return cell;
  });
}

/** The table as the editor shows it: the package the client last published. */
function shownTable(batch: OrderedTranslateBatch, tableId: string): OoxmlElement {
  const table = containerOf(batch.doc, tableId, ['tbl']);
  const part = batch.shown()?.parts.get(table.story.name);
  const node = part ? findNode(part, tableId) : null;
  if (!node || node.kind === 'textValue')
    throw new OrderedClientRefusal(
      'ordered-refused',
      'out-of-step',
      `no table ${tableId.slice(0, 160)}`
    );
  return node;
}

/**
 * The cell at a grid column's edge, as the editor shows the table: in the first row that has
 * one, the cell that starts at the column (`start`) or ends at it (`end`). A column op names the
 * cell that starts there; the change names the cell, which keeps its address while columns come
 * and go, and the model reads the grid column it starts at again where it applies. A resize names
 * the cell whose right edge moves, the one that ends there.
 */
function cellAtGridEdge(
  batch: OrderedTranslateBatch,
  tableId: string,
  gridColumnId: string,
  edge: 'start' | 'end'
): string {
  const table = shownTable(batch, tableId);
  const index = gridColumnsOf(table).findIndex((column) => column.id === gridColumnId);
  for (const row of index < 0 ? [] : rowsOf(table)) {
    for (const cell of rowCells(row)) {
      const at = cellColumns(table, cell.id);
      if (at && (edge === 'start' ? at.start : at.start + at.span - 1) === index) {
        return cell.id;
      }
    }
  }

  const missing = edge === 'start' ? 'no cell at' : 'no cell ends at';
  throw new OrderedClientRefusal(
    'ordered-refused',
    'out-of-step',
    `${missing} grid column ${gridColumnId.slice(0, 160)}`
  );
}

/** A column resize as one `table-property` change on the cell whose right edge moves. */
const resizeChange = (kind: 'setTableColumnWidths' | 'setTableRightEdgeWidth') =>
  translator(kind, (op, batch) => {
    let cellId: string;
    let tableOp: OrderedTableOp;
    if (op.op === 'setTableColumnWidths') {
      const table = shownTable(batch, op.tableId);
      const ids = gridColumnsOf(table).map((column) => column.id);

      // The engine moves the edge between two neighbours, in document order
      if (ids.indexOf(op.rightGridColumnId) !== ids.indexOf(op.leftGridColumnId) + 1)
        throw new OrderedClientRefusal(
          'ordered-refused',
          'out-of-step',
          'columns not side by side'
        );
      cellId = cellAtGridEdge(batch, op.tableId, op.leftGridColumnId, 'end');
      tableOp = {
        op: op.op,
        leftWidthTwips: op.leftWidthTwips,
        rightWidthTwips: op.rightWidthTwips,
      };
    } else {
      cellId = cellAtGridEdge(batch, op.tableId, op.gridColumnId, 'end');
      tableOp = {
        op: op.op,
        columnWidthTwips: op.columnWidthTwips,
        tableWidthTwips: op.tableWidthTwips,
      };
    }

    return () => {
      const cell = containerOf(batch.doc, cellId, ['tc']);
      return [press(cell.story, [cell.open], tableOp)];
    };
  });

/** A column op as one change on the cell it names. */
const columnChange = (kind: 'insertTableColumn' | 'deleteTableColumn'): OrderedOpTranslator =>
  translator(kind, (op, batch) => {
    // The editor names a grid column, or the caret's cell in a table with no grid
    const cellId =
      'gridColumnId' in op
        ? cellAtGridEdge(batch, op.tableId, op.gridColumnId, 'start')
        : op.referenceCellId;
    return () => {
      const cell = containerOf(batch.doc, cellId, ['tc']);
      const part = cell.story.name;
      return op.op === 'insertTableColumn'
        ? [{ kind: 'insert-column', part, cell: addressOf(cell.open), where: op.where }]
        : [{ kind: 'delete-column', part, cell: addressOf(cell.open) }];
    };
  });

type TableOpKind =
  | 'setTableProperties'
  | 'setTableRowHeight'
  | 'setTableCellFill'
  | 'setTableCellVerticalAlignment'
  | 'setTableCellBorders';

/** A table op the change carries, with its node ids and attribution left out. */
function tableOpOf(op: Extract<TreeDocOp, { readonly op: TableOpKind }>): OrderedTableOp {
  switch (op.op) {
    case 'setTableProperties':
      return {
        op: op.op,
        ...(op.styleId === undefined ? {} : { styleId: op.styleId }),
        ...(op.headerRowCount === undefined ? {} : { headerRowCount: op.headerRowCount }),
        ...(op.columnWidthsTwips === undefined ? {} : { columnWidthsTwips: op.columnWidthsTwips }),
        ...(op.justification === undefined ? {} : { justification: op.justification }),
        ...(op.preferredWidth === undefined ? {} : { preferredWidth: op.preferredWidth }),
      };
    case 'setTableRowHeight':
      return { op: op.op, heightTwips: op.heightTwips };
    case 'setTableCellFill':
      return { op: op.op, color: op.color };
    case 'setTableCellVerticalAlignment':
      return { op: op.op, alignment: op.alignment };
    case 'setTableCellBorders':
      return 'spec' in op
        ? { op: op.op, scope: op.scope, spec: op.spec }
        : { op: op.op, scope: 'none', target: op.target };
  }
}

/** The containers a table op names: the table, the row, or the cells of the table. */
function tableTargets(
  doc: OrderedDocument,
  op: Extract<TreeDocOp, { readonly op: TableOpKind }>
): ListContainer[] {
  switch (op.op) {
    case 'setTableProperties':
      return [containerOf(doc, op.tableId, ['tbl'])];
    case 'setTableRowHeight':
      return [containerOf(doc, op.rowId, ['tr'])];
    default:
      return cellsOf(doc, op.tableId, op.cellIds);
  }
}

/** A table, row or cell press: one `table-property` change over its containers. */
const tablePress = (kind: TableOpKind): OrderedOpTranslator =>
  translator(kind, (op, batch) => {
    const tableOp = tableOpOf(op);
    return () => {
      const on = tableTargets(batch.doc, op);
      const first = on[0];
      if (!first) {
        return [];
      }

      const opens = on.map((container) => container.open);
      return [press(first.story, opens, tableOp)];
    };
  });

const PRESS_NOTE = "The engine's own table op over the containers it names (design D6).";

export const TABLE_OPS: OrderedOpTable<
  | TableOpKind
  | 'insertTableRow'
  | 'deleteTableRow'
  | 'insertTable'
  | 'insertTableColumn'
  | 'deleteTableColumn'
  | 'setTableColumnWidths'
  | 'setTableRightEdgeWidth'
  | 'authorTable'
> = {
  setTableProperties: {
    status: 'translated',
    group: '6.1',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: `A table's style, header rows, widths, alignment. ${PRESS_NOTE}`,
    translate: tablePress('setTableProperties'),
  },
  setTableRowHeight: {
    status: 'translated',
    group: '6.1',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: `A row's height. ${PRESS_NOTE}`,
    translate: tablePress('setTableRowHeight'),
  },
  setTableCellFill: {
    status: 'translated',
    group: '6.1',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: `Cell shading over a selection of cells. ${PRESS_NOTE}`,
    translate: tablePress('setTableCellFill'),
  },
  setTableCellVerticalAlignment: {
    status: 'translated',
    group: '6.1',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: `Cell vertical alignment. ${PRESS_NOTE}`,
    translate: tablePress('setTableCellVerticalAlignment'),
  },
  setTableCellBorders: {
    status: 'translated',
    group: '6.1',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: `Cell borders, read over the selection's outside and inside edges. ${PRESS_NOTE}`,
    translate: tablePress('setTableCellBorders'),
  },
  insertTableRow: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-row'],
    party: PARTY_SUGGESTS,
    note: 'A row above or below a row.',
    translate: translator('insertTableRow', (op, batch) => () => {
      const row = containerOf(batch.doc, op.rowId, ['tr']);
      const part = row.story.name;
      return [{ kind: 'insert-row', part, row: addressOf(row.open), where: op.where }];
    }),
  },
  deleteTableRow: {
    status: 'translated',
    group: '6.1',
    changes: ['delete-row'],
    party: PARTY_SUGGESTS,
    note: '`referenceCellId` only places the caret after the deletion.',
    translate: translator('deleteTableRow', (op, batch) => () => {
      const row = containerOf(batch.doc, op.rowId, ['tr']);
      return [{ kind: 'delete-row', part: row.story.name, row: addressOf(row.open) }];
    }),
  },
  insertTable: {
    status: 'translated',
    group: '6.1',
    changes: ['insert-table'],
    party: PARTY_SUGGESTS,
    note: 'A new table before a paragraph, every column one width.',
    translate: translator('insertTable', (op, batch) => () => {
      const p = batch.paragraph(op.beforeParagraphId);
      return [
        {
          kind: 'insert-table',
          part: p.story.name,
          place: placeAt(p, 0),
          rows: op.rows,
          columns: op.cols,
          columnWidth: op.columnWidthTwips,
        },
      ];
    }),
  },
  insertTableColumn: {
    status: 'translated',
    group: '9.7',
    changes: ['insert-column'],
    party: PARTY_SUGGESTS,
    note: "Tracked: `w:cellIns` (§17.13.5.4) on each new cell. The neighbour lends half its width with no record, so a reject hands it back by Word's rule for a cell revision (`rebuildRevisionTable`). The op names a grid column; the change a cell of it.",
    translate: columnChange('insertTableColumn'),
  },
  deleteTableColumn: {
    status: 'translated',
    group: '9.7',
    changes: ['delete-column'],
    party: PARTY_SUGGESTS,
    note: 'Tracked: `w:cellDel` (§17.13.5.1) on each cell of the column; accepted, the neighbour takes its width by the same rule. One op per selected column, one suggestion.',
    translate: columnChange('deleteTableColumn'),
  },
  setTableColumnWidths: {
    status: 'translated',
    group: '9.7',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: 'Dragging the edge between two columns: their widths, fixed layout, tracked as Word records widths (`w:tblGridChange` with the old grid, `w:tblPrChange`, `w:tcPrChange` on every cell of each row). The change names the cell whose right edge moves.',
    translate: resizeChange('setTableColumnWidths'),
  },
  setTableRightEdgeWidth: {
    status: 'translated',
    group: '9.7',
    changes: ['table-property'],
    party: PARTY_SUGGESTS,
    note: "Dragging the table's right edge: the last column and `w:tblW`, recorded as a column drag is. The change names a cell that ends at the last column.",
    translate: resizeChange('setTableRightEdgeWidth'),
  },
  authorTable: {
    status: 'translated',
    group: '9.7',
    changes: [
      'split',
      'insert-table',
      'insert-row',
      'delete-row',
      'insert-column',
      'delete-column',
      'table-property',
      'insert',
      'delete',
    ],
    party: PARTY_SUGGESTS,
    note: "The editing API's table authoring (`Word.Table`): the engine's plan, refused where the editor alone refuses it (a tracked deletion of a whole table: `tracked-table-deletion-unsupported`), else the changes of its parts, tracked table and cell properties and columns included. A whole table's deletion is `deleteBlock` (9.10).",
    translate: authorTable,
  },
};

/** Options of the table ops that change their meaning. */
export const TABLE_OPTIONS: readonly OrderedOpOption[] = [];
