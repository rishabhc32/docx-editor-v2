// A decided column, written into its table (task 9.7).

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import { composeBlocks } from '../compose/blocks.ts';
import { ComposeContext } from '../compose/context.ts';
import type { RemovedColumns } from '../compose/table-columns.ts';
import { cellsOf, rowsOf } from '../edits/table-nodes.ts';
import { itemAt, statusOf, storyOfItem } from '../model/document.ts';
import { nextSeq, relabel } from '../model/mutations.ts';
import {
  tokenSpace,
  type OrderedDocument,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedPropertyTarget,
  type OrderedStory,
} from '../model/types.ts';
import { childNamed, isWml, withChildren } from '../ooxml-nodes.ts';
import type { OrderedIdMinter } from './ids.ts';

/** The placement children a removal and a join change, per container. */
const PLACEMENT = {
  tcPr: ['tcW', 'gridSpan'],
  trPr: ['gridBefore', 'gridAfter'],
} as const;

const TABLE_TOKENS = new Set(['tbl', 'tr', 'tc']);

/**
 * Write each column the decided suggestions removed into its table's items. Compose takes out
 * the cells a decision removed, gives their width to a neighbour and joins the grid, by the
 * engine's rules (`compose/table-columns.ts`). Rows and cells made after the decision are made
 * on that table, so the items must hold it too, as the engine's own table does once it decides:
 * else a new row's spans are read against the old grid, and the live list and its saved file
 * draw the table differently. Only stories with a table, row or cell the suggestions label are
 * drawn.
 */
export function settleRemovedColumns(
  doc: OrderedDocument,
  decided: readonly string[],
  ids: OrderedIdMinter
): void {
  const stories = new Set<OrderedStory>();
  for (const id of decided)
    for (const item of doc.labelled.get(id) ?? []) {
      if (item.kind !== 'open' || !TABLE_TOKENS.has(item.token ?? '')) continue;
      const story = storyOfItem(doc, item);
      if (story) stories.add(story);
    }

  for (const story of stories) {
    const ctx = new ComposeContext(doc, story, {});
    composeBlocks(ctx);
    for (const table of ctx.removedColumns) settleTable(doc, story, table, ids);
  }
}

/** The placement children of a container, as plain data for a comparison. */
function placementOf(props: OoxmlElement | undefined, names: readonly string[]): string {
  if (!props) return '[]';
  return JSON.stringify(
    props.children.flatMap((c) =>
      c.kind === 'textValue' || !names.includes(c.localName)
        ? []
        : [[c.localName, c.attributes.map((a) => `${a.localName}=${a.value}`).sort()]]
    )
  );
}

/** What every record of one decided table shares. */
interface TableWrite {
  readonly doc: OrderedDocument;
  readonly story: OrderedStory;
  readonly suggestion: string;
  readonly seq: number;
  readonly ids: OrderedIdMinter;
}

/** The open item of a table, row or cell node. */
function openOf(
  where: { readonly doc: OrderedDocument; readonly story: OrderedStory },
  id: string
): OrderedItem | null {
  return itemAt(where.story, { space: tokenSpace(where.doc.file, id), offset: 0 });
}

/** A record of the accepted deletion on a container's open item. */
function addRecord(
  write: TableWrite,
  open: OrderedItem,
  record: Omit<OrderedElementRecord, 'kind' | 'seq' | 'suggestion'>
): void {
  const { suggestion, seq } = write;
  relabel(write.doc, write.story, open, {
    props: [...open.props, { kind: 'element', suggestion, seq, ...record }],
  });
}

/** A copy of a node with new ids. */
function copied(write: TableWrite, node: OoxmlElement): OoxmlElement {
  // `cloneWithNewIds` keeps the node's kind, so the copy is an element as the node is
  return cloneWithNewIds(node, () => write.ids.node()) as OoxmlElement;
}

/**
 * Write the widths, spans and grid compose drew into a table's items, as records of the
 * accepted deletion, and mark the removed cells gone: compose then draws the same table from the
 * items alone.
 */
function settleTable(
  doc: OrderedDocument,
  story: OrderedStory,
  { before, after, cells }: RemovedColumns,
  ids: OrderedIdMinter
): void {
  // Only an accepted deletion is written: its decision is final. A rejected insertion, and a
  // table whose grid holds a pending change record (its join waits for that decision), stay
  // drawn from the items at every compose.
  if (cells.some((c) => statusOf(doc, c.suggestion) !== 'accepted')) {
    return;
  }

  const grid = childNamed(after, 'tblGrid');
  if (!grid || grid.children.some((c) => isWml(c, 'tblGridChange'))) {
    return;
  }

  const table = openOf({ doc, story }, before.id);
  if (!table) {
    return;
  }

  const suggestion = cells[0]!.suggestion;
  const write: TableWrite = { doc, story, suggestion, seq: nextSeq(doc), ids };
  recordGrid(write, table, before, after, grid);
  const oldRows = new Map(rowsOf(before).map((row) => [row.id, row]));
  for (const row of rowsOf(after)) {
    recordPlacement(write, openOf(write, row.id), oldRows.get(row.id), row, 'trPr');
  }

  const oldCells = cellsOf(before);
  for (const [id, cell] of cellsOf(after)) {
    recordPlacement(write, openOf(write, id), oldCells.get(id), cell, 'tcPr');
  }

  dropRemovedCells(write, cells);
}

/**
 * The grid, with the cells that show now, when the columns changed: a later decision on one of
 * them still gives its width back (`holdsCell`).
 */
function recordGrid(
  write: TableWrite,
  table: OrderedItem,
  before: OoxmlElement,
  after: OoxmlElement,
  grid: OoxmlElement
): void {
  const columns = (t: OoxmlElement) => placementOf(childNamed(t, 'tblGrid'), ['gridCol']);
  if (columns(before) === columns(after)) {
    return;
  }

  const tableCells: string[] = [];
  for (const id of cellsOf(after).keys()) {
    const open = openOf(write, id);
    if (open) {
      tableCells.push(open.space);
    }
  }

  const gridCols = grid.children.filter((c) => isWml(c, 'gridCol'));
  addRecord(write, table, {
    target: 'tblGrid',
    element: copied(write, withChildren(grid, gridCols)),
    tableCells,
  });
}

/** A row's or a cell's placement children, when the decision changed them. */
function recordPlacement(
  write: TableWrite,
  open: OrderedItem | null,
  was: OoxmlElement | undefined,
  now: OoxmlElement,
  target: keyof typeof PLACEMENT
): void {
  const names = PLACEMENT[target];
  const nowProps = childNamed(now, target);
  if (
    !open ||
    placementOf(was && childNamed(was, target), names) === placementOf(nowProps, names)
  ) {
    return;
  }

  const given: OoxmlNode[] = (nowProps?.children ?? []).filter(
    (c) => c.kind !== 'textValue' && (names as readonly string[]).includes(c.localName)
  );
  addRecord(write, open, {
    target: target as OrderedPropertyTarget,
    // Merged over the container's own children: the others (a pending fill) stay theirs
    element: nowProps ? copied(write, withChildren(nowProps, given)) : null,
    merge: names.map((name) => `${WML_NAMESPACE_URI} ${name}`),
  });
}

/** The removed cells show nowhere now, and compose takes nothing out again. */
function dropRemovedCells(write: TableWrite, cells: RemovedColumns['cells']): void {
  for (const removed of cells) {
    const open = openOf(write, removed.id);
    if (open && !open.gone) {
      relabel(write.doc, write.story, open, { gone: true });
    }
  }
}
