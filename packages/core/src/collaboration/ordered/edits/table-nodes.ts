// Small helpers the table edits share: a table's rows and cells in the drawn view, their open
// items in the list, and properties compared and copied without their revision markers.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { itemAt } from '../model/document.ts';
import { nextSeq, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import {
  tokenSpace,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedPropertyTarget,
} from '../model/types.ts';
import { attributeOf, childNamed, isWml, withChildren } from '../ooxml-nodes.ts';
import { withNewIds, type Ctx } from './context.ts';
import { withPropertyRecord } from './text.ts';

/** The innermost block container of a kind around a paragraph. */
export function blockOf(p: OrderedViewParagraph, token: 'tbl' | 'tr' | 'tc'): OrderedItem {
  const found = [...p.blocks].reverse().find((block) => block.token === token);
  if (!found) throw new OrderedRefusal('not-in-table', 'the place is not in a table');
  return found;
}

const REVISION_MARKERS = new Set([
  'cellIns',
  'cellDel',
  'cellMerge',
  'ins',
  'del',
  'tcPrChange',
  'trPrChange',
  'tblPrChange',
  'tblGridChange',
]);

const isRevisionMarker = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' && REVISION_MARKERS.has(node.localName);

/** A properties element without revision markers, as plain data for a comparison. */
function plain(node: OoxmlNode | undefined): unknown {
  if (!node) return null;
  if (node.kind === 'textValue') return node.value;
  return [
    node.localName,
    node.attributes.map((a) => `${a.localName}=${a.value}`).sort(),
    node.children.filter((c) => !isRevisionMarker(c)).map(plain),
  ];
}

/** Whether two properties elements say the same, revision markers aside. */
export const same = (a: OoxmlElement | undefined, b: OoxmlElement | undefined): boolean =>
  JSON.stringify(plain(a)) === JSON.stringify(plain(b));

/**
 * The children (`<namespace URI> <local name>`) whose value differs between two properties
 * elements, revision markers aside: what a record that sets only them changes (`merge`).
 */
export function changedKeys(was: OoxmlElement | undefined, now: OoxmlElement): string[] {
  const keyed = (element: OoxmlElement | undefined) => {
    const out = new Map<string, string>();
    for (const c of element?.children ?? [])
      if (c.kind !== 'textValue' && !isRevisionMarker(c))
        out.set(`${c.namespaceUri} ${c.localName}`, JSON.stringify(plain(c)));
    return out;
  };
  const [before, after] = [keyed(was), keyed(now)];
  return [...new Set([...before.keys(), ...after.keys()])].filter(
    (key) => before.get(key) !== after.get(key)
  );
}

/** A copy of a properties element without its revision markers, with the edit's ids. */
export function stripped(ctx: Ctx, element: OoxmlElement): OoxmlElement {
  return withNewIds(
    ctx,
    withChildren(
      element,
      element.children.filter((c) => !isRevisionMarker(c))
    )
  );
}

export const isRow = (node: OoxmlNode): boolean => node.kind === 'tableRow' || isWml(node, 'tr');
export const isCell = (node: OoxmlNode): boolean => node.kind === 'tableCell' || isWml(node, 'tc');

/** A table's own rows (not a nested table's). */
export function rowsOf(table: OoxmlElement): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const row of table.children) if (row.kind !== 'textValue' && isRow(row)) out.push(row);
  return out;
}

/** A row's own cells. */
export function rowCells(row: OoxmlElement): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const cell of row.children) if (cell.kind !== 'textValue' && isCell(cell)) out.push(cell);
  return out;
}

/** A table's own cells by node id. */
export function cellsOf(table: OoxmlElement): Map<string, OoxmlElement> {
  const out = new Map<string, OoxmlElement>();
  for (const row of rowsOf(table)) for (const cell of rowCells(row)) out.set(cell.id, cell);
  return out;
}

/** The open item of a container node of the view. */
export function openOf(ctx: Ctx, node: OoxmlElement): OrderedItem | null {
  return itemAt(ctx.story, { space: tokenSpace(ctx.doc.file, node.id), offset: 0 });
}

/**
 * The spaces of a drawn table's own cells: what a record of the table holds when it is taken
 * (`tableCells`, read by `compose/table-columns.ts`).
 */
export function tableCellSpaces(ctx: Ctx, table: OoxmlElement): string[] {
  const out: string[] = [];
  for (const cell of cellsOf(table).values()) {
    const open = openOf(ctx, cell);
    if (open) out.push(open.space);
  }

  return out;
}

/**
 * A record of the suggestion that gives a table's container `element`, copied without its
 * revision markers, and holds the cells the table showed (`tableCells`).
 */
export function elementRecord(
  ctx: Ctx,
  suggestion: string,
  target: OrderedPropertyTarget,
  element: OoxmlElement,
  tableCells: readonly string[],
  flags: Pick<OrderedElementRecord, 'untracked' | 'recordUnchanged' | 'merge'> = {}
): OrderedElementRecord {
  return {
    kind: 'element',
    suggestion,
    seq: nextSeq(ctx.doc),
    target,
    element: stripped(ctx, element),
    tableCells,
    ...flags,
  };
}

/** Add a property record to a container's open item (policy 2 for a suggestion). */
export function withRecord(ctx: Ctx, open: OrderedItem, record: OrderedElementRecord): void {
  relabel(ctx.doc, ctx.story, open, { props: withPropertyRecord(ctx, open, record) });
}

/** A cell's span over grid columns (`w:gridSpan`), 1 when absent. */
function gridSpanOf(cell: OoxmlElement): number {
  const tcPr = childNamed(cell, 'tcPr');
  const span = tcPr ? childNamed(tcPr, 'gridSpan') : undefined;
  const value = span ? Number.parseInt(attributeOf(span, 'val') ?? '1', 10) : 1;
  return Number.isSafeInteger(value) && value > 0 ? value : 1;
}

/** A table's `w:gridCol` elements, in order; empty when it has no grid. */
export function gridColumnsOf(table: OoxmlElement): OoxmlElement[] {
  const out: OoxmlElement[] = [];
  for (const column of childNamed(table, 'tblGrid')?.children ?? [])
    if (column.kind !== 'textValue' && isWml(column, 'gridCol')) out.push(column);
  return out;
}

/** The grid columns a cell covers: the index it starts at and its span; null when not found. */
export function cellColumns(
  table: OoxmlElement,
  cellId: string
): { readonly start: number; readonly span: number } | null {
  for (const row of rowsOf(table)) {
    let index = 0;
    for (const cell of rowCells(row)) {
      const span = gridSpanOf(cell);
      if (cell.id === cellId) return { start: index, span };
      index += span;
    }
  }

  return null;
}

/** The `w:gridCol` a cell starts at. */
export function gridColumnOfCell(table: OoxmlElement, cellId: string): string | null {
  const at = cellColumns(table, cellId);
  return at === null ? null : (gridColumnsOf(table)[at.start]?.id ?? null);
}
