// Cells a column decision removed from a row that still shows (task 9.7).

import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { resolveRevisions } from '../../../store/store/tree-op-revisions.ts';
import { compactRevisionGrid } from '../../../store/store/revision-table-rebuild.ts';
import { cellsCoveringColumn } from '../../../store/store/tree-op-table-merged-columns.ts';
import { readEditableTableTopology } from '../../../store/store/tree-op-table-topology.ts';
import { build, revisionAttributes } from '../../../store/store/tree-op-tracked.ts';
import { mintedChange } from '../change/ids.ts';
import { itemAt, statusOf } from '../model/document.ts';
import {
  loadedSpace,
  tokenSpace,
  type OrderedElementRecord,
  type OrderedItem,
  type OrderedPropertyTarget,
} from '../model/types.ts';
import { attributeOf, childNamed, isWml, withChildren, withId } from '../ooxml-nodes.ts';
import type { ComposeContext } from './context.ts';
import { withStoryBlocks } from './paragraph-ops.ts';
import { acceptedAndPending, containerPropertiesAfter } from './properties.ts';

/** A cell a column decision removed, drawn as the engine holds it before the decision. */
export interface RemovedCell {
  /** The cell's node id in the drawn table. */
  readonly id: string;
  /** `direct`: an organiser's direct column; `tracked`: a suggestion's. */
  readonly how: 'tracked' | 'direct';
  readonly suggestion: string;
  /** The shown grid was taken while the cell showed (a tie-break; the count decides). */
  readonly held: boolean;
}

/**
 * A table whose removed cells one drawing took out: the table as drawn with them (`before`)
 * and without them (`after`). A decision writes the difference into the table's items
 * (`edits/table-settle.ts`).
 */
export interface RemovedColumns {
  readonly before: OoxmlElement;
  readonly after: OoxmlElement;
  readonly cells: readonly RemovedCell[];
}

export const isRow = (n: OoxmlNode): boolean => n.kind === 'tableRow' || isWml(n, 'tr');
export const isCell = (n: OoxmlNode): boolean => n.kind === 'tableCell' || isWml(n, 'tc');

/** The change that made an item (its node id is minted from it, `n:<change>.<k>`), or `file`. */
function creation(item: OrderedItem): string {
  return mintedChange(loadedSpace(item.space)?.node ?? '', 'n') ?? 'file';
}

/**
 * The element records of one target on a container, in the order compose applies them: the
 * accepted ones, then the pending ones (as `containerPropertiesAfter` reads them).
 */
function appliedRecords(
  ctx: ComposeContext,
  open: OrderedItem,
  target: OrderedPropertyTarget
): OrderedElementRecord[] {
  const records = open.props.filter(
    (r): r is OrderedElementRecord => r.kind === 'element' && r.target === target
  );
  const { accepted, pending } = acceptedAndPending(ctx, records);
  return [...accepted, ...pending];
}

/** The record whose properties a target shows, or null for the container's own (the file's). */
export function shownRecord(
  ctx: ComposeContext,
  open: OrderedItem,
  target: OrderedPropertyTarget
): OrderedElementRecord | null {
  const applied = appliedRecords(ctx, open, target);
  return applied[applied.length - 1] ?? null;
}

/**
 * Whether properties were taken while a cell showed: a record's own list, or, for the
 * properties an item was made with (null), whether the cell was made with it. These are the
 * widths that hold the cell's column: every table record (`w:tblGrid`, `w:tcPr`) notes the cells
 * the table showed when it was taken (`tableCells`); the file's own properties hold the cells
 * the file had. A record taken after a column was removed already holds its widths, so that
 * column is not removed again.
 */
export function holdsCell(
  record: OrderedElementRecord | null,
  owner: OrderedItem,
  cell: OrderedItem
): boolean {
  if (record?.tableCells !== undefined) return record.tableCells.includes(cell.space);
  return creation(owner) === creation(cell);
}

/**
 * Why a cell that shows nowhere was removed by a decision on its column, while the table's
 * widths still hold that column; null otherwise (its row or table went, or the widths were
 * taken after it went). Only called for a cell of a row that shows.
 */
export function removedColumnCell(
  ctx: ComposeContext,
  open: OrderedItem,
  table: OrderedItem
): Omit<RemovedCell, 'id'> | null {
  if (open.kind !== 'open' || open.token !== 'tc' || open.gone) return null;

  // A container around it went: not a column decision
  if (ctx.containerWent(open)) return null;
  const held = holdsCell(shownRecord(ctx, table, 'tblGrid'), table, open);
  const removedBy = (suggestion: string): Omit<RemovedCell, 'id'> => ({
    held,
    how: ctx.doc.suggestions.get(suggestion)?.direct === true ? 'direct' : 'tracked',
    suggestion,
  });
  if (open.by !== null && (open.retracted || ctx.rejectedOrDrawnSo(open.by)))
    return removedBy(open.by);
  if (open.retracted) return null;
  const deleter = open.dels.find((id) => statusOf(ctx.doc, id) === 'accepted');
  return deleter === undefined ? null : removedBy(deleter);
}

/** The removed cell as the engine holds it before the decision: its properties and a paragraph. */
export function removedCellStub(ctx: ComposeContext, open: OrderedItem): OoxmlElement | null {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') return null;
  const props = containerPropertiesAfter(ctx, open, shell.children);
  const paragraph = build(ctx.mint(), 'paragraph', 'p', [], []);
  return withChildren(shell, [...props.filter((c) => isWml(c, 'tcPr')), paragraph]);
}

/** Every `w:id` value in a node, to mint marker ids that meet none of them. */
function revisionIdsIn(node: OoxmlNode, out: Set<string>, depth = 0): void {
  if (node.kind === 'textValue' || depth > 64) return;
  const id = attributeOf(node, 'id');
  if (id !== undefined) out.add(id);
  for (const child of node.children) revisionIdsIn(child, out, depth + 1);
}

/** Every node id in a node. */
function nodeIdsIn(node: OoxmlNode, out: Set<string>, depth = 0): void {
  out.add(node.id);
  if (node.kind === 'textValue' || depth > 64) return;
  for (const child of node.children) nodeIdsIn(child, out, depth + 1);
}

/** The node with every id the engine minted (one not in `known`) replaced by compose's own. */
function withComposeIds(ctx: ComposeContext, node: OoxmlNode, known: Set<string>): OoxmlNode {
  if (node.kind === 'textValue') return known.has(node.id) ? node : { ...node, id: ctx.mint() };
  const children = node.children.map((child) => withComposeIds(ctx, child, known));
  const same = children.every((child, k) => child === node.children[k]);
  const element = same ? node : withChildren(node, children);
  return known.has(node.id) ? element : withId(element, ctx.mint());
}

/** A cell with a marker added to its properties (`w:tcPr` leads a cell). */
function withCellMarker(cell: OoxmlElement, marker: OoxmlElement, mint: () => string) {
  const tcPr = cell.children.find((c) => isWml(c, 'tcPr'));
  if (!tcPr || tcPr.kind === 'textValue')
    return withChildren(cell, [build(mint(), 'generic', 'tcPr', [], [marker]), ...cell.children]);
  const children = [...tcPr.children];
  const record = children.findIndex((c) => c.kind !== 'textValue' && c.localName === 'tcPrChange');
  children.splice(record < 0 ? children.length : record, 0, marker);
  return withChildren(
    cell,
    cell.children.map((c) => (c === tcPr ? withChildren(tcPr, children) : c))
  );
}

/** The table with each of its own cells (not a nested table's) mapped. */
function withTableCells(
  table: OoxmlElement,
  map: (cell: OoxmlElement) => OoxmlElement
): OoxmlElement {
  let changed = false;
  const rows = table.children.map((row) => {
    if (row.kind === 'textValue' || !isRow(row)) return row;
    const cells = row.children.map((cell) =>
      cell.kind === 'textValue' || !isCell(cell) ? cell : map(cell)
    );
    if (cells.every((cell, k) => cell === row.children[k])) return row;
    changed = true;
    return withChildren(row, cells);
  });
  return changed ? withChildren(table, rows) : table;
}

/** The table's own cells, by node id. */
function cellsById(table: OoxmlElement): Map<string, OoxmlElement> {
  const out = new Map<string, OoxmlElement>();
  for (const row of table.children)
    if (row.kind !== 'textValue' && isRow(row))
      for (const cell of row.children)
        if (cell.kind !== 'textValue' && isCell(cell)) out.set(cell.id, cell);
  return out;
}

/** The table without the cells `ids` names (its own rows only). */
function withoutCells(table: OoxmlElement, ids: ReadonlySet<string>): OoxmlElement {
  return withChildren(
    table,
    table.children.map((row) =>
      row.kind === 'textValue' || !isRow(row) || !row.children.some((c) => ids.has(c.id))
        ? row
        : withChildren(
            row,
            row.children.filter((c) => !ids.has(c.id))
          )
    )
  );
}

/**
 * The direct removed cells that fill whole grid columns: those grid columns' ids. An organiser's
 * direct column follows the editor's untracked ops: its grid column goes with no width given to
 * anyone, as the engine's `deleteTableColumn` removes it.
 */
function wholeDirectColumns(
  table: OoxmlElement,
  root: OoxmlElement,
  direct: ReadonlySet<string>
): string[] {
  const read = readEditableTableTopology(root, table.id);
  if (!read.ok || !read.topology.grid) return [];
  const { topology } = read;
  const { gridColumns, rows } = topology;
  if (!topology.hasMerge) {
    const out: string[] = [];
    gridColumns.forEach((column, index) => {
      if (rows.length > 0 && rows.every(({ cells }) => direct.has(cells[index]?.id ?? '')))
        out.push(column.id);
    });
    return out;
  }

  // With merged cells, the engine deletes the cell of each row that covers the column (Word's
  // rule), and keeps the ids of the columns left of it: from the last column, each covering set
  // once
  const out: string[] = [];
  const taken = new Set<string>();
  for (const column of [...gridColumns].reverse()) {
    const covering = (cellsCoveringColumn(topology, column.id) ?? []).filter((c) => c !== null);
    if (covering.length === 0 || !covering.every((c) => direct.has(c.id))) continue;
    if (covering.every((c) => taken.has(c.id))) continue;
    for (const c of covering) taken.add(c.id);
    out.push(column.id);
  }

  return out;
}

/** The survivors of the rows a decision removes cells from, as the engine holds them before it. */
interface Survivors {
  /** The cells whose shown `w:tcPr` the decision must not widen. */
  readonly kept: ReadonlySet<string>;
  /** The `w:tcW` a removed column lent each cell it took width from, by the cell's node id. */
  readonly lent: ReadonlyMap<string, OoxmlElement>;
}

/**
 * The survivors of each row with a removed cell, by how the engine holds them before the
 * decision. A cell whose shown `w:tcPr` was taken while the removed cell showed is widened by
 * the engine's rule. A cell a removed column took width from, under that column's record, shows
 * the width it lent (`lent`): Word writes no record of it, so the cell shows its width from before
 * the column, where the engine holds it lent until the decision gives the removed cell's width
 * back. Any other cell whose shown `w:tcPr` was taken when the removed cell did not show (already
 * removed, or not added yet) holds no width of it (`kept`). A column's record names the cell it
 * took width from, so every other cell's own properties hold that column's cell.
 */
function survivorsOf(
  ctx: ComposeContext,
  drawn: OoxmlElement,
  removed: readonly RemovedCell[]
): Survivors {
  const columns = new Map(removed.map((r) => [r.id, r.suggestion]));
  const suggestions = new Set(columns.values());
  const kept = new Set<string>();
  const lent = new Map<string, OoxmlElement>();
  for (const row of drawn.children) {
    if (row.kind === 'textValue' || !isRow(row)) {
      continue;
    }

    const cells = row.children.flatMap((c) => {
      const space = tokenSpace(ctx.doc.file, c.id);
      const open =
        c.kind !== 'textValue' && isCell(c) ? itemAt(ctx.story, { space, offset: 0 }) : null;
      return open ? [{ id: c.id, open, lends: lendingsOf(ctx, open, suggestions) }] : [];
    });
    const gone = cells.filter((c) => columns.has(c.id));
    const lending = new Set(cells.flatMap((c) => c.lends.map((r) => r.suggestion)));
    for (const { id, open, lends } of gone.length === 0 ? [] : cells) {
      if (columns.has(id)) {
        continue;
      }

      const shown = shownRecord(ctx, open, 'tcPr');
      const holds = (cell: (typeof gone)[number]): boolean => {
        const column = columns.get(cell.id) ?? '';
        if (shown !== null) {
          return holdsCell(shown, open, cell.open);
        }

        const lender = lends.some((r) => r.suggestion === column);
        return !lender && (lending.has(column) || holdsCell(null, open, cell.open));
      };
      if (gone.every(holds)) {
        continue;
      }

      const last = lends.at(-1)?.element;
      const width = last ? childNamed(last, 'tcW') : undefined;
      if (width) {
        lent.set(id, width);
      } else {
        kept.add(id);
      }
    }
  }

  return { kept, lent };
}

/** The records of what removed columns (of `suggestions`) lent a cell, once they show no more. */
function lendingsOf(
  ctx: ComposeContext,
  open: OrderedItem,
  suggestions: ReadonlySet<string>
): OrderedElementRecord[] {
  const records = open.props.filter(
    (r): r is OrderedElementRecord =>
      r.kind === 'element' &&
      r.target === 'tcPr' &&
      r.untracked === true &&
      suggestions.has(r.suggestion) &&
      ctx.rejectedOrDrawnSo(r.suggestion)
  );
  return records.sort((a, b) => a.seq - b.seq);
}

/** A cell's span over grid columns (`w:gridSpan`), 1 when absent. */
function spanOf(cell: OoxmlElement): number {
  const tcPr = cell.children.find((c) => isWml(c, 'tcPr'));
  const span =
    tcPr && tcPr.kind !== 'textValue' ? tcPr.children.find((c) => isWml(c, 'gridSpan')) : undefined;
  const value =
    span && span.kind !== 'textValue' ? Number.parseInt(attributeOf(span, 'val') ?? '1', 10) : 1;
  return Number.isSafeInteger(value) && value > 0 ? value : 1;
}

/** The grid columns a row skips before or after its cells (`w:gridBefore`, `w:gridAfter`). */
function skippedOf(row: OoxmlElement): number {
  const trPr = row.children.find((c) => isWml(c, 'trPr'));
  if (!trPr || trPr.kind === 'textValue') return 0;
  let out = 0;
  for (const child of trPr.children) {
    if (child.kind === 'textValue' || !(isWml(child, 'gridBefore') || isWml(child, 'gridAfter')))
      continue;
    const value = Number.parseInt(attributeOf(child, 'val') ?? '0', 10);
    if (Number.isSafeInteger(value) && value > 0) out += value;
  }

  return out;
}

/**
 * The removed cells whose columns the drawn grid still holds: in each row, as many as fill the
 * grid past the cells that show, the ones the shown grid was taken with first, then in order.
 * The others' columns are gone from the grid already (a later record was taken after they
 * went, or before they came).
 */
function wantedCells(drawn: OoxmlElement, removed: readonly RemovedCell[]): RemovedCell[] {
  const grid = drawn.children.find((c) => isWml(c, 'tblGrid'));
  const columns =
    grid && grid.kind !== 'textValue' ? grid.children.filter((c) => isWml(c, 'gridCol')).length : 0;
  const byId = new Map(removed.map((r) => [r.id, r]));
  const out: RemovedCell[] = [];
  for (const row of drawn.children) {
    if (row.kind === 'textValue' || !isRow(row)) continue;
    let shown = skippedOf(row);
    const candidates: { cell: OoxmlElement; entry: RemovedCell }[] = [];
    for (const cell of row.children) {
      if (cell.kind === 'textValue' || !isCell(cell)) continue;
      const entry = byId.get(cell.id);
      if (entry) candidates.push({ cell, entry });
      else shown += spanOf(cell);
    }

    let need = columns - shown;
    const ordered = [
      ...candidates.filter((c) => c.entry.held),
      ...candidates.filter((c) => !c.entry.held),
    ];
    for (const { cell, entry } of ordered) {
      if (need <= 0) break;
      out.push(entry);
      need -= spanOf(cell);
    }
  }

  return out;
}

/** A node without its change records (`w:...Change`), deep; bounded by the XML depth. */
function withoutHistory(node: OoxmlElement, depth = 0): OoxmlElement {
  if (depth > 64) return node;
  const children: OoxmlNode[] = [];
  let changed = false;
  for (const child of node.children) {
    if (child.kind !== 'textValue' && /Change$/.test(child.localName)) {
      changed = true;
      continue;
    }

    const next = child.kind === 'textValue' ? child : withoutHistory(child, depth + 1);
    if (next !== child) changed = true;
    children.push(next);
  }

  return changed ? withChildren(node, children) : node;
}

/** A row's or a cell's properties with the grid placement children `names` from `from`. */
function withPlacement(
  owner: OoxmlElement,
  from: OoxmlElement | undefined,
  container: 'trPr' | 'tcPr',
  names: readonly string[],
  mint: () => string
): OoxmlElement {
  const own = owner.children.find((c) => isWml(c, container));
  const theirs = from?.children.find((c) => isWml(c, container));
  const placement = (props: OoxmlNode | undefined): OoxmlNode[] =>
    props && props.kind !== 'textValue'
      ? props.children.filter((c) => names.some((name) => isWml(c, name)))
      : [];
  const next = placement(theirs);
  const was = placement(own);
  if (JSON.stringify(next.map(plainOf)) === JSON.stringify(was.map(plainOf))) return owner;
  if (!own || own.kind === 'textValue') {
    if (next.length === 0) return owner;

    // First: the engine's `compactRevisionGrid` puts a new `w:trPr` first; `w:tcPr` leads a cell
    const made = build(mint(), 'generic', container, [], next);
    return withChildren(owner, [made, ...owner.children]);
  }

  const kept = own.children.filter((c) => !names.some((name) => isWml(c, name)));

  // The placement goes where the schema has it: `w:gridBefore`, `w:gridAfter` lead `w:trPr`
  // after `w:cnfStyle`; `w:gridSpan` follows `w:tcW`
  const at = kept.findIndex((c) =>
    container === 'trPr'
      ? !isWml(c, 'cnfStyle') && !isWml(c, 'divId')
      : !isWml(c, 'cnfStyle') && !isWml(c, 'tcW')
  );
  const children = [...kept];
  children.splice(at < 0 ? children.length : at, 0, ...next);
  return withChildren(
    owner,
    owner.children.map((c) => (c === own ? withChildren(own, children) : c))
  );
}

const plainOf = (node: OoxmlNode): unknown =>
  node.kind === 'textValue'
    ? node.value
    : [node.localName, node.attributes.map((a) => `${a.localName}=${a.value}`).sort()];

/** The grid's widths, in order. */
const gridWidths = (grid: OoxmlElement): string =>
  JSON.stringify(
    grid.children.flatMap((c) =>
      c.kind !== 'textValue' && isWml(c, 'gridCol') ? [attributeOf(c, 'w') ?? ''] : []
    )
  );

/**
 * The grid's record when it holds the grid as it is (a row height's anchor, Word's form): it
 * holds no old grid, so the join goes on and the record follows it.
 */
function unchangedGridRecord(grid: OoxmlElement): OoxmlElement | undefined {
  const record = grid.children.find((c) => isWml(c, 'tblGridChange'));
  const old =
    record && record.kind !== 'textValue' && record.children.find((c) => isWml(c, 'tblGrid'));
  if (!record || record.kind === 'textValue' || !old || old.kind === 'textValue') return undefined;
  return gridWidths(old) === gridWidths(grid) ? record : undefined;
}

/**
 * The resolved table with its grid joined as the engine joins it (`compactRevisionGrid`: a
 * grid boundary no remaining cell uses goes, with its width given to the column before it).
 * The engine keeps the grid as it was while the table holds any pending change record; the
 * list joins it unless a pending grid record (`w:tblGridChange`) holds the old grid, so a
 * decision on a column gives the same table live and after a reopen.
 */
function withJoinedGrid(ctx: ComposeContext, resolved: OoxmlElement): OoxmlElement {
  const grid = resolved.children.find((c) => isWml(c, 'tblGrid'));
  if (!grid || grid.kind === 'textValue') return resolved;
  const unchanged = unchangedGridRecord(grid);
  if (!unchanged && grid.children.some((c) => isWml(c, 'tblGridChange'))) return resolved;
  const bare = withoutHistory(resolved);
  const compacted = compactRevisionGrid(bare, ctx.mint);
  if (compacted === bare) return resolved;
  const newGrid = compacted.children.find((c) => isWml(c, 'tblGrid'));
  if (!newGrid || newGrid.kind === 'textValue') return resolved;
  const rows = new Map<string, OoxmlElement>();
  const cells = new Map<string, OoxmlElement>();
  for (const row of compacted.children) {
    if (row.kind === 'textValue' || !isRow(row)) continue;
    rows.set(row.id, row);
    for (const cell of row.children)
      if (cell.kind !== 'textValue' && isCell(cell)) cells.set(cell.id, cell);
  }

  const columns = newGrid.children.filter((c) => isWml(c, 'gridCol'));
  return withChildren(
    resolved,
    resolved.children.map((child) => {
      if (child === grid) {
        // An unchanged grid record holds the joined grid too
        const rest = grid.children.filter((c) => !isWml(c, 'gridCol'));
        return withChildren(grid, [
          ...columns,
          ...rest.map((c) =>
            c === unchanged ? withJoinedGridRecord(unchanged, columns, ctx.mint) : c
          ),
        ]);
      }

      if (child.kind === 'textValue' || !isRow(child)) return child;
      return withRowPlacements(ctx, child, rows, cells);
    })
  );
}

/** A grid record whose old grid takes the joined grid's columns, under new ids. */
function withJoinedGridRecord(
  record: OoxmlElement,
  columns: readonly OoxmlNode[],
  mint: () => string
): OoxmlElement {
  return withChildren(
    record,
    record.children.map((old) => {
      if (old.kind === 'textValue' || !isWml(old, 'tblGrid')) return old;
      const copies = columns.map((col) => (col.kind === 'textValue' ? col : withId(col, mint())));
      return withChildren(old, [...copies, ...old.children.filter((o) => !isWml(o, 'gridCol'))]);
    })
  );
}

/** A row with its grid placement, and its cells', as the joined table (`rows`, `cells`) has it. */
function withRowPlacements(
  ctx: ComposeContext,
  row: OoxmlElement,
  rows: ReadonlyMap<string, OoxmlElement>,
  cells: ReadonlyMap<string, OoxmlElement>
): OoxmlElement {
  const placed = withPlacement(
    row,
    rows.get(row.id),
    'trPr',
    ['gridBefore', 'gridAfter'],
    ctx.mint
  );
  return withChildren(
    placed,
    placed.children.map((cell) =>
      cell.kind === 'textValue' || !isCell(cell)
        ? cell
        : withPlacement(cell, cells.get(cell.id), 'tcPr', ['gridSpan'], ctx.mint)
    )
  );
}

/**
 * The drawn table with its removed cells gone, by the engine's rules: a direct column's grid
 * column goes with no width given to anyone; every other removed cell's width goes to its
 * neighbour (the engine accepts a cell deletion). The result is the same on every copy: the
 * engine is deterministic, and the ids it mints are replaced by compose's own.
 *
 * The engine's tracked column insertion lends half the neighbour's width to the new cells, with
 * no change record (Word's own form of `w:cellIns` cells). A decision gives the width back by
 * the engine's rule (`rebuildRevisionTable`: a removed cell's grid space goes to the cell before
 * it, or to the first one after it at the row's start), and an accepted `w:cellDel` does the
 * same. So while the widths still hold a removed cell's column, compose draws the cell as the
 * engine holds it before the decision and lets the engine remove it: live, after a save and a
 * reopen, and for several column suggestions in one table decided in any order.
 */
export function withoutRemovedCells(
  ctx: ComposeContext,
  table: OoxmlElement,
  candidates: readonly RemovedCell[]
): OoxmlElement {
  if (candidates.length === 0) return table;
  const after = withoutCandidates(ctx, table, candidates);
  ctx.removedColumns.push({ before: table, after, cells: candidates });
  return after;
}

function withoutCandidates(
  ctx: ComposeContext,
  table: OoxmlElement,
  candidates: readonly RemovedCell[]
): OoxmlElement {
  const known = new Set<string>();
  nodeIdsIn(table, known);

  // A candidate whose column the grid no longer holds just goes
  const removed = wantedCells(table, candidates);
  const wanted = new Set(removed.map((r) => r.id));
  const drawn =
    removed.length < candidates.length
      ? withoutCells(table, new Set(candidates.filter((r) => !wanted.has(r.id)).map((r) => r.id)))
      : table;

  // No width to give back, but the decision still joins the grid, as the engine does
  if (removed.length === 0) return withJoinedGrid(ctx, drawn);
  const { kept, lent } = survivorsOf(ctx, drawn, removed);
  const before = withLentWidths(drawn, lent);
  let part = withStoryBlocks(ctx.story.part, [before]);
  const direct = new Set(removed.filter((r) => r.how === 'direct').map((r) => r.id));
  for (const gridColumnId of wholeDirectColumns(before, part.root, direct)) {
    const result = applyTreeOp(part, { op: 'deleteTableColumn', tableId: drawn.id, gridColumnId });
    if (result.ok) part = result.part;
  }

  const current = findNode(part, drawn.id);
  if (!current || current.kind === 'textValue') return drawn;
  const joined = withJoinedGrid(ctx, withCellsAccepted(ctx, current, removed));
  const out = withComposeIds(ctx, withKeptCellProperties(joined, drawn, kept), known);
  return out.kind === 'textValue' ? drawn : out;
}

/** The table with every removed cell still in it gone, its width given to its neighbour. */
function withCellsAccepted(
  ctx: ComposeContext,
  current: OoxmlElement,
  removed: readonly RemovedCell[]
): OoxmlElement {
  const used = new Set<string>();
  revisionIdsIn(current, used);
  let next = 2_000_000_000;
  const freshId = (): string => {
    while (used.has(String(next))) next += 1;
    used.add(String(next));
    return String(next);
  };
  const byId = new Map(removed.map((r) => [r.id, r]));
  const sites: string[] = [];
  const marked = withTableCells(current, (cell) => {
    const entry = byId.get(cell.id);
    if (!entry) return cell;
    const attributes = revisionAttributes(freshId(), ctx.attribution(entry.suggestion));
    const marker = build(ctx.mint(), 'generic', 'cellDel', attributes, []);
    sites.push(marker.id);
    return withCellMarker(cell, marker, ctx.mint);
  });
  if (sites.length === 0) return marked;
  const scratch = withStoryBlocks(ctx.story.part, [marked]);
  const resolved = resolveRevisions(scratch, 'accept', undefined, { siteNodeIds: sites });
  const found = resolved.ok && resolved.part ? findNode(resolved.part, marked.id) : null;

  // A table the engine cannot resolve loses the removed cells with no width given: never a
  // cell drawn that a decision removed
  if (found && found.kind !== 'textValue') return found;
  return withoutCells(current, new Set(byId.keys()));
}

/** The table with each cell of `lent` at the width it lent, as the engine holds it. */
function withLentWidths(
  table: OoxmlElement,
  lent: ReadonlyMap<string, OoxmlElement>
): OoxmlElement {
  if (lent.size === 0) {
    return table;
  }

  return withTableCells(table, (cell) => {
    const width = lent.get(cell.id);
    const tcPr = childNamed(cell, 'tcPr');
    if (!width || !tcPr) {
      return cell;
    }

    const props = withChildren(
      tcPr,
      tcPr.children.map((c) => (isWml(c, 'tcW') ? width : c))
    );
    return withChildren(
      cell,
      cell.children.map((c) => (c === tcPr ? props : c))
    );
  });
}

/** The table with the `w:tcPr` that `drawn` gave each cell of `kept` put back. */
function withKeptCellProperties(
  table: OoxmlElement,
  drawn: OoxmlElement,
  kept: ReadonlySet<string>
): OoxmlElement {
  if (kept.size === 0) return table;
  const before = cellsById(drawn);
  return withTableCells(table, (cell) => {
    const was = kept.has(cell.id) ? before.get(cell.id) : undefined;
    const tcPr = was?.children.find((c) => isWml(c, 'tcPr'));
    if (!tcPr) return cell;
    return withChildren(cell, [tcPr, ...cell.children.filter((c) => !isWml(c, 'tcPr'))]);
  });
}
