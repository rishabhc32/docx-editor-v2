// A rejected row formatting decision resets the row's cells that hold no record of their own to
// default formatting, as the engine's reject does (`restoreUnrecordedCells`, confirmed against
// Word; Word for Mac did the same on a row's `w:tblPrExChange`, 2026-10-29).

import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { withUnrecordedCellDefaults } from '../../../store/store/revision-table-rebuild.ts';
import type { OrderedItem, OrderedPropertyTarget } from '../model/types.ts';
import type { ComposeContext } from './context.ts';

const ROW_TARGETS: ReadonlySet<OrderedPropertyTarget> = new Set(['trPr', 'tblPrEx']);
const CELL_TARGETS: ReadonlySet<OrderedPropertyTarget> = new Set(['tcPr']);

/** The item's tracked records of these properties: a change record shows while one is pending. */
function trackedRecords(
  ctx: ComposeContext,
  item: OrderedItem,
  targets: ReadonlySet<OrderedPropertyTarget>
): string[] {
  return item.props.flatMap((p) =>
    p.kind === 'element' &&
    targets.has(p.target) &&
    p.untracked !== true &&
    p.acceptedOnly !== true &&
    ctx.doc.suggestions.get(p.suggestion)?.direct !== true
      ? [p.suggestion]
      : []
  );
}

/**
 * Reset, in place, the drawn cells of a row (`cells` gives each one's open item) that hold no
 * record, when a record of the row is rejected: its `w:trPr`, its `w:tblPrEx`, or one of its
 * cells' `w:tcPr`. Whether it changed a cell.
 *
 * A cell with a record of its own is that record's to decide and is never reset. The engine
 * decides in order: a cell's record decided before the row's leaves the cell with no record, so
 * the row's reject resets it. The room's decisions give one result in every order, the one where
 * the row's record goes first.
 */
export function resetUnrecordedCells(
  ctx: ComposeContext,
  row: OrderedItem,
  children: OoxmlNode[],
  cells: ReadonlyMap<OoxmlNode, OrderedItem> | undefined
): boolean {
  if (!cells || cells.size === 0) return false;
  const records = [
    ...trackedRecords(ctx, row, ROW_TARGETS),
    ...[...cells.values()].flatMap((cell) => trackedRecords(ctx, cell, CELL_TARGETS)),
  ];
  if (!records.some((suggestion) => ctx.rejectedOrDrawnSo(suggestion))) return false;
  let changed = false;
  children.forEach((child, index) => {
    const cell = cells.get(child);
    if (!cell || child.kind === 'textValue') return;
    if (trackedRecords(ctx, cell, CELL_TARGETS).length > 0) return;
    children[index] = withUnrecordedCellDefaults(child, ctx.mint);
    changed = true;
  });
  return changed;
}
