// Repeat Header Rows and table placement — the table PROPERTY commands of the planner.
//
// `toggleHeaderRow` follows Word's rule that header rows are the table's LEADING rows: a row
// repeats atop each page the table continues onto only while every row above it does, which is
// how layout collects them (`table-flow-pagination.ts`). So turning it on marks every row from
// the first through the last selected one, and turning it off clears it from the first selected
// row down — a flag left on a row below a plain row would repeat nothing. Both write the store's
// `headerRowCount`, which keeps the leading block exactly that many rows long.
//
// `setTableProperties` takes left and right as the reader sees them, as the planner's border
// and column commands do, and writes them in document order: the leading edge of a
// right-to-left table (`w:bidiVisual`) is on the right. A table that text wraps around is
// placed by its position (`w:tblpPr`) and ignores `w:jc`, so its alignment is refused rather
// than written where it would not move the table.

import type { EditorCommand } from '../contracts/editor.ts';
import type { SemanticLayout } from '../layout/semantic-records.ts';
import { tableInteractionIndex } from '../layout/semantic-table-interaction.ts';
import { readFlag } from '../layout/table-row-properties.ts';
import type { OoxmlElement, OoxmlPart } from '../store/package/ooxml-tree.ts';
import { wmlChildNamed } from '../store/store/tree-op-table-shared.ts';
import { readEditableTableTopology } from '../store/store/tree-op-table-topology.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { TableCommandPlan } from './table-command-plan.ts';

type SetTableProperties = Extract<EditorCommand, { type: 'setTableProperties' }>;
type PreferredWidth = Extract<TreeDocOp, { op: 'setTableProperties' }>['preferredWidth'];
type Refusal = Extract<TableCommandPlan, { ok: false }>;

/** The refusal for turning off a header row that the table style, not the row, makes one. */
export const HEADER_ROW_FROM_STYLE_REASON = 'the table style makes this row a header row';
/** The refusal for aligning a table that text wraps around. */
export const FLOATING_TABLE_ALIGNMENT_REASON =
  'text wraps around this table, so alignment does not move it';

const SET_TABLE_PROPERTIES_KEYS: readonly string[] = [
  'type',
  'width',
  'widthType',
  'justification',
];

const invalid = (reason: string): Refusal => ({ ok: false, code: 'invalidArgs', reason });

/** The row's own `w:tblHeader`, read as layout reads it. */
const ownHeaderFlag = (row: OoxmlElement): boolean =>
  readFlag(wmlChildNamed(row, 'trPr'), 'tblHeader');

interface HeaderRowState {
  readonly rows: readonly OoxmlElement[];
  /** The first and last selected row, as indexes in table order. */
  readonly first: number;
  readonly last: number;
  /** How many leading rows repeat as header rows. */
  readonly leading: number;
}

function headerRowState(
  part: OoxmlPart,
  layout: SemanticLayout,
  tableId: string,
  rowIds: readonly string[]
): HeaderRowState | null {
  const topology = readEditableTableTopology(part.root, tableId);
  if (!topology.ok) return null;
  const rows = topology.topology.rows.map((entry) => entry.row);
  const selected = new Set(rowIds);
  const indexes = rows.flatMap((row, index) => (selected.has(row.id) ? [index] : []));
  if (indexes.length === 0) return null;
  // Layout resolves the table style's `w:tblHeader` as well. A row it placed nowhere — one the
  // review view hides — answers with its own flag.
  const resolved = new Map<string, boolean>();
  for (const { table, row } of tableInteractionIndex(layout).occurrences) {
    if (table.tableId === tableId && !row.isHeaderRepeat) resolved.set(row.id, row.isHeaderRow);
  }
  const plain = rows.findIndex((row) => !(resolved.get(row.id) ?? ownHeaderFlag(row)));
  return {
    rows,
    first: indexes[0]!,
    last: indexes[indexes.length - 1]!,
    leading: plain === -1 ? rows.length : plain,
  };
}

/** Whether every selected row repeats as a header row, which is `toggleHeaderRow` pressed. */
export function headerRowsActive(
  part: OoxmlPart,
  layout: SemanticLayout,
  tableId: string,
  rowIds: readonly string[]
): boolean {
  const state = headerRowState(part, layout, tableId, rowIds);
  return state !== null && state.last < state.leading;
}

/** `toggleHeaderRow` over the selected rows, not yet validated against the store. */
export function planHeaderRowToggle(
  part: OoxmlPart,
  layout: SemanticLayout,
  tableId: string,
  rowIds: readonly string[]
): TableCommandPlan {
  const state = headerRowState(part, layout, tableId, rowIds);
  if (!state) return invalid('the table target is no longer valid');
  const on = state.last >= state.leading;
  // A row the table style makes a header has no flag of its own to clear, so turning it off
  // here would change nothing. Say so rather than report a change that did not happen.
  if (!on && state.rows.slice(state.first, state.leading).some((row) => !ownHeaderFlag(row))) {
    return { ok: false, code: 'unsupported', reason: HEADER_ROW_FROM_STYLE_REASON };
  }
  const headerRowCount = on ? state.last + 1 : state.first;
  return {
    ok: true,
    ops: [{ op: 'setTableProperties', tableId, headerRowCount }],
    selection: { kind: 'preserveSelection' },
  };
}

/** Why a `setTableProperties` command is malformed, or null; the store checks its values. */
export function tablePropertiesShapeError(command: EditorCommand): Refusal | null {
  const keys = Object.keys(command);
  return keys.length < 2 || !keys.every((key) => SET_TABLE_PROPERTIES_KEYS.includes(key))
    ? invalid('setTableProperties takes width, widthType and justification only')
    : null;
}

/**
 * The command's width as the store's `preferredWidth`: undefined leaves `w:tblW` alone, and
 * null removes it. A width travels with its unit, because the same number is twips under `dxa`
 * and fiftieths of a percent under `pct`; the store refuses one without the other.
 */
function preferredWidthOf({ width, widthType }: SetTableProperties): PreferredWidth {
  if (width === undefined && widthType === undefined) return undefined;
  const value = widthType === null || widthType === 'auto' ? (width ?? 0) : width;
  return widthType === null && value === 0 ? null : ({ type: widthType, value } as PreferredWidth);
}

/** `setTableProperties` on the selected table, not yet validated against the store. */
export function planTablePlacement(
  command: SetTableProperties,
  layout: SemanticLayout,
  tableId: string,
  visualRtl: boolean
): TableCommandPlan {
  const preferredWidth = preferredWidthOf(command);
  const { justification } = command;
  const floating = tableInteractionIndex(layout).occurrences.some(
    ({ table }) => table.tableId === tableId && table.floatingWrap !== undefined
  );
  if (justification !== undefined && floating) {
    return { ok: false, code: 'unsupported', reason: FLOATING_TABLE_ALIGNMENT_REASON };
  }
  const mirror = visualRtl && (justification === 'left' || justification === 'right');
  const jc = mirror ? (justification === 'left' ? 'right' : 'left') : justification;
  const op: TreeDocOp = {
    op: 'setTableProperties',
    tableId,
    ...(jc === undefined ? {} : { justification: jc }),
    ...(preferredWidth === undefined ? {} : { preferredWidth }),
  };
  return { ok: true, ops: [op], selection: { kind: 'preserveSelection' } };
}
