// The anchor a row height Word added with no row record leans on (`wordHeightRows`), and its
// move onto the row when a decision takes the table's record away.

import {
  createNodeIdAllocator,
  findNode,
  parentNodeOf,
  replaceNode,
  type EditOptions,
  type OoxmlEditResult,
} from '../package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import {
  tableHeightAnchor,
  wordHeightRows,
  type WordHeightRow,
} from './revision-table-implicit-height.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import type { RevisionSite } from './tree-op-revisions.ts';
import { cloneWithNewIds } from './tree-op-nodes.ts';
import { wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';
import { ROW_EXCEPTION_MEMBERS, TABLE_RECORDS } from './tree-op-tracked-table-records.ts';

/**
 * A row whose height must move onto its own anchor, and whose height it is; or (`drop`) a row
 * whose own anchor was decided, whose `w:tblPrEx` then sheds what the table's properties say:
 * that was there for the record alone.
 */
export type RowAnchorMove =
  | { readonly rowId: string; readonly revision: RevisionAttributionInput }
  | { readonly rowId: string; readonly drop: true };

/** The row and the grid that anchor a height on the row (`rowHeightAnchorEdits`). */
export interface RowAnchorEdits {
  readonly row: OoxmlElement;
  readonly grid?: OoxmlElement;
}

/** Whose height a row is: the author and date of its first cell's unchanged record. */
export function heightRevisionOf(height: WordHeightRow): RevisionAttributionInput | undefined {
  const author = wmlAttributeValue(height.cells[0]!, 'author');
  if (author === undefined) {
    return undefined;
  }

  const date = wmlAttributeValue(height.cells[0]!, 'date');
  return { author, ...(date ? { date } : {}) };
}

/**
 * The rows a decision of `matched` moves. A height leans on the table's one unchanged
 * `w:tblPrChange`, which a second author's height shares with the first, or on the row's own
 * unchanged `w:tblPrExChange`. A decided table record moves each height that leans on it onto
 * its row, unless one of the height's own cell records is decided with it. A decided row
 * record drops the row's copy of the table's properties.
 */
export function rowsLosingTableAnchor(
  part: OoxmlPart,
  matched: readonly RevisionSite[]
): RowAnchorMove[] {
  const selected = new Set(matched.map((site) => site.node.id));
  return matched.flatMap((site) => {
    if (!site.parent) {
      return [];
    }

    if (site.node.localName === 'tblPrExChange') {
      return droppedRowAnchor(part, site.node, site.parent.id);
    }

    if (site.node.localName === 'tblPrChange') {
      return heightsLeavingTableAnchor(part, site.node, site.parent.id, selected);
    }

    return [];
  });
}

/** A `drop` move for the row whose own height anchor `record` is, else none. */
function droppedRowAnchor(part: OoxmlPart, record: OoxmlElement, rowId: string): RowAnchorMove[] {
  const row = parentNodeOf(part, rowId);
  const table = row && parentNodeOf(part, row.id);
  if (!row || table?.kind !== 'table') {
    return [];
  }

  const anchored = wordHeightRows(table).some(
    (height) => height.rowAnchor && height.anchor.id === record.id
  );
  return anchored ? [{ rowId: row.id, drop: true }] : [];
}

/** The heights that lean on the table record `record` and move onto their rows. */
function heightsLeavingTableAnchor(
  part: OoxmlPart,
  record: OoxmlElement,
  tablePropertiesId: string,
  selected: ReadonlySet<string>
): RowAnchorMove[] {
  const table = parentNodeOf(part, tablePropertiesId);
  if (table?.kind !== 'table') {
    return [];
  }

  return wordHeightRows(table).flatMap((height): RowAnchorMove[] => {
    if (height.rowAnchor || height.anchor.id !== record.id) {
      return [];
    }

    if (height.cells.some((cell) => selected.has(cell.id))) {
      return [];
    }

    const revision = heightRevisionOf(height);
    return revision ? [{ rowId: height.row.id, revision }] : [];
  });
}

/**
 * The table properties the row states, in schema order: its own `w:tblPrEx` members, and the
 * table's `w:tblPr` members it does not override.
 */
function rowExceptions(table: OoxmlElement, row: OoxmlElement): OoxmlElement[] {
  const own = wmlChildNamed(row, 'tblPrEx');
  const tblPr = wmlChildNamed(table, 'tblPr');
  return [...ROW_EXCEPTION_MEMBERS].flatMap((name) => {
    const member = (own && wmlChildNamed(own, name)) ?? (tblPr && wmlChildNamed(tblPr, name));
    return member ? [member] : [];
  });
}

/**
 * The row anchored on its own unchanged `w:tblPrExChange` of `revision`, and the table's grid
 * with a `w:tblGridChange` when it has none: Word's form for a height whose table holds no
 * record for it. The row's `w:tblPrEx` holds the table's properties, and the record holds them
 * unchanged (Word for Mac, `store/__tests__/fixtures/word-row-height-authors.ts`). Without the
 * anchor the height would read as plain formatting, and its reject would keep it.
 */
export function rowHeightAnchorEdits(
  table: OoxmlElement,
  row: OoxmlElement,
  revision: RevisionAttributionInput,
  mint: () => string,
  revisionId: () => string
): RowAnchorEdits {
  const exceptions = rowExceptions(table, row);
  const copy = () => exceptions.map((child) => cloneWithNewIds(child, mint));
  const change = build(
    mint(),
    'generic',
    'tblPrExChange',
    revisionAttributes(revisionId(), revision),
    [build(mint(), 'generic', 'tblPrEx', [], copy())]
  );
  const tblPrEx = build(mint(), 'generic', 'tblPrEx', [], [...copy(), change]);
  const replaced = wmlChildNamed(row, 'tblPrEx');
  const anchored = {
    ...row,
    children: [tblPrEx, ...row.children.filter((child) => child.id !== replaced?.id)],
  } as OoxmlElement;
  const grid = wmlChildNamed(table, 'tblGrid');
  if (!grid || wmlChildNamed(grid, 'tblGridChange')) {
    return { row: anchored };
  }

  const columns = grid.children.filter((child) => child.kind !== 'textValue');
  const snapshot = build(
    mint(),
    TABLE_RECORDS.tableGrid.recordKind,
    'tblGrid',
    [],
    columns.map((child) => cloneWithNewIds(child, mint))
  );

  // `CT_Markup`: an `@w:id` and nothing else, as Word writes it
  const [id] = revisionAttributes(revisionId(), revision);
  const record = build(mint(), 'generic', 'tblGridChange', [id!], [snapshot]);
  return { row: anchored, grid: { ...grid, children: [...grid.children, record] } as OoxmlElement };
}

/** `part` with the row and the grid of `table` replaced by `edits`. */
export function applyRowAnchorEdits(
  part: OoxmlPart,
  table: OoxmlElement,
  edits: RowAnchorEdits,
  options?: EditOptions
): OoxmlEditResult {
  const replaced = replaceNode(part, edits.row.id, edits.row, options);
  const grid = wmlChildNamed(table, 'tblGrid');
  if (!replaced.ok || !edits.grid || !grid) {
    return replaced;
  }

  return replaceNode(replaced.part, grid.id, edits.grid, options);
}

/** The members of a row's `w:tblPrEx`, with no record, that the table's properties do not hold. */
function ownExceptions(table: OoxmlElement, exceptions: OoxmlElement): OoxmlNode[] {
  const tblPr = wmlChildNamed(table, 'tblPr');
  return exceptions.children.filter(
    (child) =>
      child.kind === 'textValue' ||
      tblPr === undefined ||
      !ROW_EXCEPTION_MEMBERS.has(child.localName) ||
      !tblPr.children.some(
        (own) =>
          own.kind !== 'textValue' &&
          own.localName === child.localName &&
          ooxmlTreesEqual(own, child)
      )
  );
}

/** `part` after a decision, each row of `moves` on its own anchor (`rowsLosingTableAnchor`). */
export function withRowHeightAnchors(
  part: OoxmlPart,
  moves: readonly RowAnchorMove[],
  options?: EditOptions
): OoxmlPart | null {
  if (moves.length === 0) return part;
  const mint = createNodeIdAllocator(part);
  const revisionId = nextRevisionId(part);
  let next = part;
  for (const move of moves) {
    const row = findNode(next, move.rowId);
    const table = row && parentNodeOf(next, row.id);
    if (row?.kind === 'textValue' || !row || table?.kind !== 'table') continue;
    if ('drop' in move) {
      const exceptions = wmlChildNamed(row, 'tblPrEx');
      const own = exceptions && ownExceptions(table, exceptions);
      if (!exceptions || !own || own.length === exceptions.children.length) continue;

      // The row keeps the exceptions it held before its anchor
      const kept = own.length === 0 ? [] : [{ ...exceptions, children: own } as OoxmlElement];
      const children = row.children.flatMap((child) =>
        child.id === exceptions.id ? kept : [child]
      );
      const dropped = replaceNode(next, row.id, { ...row, children } as OoxmlElement, options);
      if (!dropped.ok) return null;
      next = dropped.part;
      continue;
    }

    if (tableHeightAnchor(table)) continue;
    const edits = rowHeightAnchorEdits(table, row, move.revision, mint, revisionId);
    const applied = applyRowAnchorEdits(next, table, edits, options);
    if (!applied.ok) return null;
    next = applied.part;
  }

  return next;
}
