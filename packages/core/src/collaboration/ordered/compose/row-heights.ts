// Word's anchor for a pending row height with no row record (D17).

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  tableHeightAnchor,
  wordHeightRows,
} from '../../../store/store/revision-table-implicit-height.ts';
import { rowHeightAnchorEdits } from '../../../store/store/revision-table-row-anchor.ts';
import { inserter, statusOf } from '../model/document.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import type { OrderedItem } from '../model/types.ts';
import { childNamed, isWml, withChildren } from '../ooxml-nodes.ts';
import type { ComposeContext } from './context.ts';

/**
 * The pending suggestion whose height the row shows with no row record, if any. A height of the
 * row's own insertion is part of it: Word writes it with no anchor.
 */
export function pendingHeightOf(ctx: ComposeContext, open: OrderedItem): string | undefined {
  if (ctx.options.pending === 'reject') return undefined;
  for (const p of open.props)
    if (
      p.kind === 'element' &&
      p.target === 'trPr' &&
      p.untracked === true &&
      p.element !== null &&
      childNamed(p.element, 'trHeight') !== undefined &&
      statusOf(ctx.doc, p.suggestion) === 'pending' &&
      p.suggestion !== inserter(ctx.doc, open)
    )
      return p.suggestion;
  return undefined;
}

/**
 * The table with each row of `heights` (row id to suggestion) on its own anchor, when needed. On
 * a row with no `w:trPr`, Word writes the height with no `w:trPrChange` and anchors it on the
 * table's one unchanged `w:tblPrChange` (the first author's, which a later author's height
 * shares); that record comes from the suggestion's own `w:tblPr` record. When the table holds no
 * such record (its author's change was decided), the height leans on its row's own: a
 * `w:tblPrEx` that holds the table's properties, with an unchanged `w:tblPrExChange` of the
 * height's author, and a `w:tblGridChange` on the grid (Word for Mac,
 * `store/__tests__/fixtures/word-row-height-authors.ts`). This writes the row's, as the engine's
 * decision does (`store/store/revision-table-row-anchor.ts`), on a row Word's form reads as the
 * height (`wordHeightRows`): on another (a row with a record of its own, a cell with none), a
 * load would read the anchor as a record of the row.
 */
export function withRowHeightAnchors(
  ctx: ComposeContext,
  table: OoxmlNode,
  heights: ReadonlyMap<string, string> | undefined
): OoxmlNode {
  if (!heights || heights.size === 0 || table.kind === 'textValue' || tableHeightAnchor(table))
    return table;
  let current: OoxmlElement = table;
  for (const [rowId, suggestion] of heights) {
    const row = current.children.find((c) => c.id === rowId);
    if (!row || row.kind === 'textValue') continue;

    // A row that holds its own anchor already (a file's, or one this suggestion wrote)
    const exceptions = childNamed(row, 'tblPrEx');
    if (exceptions && childNamed(exceptions, 'tblPrExChange')) continue;
    const role = REVISION_ROLE.tableProperties;
    const edits = rowHeightAnchorEdits(current, row, ctx.attribution(suggestion), ctx.mint, () =>
      ctx.revisionId(suggestion, role)
    );
    const marked = edits.row.children.map((child) =>
      child.kind !== 'textValue' && isWml(child, 'tblPrEx')
        ? withChildren(child, ctx.withRecordAttributes(child.children, suggestion, role))
        : child
    );
    const grid = edits.grid;
    const anchored = withChildren(
      current,
      current.children.map((child) => {
        if (child.id === rowId) return withChildren(edits.row, marked);
        if (grid && child.kind !== 'textValue' && isWml(child, 'tblGrid')) return grid;
        return child;
      })
    );
    if (wordHeightRows(anchored).some((height) => height.row.id === rowId)) {
      current = anchored;
    }
  }

  return current;
}
