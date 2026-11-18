// A tracked row height in Word's form (Word for Mac, 2026-10-30;
// `store/__tests__/fixtures/word-created-row-height.ts` and `word-row-height-authors.ts`). The
// engine reads the form back as one change per author (`revision-table-implicit-height.ts`).

import {
  findNode,
  parentNodeOf,
  type EditOptions,
  type OoxmlEditResult,
} from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlPart } from '../package/ooxml-tree.ts';
import { tableHeightAnchor, wordHeightRows } from './revision-table-implicit-height.ts';
import { applyRowAnchorEdits, rowHeightAnchorEdits } from './revision-table-row-anchor.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { isWmlElement, wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';

const elements = (node: OoxmlElement | undefined): OoxmlElement[] =>
  (node?.children ?? []).filter((child): child is OoxmlElement => child.kind !== 'textValue');

/**
 * Whether a tracked write gives `row` a height in Word's form: the row had no `w:trPr`, or
 * only a height this author added in that form, and now holds a `w:trHeight` and nothing else.
 * Word writes such a height with no `w:trPrChange`, records each cell of the row unchanged
 * (`w:tcPrChange`), and anchors the change on the table (`heightAnchorOf`).
 */
export function addsPlainHeight(
  table: OoxmlElement,
  prior: OoxmlElement,
  row: OoxmlElement,
  author: string
): boolean {
  const now = elements(wmlChildNamed(row, 'trPr'));
  if (now.length !== 1 || !isWmlElement(now[0]!, 'trHeight')) {
    return false;
  }

  if (!wmlChildNamed(prior, 'trPr')) return true;
  return wordHeightRows(table).some(
    (height) =>
      height.row.id === prior.id && wmlAttributeValue(height.cells[0]!, 'author') === author
  );
}

/** How the write anchors its plain heights on the table `before` it. */
export type HeightAnchor = 'held' | 'table' | 'rows';

/**
 * `held`: the table already holds an anchor (anyone's), which the heights share: a table holds
 * one `w:tblPrChange`, so a later height of any author leans on it. `table`: the table holds no
 * `w:tblPrChange`, so the write adds the anchor, an unchanged `w:tblPrChange` and
 * `w:tblGridChange`. `rows`: the table holds a `w:tblPrChange` that records a change, or has no
 * `w:tblPr` or no `w:tblGrid`, so each row gets Word's row-level anchor, an unchanged
 * `w:tblPrExChange` (`rowHeightAnchorEdits`).
 */
export function heightAnchorOf(before: OoxmlElement): HeightAnchor {
  if (tableHeightAnchor(before)) return 'held';
  const tblPr = wmlChildNamed(before, 'tblPr');
  const grid = wmlChildNamed(before, 'tblGrid');
  const recorded = elements(tblPr).some((child) => child.localName === 'tblPrChange');
  return tblPr && grid && !recorded ? 'table' : 'rows';
}

/** The edit that anchors `rowId`'s plain height on its own row record, at apply time. */
export function rowAnchorEdit(
  rowId: string,
  revision: RevisionAttributionInput,
  mint: () => string,
  /** The op's own id; absent for another author's height, which takes a fresh one. */
  revisionId: (() => string) | undefined,
  options?: EditOptions
): (latest: OoxmlPart) => OoxmlEditResult {
  return (latest) => {
    const row = findNode(latest, rowId);
    const table = parentNodeOf(latest, rowId);
    if (!row || row.kind === 'textValue' || table?.kind !== 'table') {
      return { ok: true, part: latest };
    }

    const edits = rowHeightAnchorEdits(
      table,
      row,
      revision,
      mint,
      revisionId ?? nextRevisionId(latest)
    );
    return applyRowAnchorEdits(latest, table, edits, options);
  };
}
