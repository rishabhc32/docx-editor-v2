// Word's row heights with no record of the row's own, as the room's suggestions.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import {
  anchorsOwnHeight,
  tableHeightAnchor,
  unchangedRecord,
  wordHeightRows,
  type WordHeightRow,
} from '../../../store/store/revision-table-implicit-height.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';

const wmlChildren = (node: OoxmlElement, name: string): OoxmlElement[] =>
  node.children.filter((c) => c.kind !== 'textValue' && isWml(c, name)) as OoxmlElement[];

/** One change of Word's row heights: its records, and the `w:trPr` of each row it added. */
export interface WordHeightChange {
  readonly records: readonly OoxmlElement[];
  readonly added: ReadonlyMap<OoxmlElement, OoxmlElement>;
}

const actorOf = (record: OoxmlElement): string =>
  `${attributeOf(record, 'author') ?? ''}\u0000${attributeOf(record, 'date') ?? ''}`;

/**
 * Word's row heights in a table, one change for each author (`wordHeightRows`, the engine's
 * reading). On a row with no `w:trPr`, Word writes the new `w:trHeight` with no `w:trPrChange`
 * and records the change on the table instead (a file Word saved:
 * `store/__tests__/fixtures/word-created-row-height.ts`); the room reads those records as one
 * suggestion, as Word and the engine reject them as one change. The table's anchor (its
 * `w:tblPrChange` and grid record) is the change of its own author, with that author's heights,
 * and with each row that author recorded with a `w:trPrChange` and unchanged cells. Another
 * author's heights that lean on it are a change of their own: their cells' records
 * (`store/__tests__/fixtures/word-row-height-authors.ts`). A height on its row's own anchor is a
 * change of its own: the row's `w:tblPrExChange` and its cells' records, and the grid record
 * when the table has no anchor (Word's form once the anchor's change was decided).
 */
export function wordRowHeights(table: OoxmlElement): WordHeightChange[] {
  const anchor = tableHeightAnchor(table);
  const tblGrid = wmlChildren(table, 'tblGrid')[0];
  const grid = tblGrid && wmlChildren(tblGrid, 'tblGridChange')[0];
  const plain = wordHeightRows(table);
  const changes = new Map<
    string,
    { records: OoxmlElement[]; added: Map<OoxmlElement, OoxmlElement> }
  >();
  const changeOf = (key: string, records: readonly OoxmlElement[] = []) => {
    let change = changes.get(key);
    if (!change) changes.set(key, (change = { records: [...records], added: new Map() }));
    return change;
  };
  const trPrOf = (row: OoxmlElement) => wmlChildren(row, 'trPr')[0]!;
  if (anchor && grid) {
    const own = changeOf('table', [anchor, grid]);
    for (const row of wmlChildren(table, 'tr')) {
      const records = anchorAuthorRecords(row, anchor, plain);
      if (records) {
        own.records.push(...records);
      }
    }
  }

  for (const height of plain) {
    const change = changeOf(heightKey(height), height.rowAnchor ? [height.anchor] : []);

    // With no anchor on the table, the grid record goes with the first height on its row's own
    if (
      height.rowAnchor &&
      !anchor &&
      grid &&
      ![...changes.values()].some((c) => c.records.includes(grid))
    )
      change.records.push(grid);
    change.records.push(...height.cells);
    change.added.set(height.row, trPrOf(height.row));
  }

  const own = changes.get('table');
  if (own && own.records.length <= 2) changes.delete('table');
  return [...changes.values()];
}

/**
 * The records of a row the anchor's author recorded with a `w:trPrChange` and its cells
 * unchanged: the row's record and its cells', or null for any other row.
 */
function anchorAuthorRecords(
  row: OoxmlElement,
  anchor: OoxmlElement,
  plain: readonly WordHeightRow[]
): OoxmlElement[] | null {
  if (plain.some((height) => height.row === row)) {
    return null;
  }

  const cells = wmlChildren(row, 'tc');
  if (cells.length === 0) {
    return null;
  }

  if (row.children.some((c) => c.kind !== 'textValue' && !isWml(c, 'trPr') && !isWml(c, 'tc'))) {
    return null;
  }

  const props = wmlChildren(row, 'trPr');
  const rowRecords = props.length === 1 ? wmlChildren(props[0]!, 'trPrChange') : [];
  const rowRecord = rowRecords.length === 1 ? rowRecords[0] : undefined;
  const cellRecords = cells.map((cell) => {
    const tcPr = wmlChildren(cell, 'tcPr');
    return tcPr.length === 1 ? unchangedRecord(tcPr[0]!, 'tcPrChange') : undefined;
  });
  const same = (record: OoxmlElement | undefined): record is OoxmlElement =>
    record !== undefined && actorOf(record) === actorOf(anchor);
  if (!same(rowRecord) || !cellRecords.every(same)) {
    return null;
  }

  return [rowRecord, ...cellRecords];
}

/** The change of a height with no row record: its row's own, the table's or its author's. */
function heightKey(height: WordHeightRow): string {
  if (height.rowAnchor) {
    return `row\u0000${height.row.id}`;
  }

  if (anchorsOwnHeight(height)) {
    return 'table';
  }

  return `actor\u0000${actorOf(height.cells[0]!)}`;
}
