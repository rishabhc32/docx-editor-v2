// The TABLE containers of a tracked format change, beside the run's and the paragraph's in
// `withPropertyChangeRecord`, which records all of them by the same two rules: the record is
// the original, and a write back at the original leaves none.
//
//   w:tcPr    → w:tcPrChange/w:tcPr       (CT_TcPrChange over CT_TcPrInner, §17.13.5.37)
//   w:trPr    → w:trPrChange/w:trPr       (CT_TrPrChange over CT_TrPrBase)
//   w:tblPr   → w:tblPrChange/w:tblPr     (CT_TblPrChange over CT_TblPrBase)
//   w:tblGrid → w:tblGridChange/w:tblGrid (CT_TblGridChange over CT_TblGridBase)
//
// Each change wrapper is the last member of its container's sequence, as `w:rPrChange` and
// `w:pPrChange` are, so the record engine's tail slot is the legal one here too.

import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { preservedRevisionProperty } from './revision-property-records.ts';

/** A table container that records its tracked format changes. */
export type TableContainer = 'cellProperties' | 'rowProperties' | 'tableProperties' | 'tableGrid';

/** How one table container records. */
interface TableRecord {
  /** The change wrapper. */
  readonly changeName: 'tcPrChange' | 'trPrChange' | 'tblPrChange' | 'tblGridChange';
  /** The element the wrapper holds its copy in: the container's own name. */
  readonly recordName: 'tcPr' | 'trPr' | 'tblPr' | 'tblGrid';
  /**
   * The kind re-reading the copy produces, which a save/reopen digest compares: `w:tcPr` and
   * `w:trPr` have no typed kind, so they re-read generic.
   */
  readonly recordKind: OoxmlElement['kind'];
  /**
   * Whether a child goes into the record. A cell's `w:cellIns`, `w:cellDel` and `w:cellMerge`
   * stay out of a `w:tcPrChange`, as the mark revisions stay out of a `w:rPrChange`: each is a
   * pending structural decision about the cell, not a property of it, and a reject restores the
   * container from the record. A row's `w:ins` and `w:del` stay out of a `w:trPrChange` the same
   * way.
   */
  readonly recordable: (child: OoxmlNode) => boolean;
}

function tableRecord(
  changeName: TableRecord['changeName'],
  recordName: TableRecord['recordName'],
  recordKind: TableRecord['recordKind']
): TableRecord {
  return {
    changeName,
    recordName,
    recordKind,
    recordable: (child) => !preservedRevisionProperty(changeName, child),
  };
}

/** How each table container records. */
export const TABLE_RECORDS: Readonly<Record<TableContainer, TableRecord>> = {
  cellProperties: tableRecord('tcPrChange', 'tcPr', 'generic'),
  rowProperties: tableRecord('trPrChange', 'trPr', 'generic'),
  tableProperties: tableRecord('tblPrChange', 'tblPr', 'tableProperties'),
  tableGrid: tableRecord('tblGridChange', 'tblGrid', 'tableGrid'),
};

/** Whether `container` is a table container, which `TABLE_RECORDS` describes. */
export function isTableContainer(container: string): container is TableContainer {
  return Object.hasOwn(TABLE_RECORDS, container);
}

/** The change wrapper each table property op writes (`PROPERTY_CHANGE_WRAPPER_OF_OP`). */
export const TABLE_CHANGE_WRAPPER_OF_OP: readonly (readonly [string, string])[] = [
  ['setTableCellFill', 'tcPrChange'],
  ['setTableCellVerticalAlignment', 'tcPrChange'],
  ['setTableCellBorders', 'tcPrChange'],
  ['setTableRowHeight', 'trPrChange'],
  ['setTableProperties', 'tblPrChange'],
  ['setTableColumnWidths', 'tblGridChange'],
  ['setTableRightEdgeWidth', 'tblGridChange'],
];
