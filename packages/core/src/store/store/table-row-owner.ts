// Which author owns a table row's pending formatting decision when two of them raced for it.
//
// Word reads every property record of a row — `w:trPrChange`, `w:tblPrExChange` and each
// cell's `w:tcPrChange` — as ONE formatting decision, and rejecting it resets each cell that has
// no record (`restoreUnrecordedCells`). So suggesting refuses a change to a row that holds
// another author's pending change (`holdsOtherAuthorsRowChange`). In a shared room two people
// can still each change one row before either sees the other's record. Every replica settles
// that race the same way, from the records alone: the record with the earlier `w:date` owns the
// row, then the lesser author, then the lesser revision id. The editor takes a local change
// that lost back (`surface-table-row-race.ts`), and a room shows a row both changes reached as
// its owner's one decision.

import type { OoxmlElement } from '../package/ooxml-tree.ts';
import { tableChildren } from './revision-table-children.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import {
  CT_TBLPR_SEQUENCE,
  CT_TCPR_SEQUENCE,
  CT_TRPR_SEQUENCE,
} from './tree-op-table-properties.ts';
import { WML_NAMESPACE_URI, wmlChildNamed } from './tree-op-table-shared.ts';
import { isWmlNamed } from './tree-op-tracked.ts';
import { REVISION_CAPABLE_TABLE_OPS } from './tree-op-tracked-table-ops.ts';
import type { TreeDocOp } from './tree-op-types.ts';

/** `CT_TblPrEx`: the `CT_TblPrBase` members a row's table property exception may hold. */
const TBLPREX_MEMBERS = new Set([
  'tblW',
  'jc',
  'tblCellSpacing',
  'tblInd',
  'tblBorders',
  'shd',
  'tblLayout',
  'tblCellMar',
  'tblLook',
]);

/**
 * The property containers of a table row and of its cells, each with its ECMA-376 child order
 * and the revision record that ends it.
 *
 * @internal
 */
export const ROW_PROPERTY_CONTAINERS: Readonly<
  Record<string, { readonly sequence: readonly string[]; readonly record: string }>
> = Object.freeze({
  trPr: { sequence: CT_TRPR_SEQUENCE, record: 'trPrChange' },
  tblPrEx: {
    sequence: [...CT_TBLPR_SEQUENCE.filter((name) => TBLPREX_MEMBERS.has(name)), 'tblPrExChange'],
    record: 'tblPrExChange',
  },
  tcPr: { sequence: CT_TCPR_SEQUENCE, record: 'tcPrChange' },
});

/**
 * Who made one pending row record, and when. A missing attribute reads as empty.
 *
 * @internal
 */
export interface RowRecordOwner {
  /** `w:author`. */
  readonly author: string;
  /** `w:date`, as written. */
  readonly date: string;
  /** `w:id`, the revision id. */
  readonly id: string;
}

/**
 * The attribution a pending row record carries in its WML attributes.
 *
 * @internal
 */
export function rowRecordOwnerOf(
  attributes: readonly {
    readonly namespaceUri: string;
    readonly localName: string;
    readonly value: string;
  }[]
): RowRecordOwner {
  const value = (name: string): string =>
    attributes.find(
      (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === name
    )?.value ?? '';
  return { author: value('author'), date: value('date'), id: value('id') };
}

/** Code-unit order, so every replica compares alike whatever its locale. */
const order = (left: number | string, right: number | string): number =>
  left < right ? -1 : left > right ? 1 : 0;

/** A record with no readable date is the latest. */
const timeOf = (date: string): number => {
  const time = Date.parse(date);
  return Number.isNaN(time) ? Number.POSITIVE_INFINITY : time;
};

/** Revision ids are decimal; any other id follows them. */
const numberOf = (id: string): number =>
  /^\d{1,15}$/.test(id) ? Number(id) : Number.POSITIVE_INFINITY;

/**
 * Negative when `left`'s record owns a row both records are in, positive when `right`'s does:
 * the earlier `w:date`, then the lesser author, then the lesser revision id.
 *
 * @internal
 */
export function compareRowRecordOwners(left: RowRecordOwner, right: RowRecordOwner): number {
  return (
    order(timeOf(left.date), timeOf(right.date)) ||
    order(left.author, right.author) ||
    order(numberOf(left.id), numberOf(right.id)) ||
    order(left.id, right.id)
  );
}

/** The pending property records of a row and its cells, in document order. */
export function rowPropertyRecords(row: OoxmlElement): OoxmlElement[] {
  const records: OoxmlElement[] = [];
  const collect = (owner: OoxmlElement, name: string): void => {
    const container = wmlChildNamed(owner, name);
    for (const child of container?.children ?? []) {
      if (child.kind !== 'textValue' && isWmlNamed(child, ROW_PROPERTY_CONTAINERS[name]!.record)) {
        records.push(child);
      }
    }
  };
  collect(row, 'tblPrEx');
  collect(row, 'trPr');
  for (const cell of tableChildren(row, 'tableCell')) collect(cell, 'tcPr');
  return records;
}

const TABLE_FORMATTING_OPS: ReadonlySet<string> = new Set(
  REVISION_CAPABLE_TABLE_OPS.filter(
    (op) => op !== 'insertTableColumn' && op !== 'deleteTableColumn'
  )
);

/**
 * Whether `op` proposes a table formatting change: a fill, border, alignment, size or table
 * property op with a revision, which the store records as row, cell and table records
 * (`withTablePropertyChanges`). A column insertion or deletion marks cells instead.
 *
 * @internal
 */
export function recordsTableFormatting(
  op: TreeDocOp
): op is TreeDocOp & { readonly tableId: string; readonly revision: RevisionAttributionInput } {
  return TABLE_FORMATTING_OPS.has(op.op) && 'revision' in op && op.revision !== undefined;
}
