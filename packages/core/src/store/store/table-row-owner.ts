// Which author owns a table row's pending formatting decision when two of them raced for it.
// Suggesting refuses a change to a row that holds another author's pending change
// (`holdsOtherAuthorsRowChange`), but in a shared room two people can still each change one
// row before either sees the other's record.

import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { WML_NAMESPACE_URI } from './tree-op-table-shared.ts';
import { REVISION_CAPABLE_TABLE_OPS } from './tree-op-tracked-table-ops.ts';
import type { TreeDocOp } from './tree-op-types.ts';

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
function order(left: number | string, right: number | string): number {
  if (left < right) {
    return -1;
  }

  if (left > right) {
    return 1;
  }

  return 0;
}

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
 * the earlier `w:date`, then the lesser author, then the lesser revision id. Every replica
 * settles the race the same way, from the records alone. The editor takes back a local change
 * that lost (`surface-table-row-race.ts`), and a room shows a row both changes reached as its
 * owner's one decision.
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
