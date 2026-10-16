/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * One owner for a table row that two authors' pending formatting changes both reached.
 *
 * Word reads every property record of a row as one formatting decision, and rejecting it resets
 * each cell of the row that has no record. Suggesting refuses a change to a row that holds
 * another author's pending change, but two people can each change one row before either sees
 * the other's record, and each change records every cell of its row. Merged, each container
 * showed the first record in shared order, so one row mixed both authors: one card named one of
 * them, and rejecting the two changes in either order lost the widths of the cells the second
 * rejection found without a record.
 *
 * Every replica now gives such a row one owner, from shared state alone: the author of its
 * record with the earliest `w:date`, then the lesser author, then the lesser revision id
 * (`compareRowRecordOwners`, the rule by which the editor takes a losing change back). Each
 * record of the row shows its owner's attribution, so the row is one decision with one card:
 * accepting it keeps every merged property, and rejecting it restores every container from
 * the record it holds. The records stay as written in shared state, so undo still finds them,
 * and a row whose records only the source file wrote stays as loaded.
 */

import {
  compareRowRecordOwners,
  ROW_PROPERTY_CONTAINERS,
  rowRecordOwnerOf,
  type RowRecordOwner,
} from '@docx-editor.dev/core/collaboration/replication';
import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import { replicaOfLogicalId, type LogicalId } from './identity.ts';
import type { DocumentRegistry } from './registry.ts';
import { isElementRecord, type ElementRecord, type SharedRecord } from './schema.ts';
import { sharedChildShell, type ChildShell } from './singleton-children.ts';

const CONTAINERS: ReadonlySet<string> = new Set(Object.keys(ROW_PROPERTY_CONTAINERS));
const RECORDS: ReadonlySet<string> = new Set(
  Object.values(ROW_PROPERTY_CONTAINERS).map(({ record }) => record)
);

const isWml = (shell: ChildShell, names: ReadonlySet<string>): boolean =>
  shell.kind === 'generic' &&
  shell.namespaceUri === WML_NAMESPACE_URI &&
  names.has(shell.localName);

/** A node whose change can move a row's owner: the row, a cell, a container, a record. */
function touchesOwner(shell: ChildShell): boolean {
  return (
    shell.kind === 'tableRow' ||
    shell.kind === 'tableCell' ||
    isWml(shell, CONTAINERS) ||
    isWml(shell, RECORDS)
  );
}

/** The row a record, container or cell belongs to, or the row itself. Null outside a row. */
function rowOf(registry: DocumentRegistry, id: LogicalId): LogicalId | null {
  let at: LogicalId | null = id;
  for (let depth = 0; at !== null && depth < 4; depth += 1) {
    const kind = registry.kindOf(at);
    if (kind === 'tableRow') return at;
    if (kind === 'table' || kind === 'paragraph') return null;
    at = registry.parentOf(at);
  }
  return null;
}

/** The live children of `parentId` whose shells `accept` takes. */
function childrenWhere(
  registry: DocumentRegistry,
  parentId: LogicalId,
  accept: (shell: ChildShell) => boolean
): LogicalId[] {
  const parent = registry.record(parentId);
  if (!parent || !isElementRecord(parent)) return [];
  return parent.childIds.filter((id) => {
    const shell = sharedChildShell(registry, id);
    return shell !== null && accept(shell);
  });
}

/** Every live pending property record of a row and its cells, hidden copies included. */
function rowRecords(registry: DocumentRegistry, rowId: LogicalId): ElementRecord[] {
  const containers = (ownerId: LogicalId): LogicalId[] =>
    childrenWhere(registry, ownerId, (shell) => isWml(shell, CONTAINERS));
  const cells = childrenWhere(registry, rowId, (shell) => shell.kind === 'tableCell');
  const records: ElementRecord[] = [];
  for (const container of [...containers(rowId), ...cells.flatMap(containers)]) {
    for (const id of childrenWhere(registry, container, (shell) => isWml(shell, RECORDS))) {
      const record = registry.record(id);
      if (record && isElementRecord(record)) records.push(record);
    }
  }
  return records;
}

/**
 * The attribution every record of a row shows, or null when the row needs none: its records
 * have one author, or no peer wrote any of them.
 */
function rowOwner(records: readonly ElementRecord[]): RowRecordOwner | null {
  if (records.every((record) => replicaOfLogicalId(record.logicalId) === null)) return null;
  const owners = records.map((record) => rowRecordOwnerOf(record.attributes));
  if (owners.every((owner) => owner.author === owners[0]!.author)) return null;
  return owners.reduce((best, owner) => (compareRowRecordOwners(owner, best) < 0 ? owner : best));
}

/** The attributes a row owner's attribution takes the place of. */
const ATTRIBUTION: ReadonlySet<string> = new Set(['author', 'date', 'id']);

/** `attributes` with the row owner's `w:author`, `w:date` and `w:id` in place of their own. */
function attributedTo(
  attributes: ElementRecord['attributes'],
  owner: RowRecordOwner
): ElementRecord['attributes'] {
  const wml = (attribute: ElementRecord['attributes'][number]): boolean =>
    attribute.namespaceUri === WML_NAMESPACE_URI;
  const prefix = attributes.find(wml)?.prefix;
  const named = (['author', 'date', 'id'] as const)
    .filter((localName) => owner[localName] !== '')
    .map((localName) => ({
      namespaceUri: WML_NAMESPACE_URI,
      localName,
      ...(prefix === undefined ? {} : { prefix }),
      value: owner[localName],
    }));
  return [
    ...attributes.filter((attribute) => !wml(attribute) || !ATTRIBUTION.has(attribute.localName)),
    ...named,
  ];
}

/**
 * `record` as a replica shows it. A pending property record of a row that another author's
 * record also reached carries the attribution of the row's owner; the owner's own records and
 * every other node show as written.
 */
export function rowOwnedRecord<T extends SharedRecord | null>(
  registry: DocumentRegistry,
  record: T
): T {
  if (!record || !isElementRecord(record) || !isWml(record, RECORDS)) return record;
  const rowId = rowOf(registry, record.logicalId);
  const owner = rowId === null ? null : rowOwner(rowRecords(registry, rowId));
  if (!owner || rowRecordOwnerOf(record.attributes).author === owner.author) return record;
  return { ...record, attributes: attributedTo(record.attributes, owner) };
}

/**
 * The records to build again because `dirty` can move their row's owner: every record of each
 * row whose records, property containers or cells changed. A record's attribution depends on
 * the whole row, so a cached one is stale as soon as any of them changes.
 */
export function rowRecordsToRebuild(
  registry: DocumentRegistry,
  dirty: Iterable<LogicalId>
): LogicalId[] {
  const rows = new Set<LogicalId>();
  for (const id of dirty) {
    const shell = sharedChildShell(registry, id);
    if (!shell || !touchesOwner(shell)) continue;
    const rowId = rowOf(registry, id);
    if (rowId !== null) rows.add(rowId);
  }
  return [...rows].flatMap((rowId) =>
    rowRecords(registry, rowId).map((record) => record.logicalId)
  );
}
