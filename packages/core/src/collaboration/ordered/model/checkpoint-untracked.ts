// The records a checkpoint's manifest keeps that Word writes no change record for (design D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { at, bool, int, oneOf, plainFields, required } from '../change/read.ts';
import { manifestList, manifestString, readNamedOnce } from './checkpoint-suggestions.ts';
import { propertiesXml, readPropertiesXml } from './checkpoint-takeovers.ts';
import type { OrderedPropertyTarget } from './types.ts';

/**
 * The properties a container's records with no change record give, and the container whose
 * properties they are.
 */
export const UNTRACKED_CONTAINERS: ReadonlyMap<OrderedPropertyTarget, string> = new Map([
  ['tblPr', 'tbl'],
  ['tblGrid', 'tbl'],
  ['trPr', 'tr'],
  ['tcPr', 'tc'],
  ['sdtPr', 'sdt'],
]);

/**
 * A record, decided or not, of properties some record changes with no change record
 * (`untracked`: a control's value, a column's lent width). The file holds only what they give, so
 * the manifest keeps each record of those properties as the room has it.
 */
export interface CheckpointUntrackedRecord {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly element: OoxmlElement | null;
  /** Shown with no change record. */
  readonly untracked?: boolean;
  readonly merge?: readonly string[];
  readonly recordUnchanged?: boolean;
  /** The cells the table showed when it was taken, by their labels. */
  readonly tableCells?: readonly number[];
  /** The suggestion the record was taken from, by its place in the manifest's list. */
  readonly replaced?: number;
}

/**
 * A container's properties that a pending record changes with no change record, by the label on
 * the container (a table, a row, a cell or a control). The file holds the properties the records
 * give; the manifest keeps the container's own and every record of them, decided or not.
 */
export interface CheckpointUntracked {
  readonly label: number;
  readonly target: OrderedPropertyTarget;
  /** The container's own properties of the target, under the records; absent for none. */
  readonly properties?: OoxmlElement;
  /** The records, in the order they apply. */
  readonly records: readonly CheckpointUntrackedRecord[];
}

/** The untracked records as the manifest lists them: JSON leaves out a field with no value. */
export function untrackedJson(untracked: readonly CheckpointUntracked[]): unknown[] {
  return untracked.map(({ properties, records, ...container }) => ({
    ...container,
    properties: properties === undefined ? undefined : propertiesXml(properties),
    records: records.map(({ element, ...record }) => ({
      ...record,
      element: element === null ? undefined : propertiesXml(element),
    })),
  }));
}

/**
 * The listed untracked records, each checked: every value comes from storage. Each names a
 * suggestion of the manifest's list of `suggestions`, and a container once for each target.
 *
 * @throws Invalid when one does not read.
 */
export function readUntracked(value: unknown, suggestions: number): CheckpointUntracked[] {
  return readNamedOnce(
    value,
    'untracked',
    (entry, path) => readEntry(entry, path, suggestions),
    ({ label, target }) => `${label} ${target}`
  );
}

function readEntry(value: unknown, path: string, suggestions: number): CheckpointUntracked {
  const f = plainFields(value, path);
  const target = oneOf(required(f, 'target'), at(path, 'target'), [...UNTRACKED_CONTAINERS.keys()]);
  const records = manifestList(required(f, 'records'), at(path, 'records'), 1);
  return {
    label: int(required(f, 'label'), at(path, 'label')),
    target,
    properties: f.has('properties')
      ? readPropertiesXml(f.get('properties'), at(path, 'properties'), new Set([target]))
      : undefined,
    records: records.map((record, j) =>
      readRecord(record, at(at(path, 'records'), j), target, suggestions)
    ),
  };
}

function readRecord(
  value: unknown,
  path: string,
  target: string,
  suggestions: number
): CheckpointUntrackedRecord {
  const f = plainFields(value, path);
  const merge = f.has('merge') ? manifestList(f.get('merge'), at(path, 'merge')) : undefined;
  const cells = f.has('tableCells')
    ? manifestList(f.get('tableCells'), at(path, 'tableCells'))
    : undefined;
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    element: f.has('element')
      ? readPropertiesXml(f.get('element'), at(path, 'element'), new Set([target]))
      : null,
    untracked: f.has('untracked') ? bool(f.get('untracked'), at(path, 'untracked')) : undefined,
    merge: merge?.map((key, k) => manifestString(key, at(at(path, 'merge'), k))),
    recordUnchanged: f.has('recordUnchanged')
      ? bool(f.get('recordUnchanged'), at(path, 'recordUnchanged'))
      : undefined,
    tableCells: cells?.map((cell, k) => int(cell, at(at(path, 'tableCells'), k))),
    replaced: f.has('replaced')
      ? int(f.get('replaced'), at(path, 'replaced'), 0, suggestions - 1)
      : undefined,
  };
}
