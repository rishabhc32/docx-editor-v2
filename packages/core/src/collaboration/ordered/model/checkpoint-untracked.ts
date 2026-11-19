// The records a checkpoint's manifest keeps that Word writes no change record for (design D14).

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { normalizePartName } from '../../../store/package/opc-names.ts';
import { at, bool, fieldsOf, int, Invalid, oneOf, required } from '../change/read.ts';
import { manifestList, manifestString, readNamedOnce } from './checkpoint-suggestions.ts';
import {
  containerKey,
  isCheckpointContainer,
  propertiesXml,
  readPropertiesXml,
} from './checkpoint-takeovers.ts';
import type { OrderedPropertyTarget } from './types.ts';

/** The properties a container's records with no change record give. */
const TARGETS: readonly OrderedPropertyTarget[] = ['tblPr', 'tblGrid', 'trPr', 'tcPr', 'sdtPr'];

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
  /** The cells the table showed when it was taken, by their places among the part's cells. */
  readonly tableCells?: readonly number[];
}

/**
 * A container's properties that a pending record changes with no change record, by where the
 * file writes the container: its story part, its local name and its place among the part's
 * containers of that name. The file holds the properties the records give; the manifest keeps the
 * container's own and every record of them, decided or not.
 */
export interface CheckpointUntracked {
  readonly part: string;
  readonly container: string;
  readonly occurrence: number;
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
    ({ part, container, occurrence, target }) =>
      `${containerKey(part, container, occurrence)}\u0000${target}`
  );
}

function readEntry(value: unknown, path: string, suggestions: number): CheckpointUntracked {
  const f = fieldsOf(
    value,
    path,
    new Set(['part', 'container', 'occurrence', 'target', 'properties', 'records'])
  );
  const part = manifestString(required(f, 'part'), at(path, 'part'));
  const name = normalizePartName(part);
  if (!name.ok || name.partName !== part) {
    throw new Invalid('invalid-value', at(path, 'part'), 'not a part name');
  }

  const container = manifestString(required(f, 'container'), at(path, 'container'));
  if (!isCheckpointContainer(container)) {
    throw new Invalid('invalid-value', at(path, 'container'), 'not a container');
  }

  const target = oneOf(required(f, 'target'), at(path, 'target'), TARGETS);
  const records = manifestList(required(f, 'records'), at(path, 'records'), 1);
  return {
    part,
    container,
    occurrence: int(required(f, 'occurrence'), at(path, 'occurrence')),
    target,
    properties: f.has('properties')
      ? readPropertiesXml(f.get('properties'), at(path, 'properties'), target)
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
  const f = fieldsOf(
    value,
    path,
    new Set(['suggestion', 'element', 'untracked', 'merge', 'recordUnchanged', 'tableCells'])
  );
  const merge = f.has('merge') ? manifestList(f.get('merge'), at(path, 'merge')) : undefined;
  const cells = f.has('tableCells')
    ? manifestList(f.get('tableCells'), at(path, 'tableCells'))
    : undefined;
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    element: f.has('element')
      ? readPropertiesXml(f.get('element'), at(path, 'element'), target)
      : null,
    untracked: f.has('untracked') ? bool(f.get('untracked'), at(path, 'untracked')) : undefined,
    merge: merge?.map((key, k) => manifestString(key, at(at(path, 'merge'), k))),
    recordUnchanged: f.has('recordUnchanged')
      ? bool(f.get('recordUnchanged'), at(path, 'recordUnchanged'))
      : undefined,
    tableCells: cells?.map((cell, k) => int(cell, at(at(path, 'tableCells'), k))),
  };
}
