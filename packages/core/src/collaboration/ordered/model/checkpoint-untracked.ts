// The records a checkpoint's manifest keeps that Word writes no change record for (design D14).

import { serializeOoxmlPart } from '../../../store/package/ooxml-serialize.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
} from '../../../store/package/ooxml-tree.ts';
import { normalizePartName } from '../../../store/package/opc-names.ts';
import { at, bool, fieldsOf, int, Invalid, list, required, str } from '../change/read.ts';
import { MAX_FIELD_UNITS, MAX_LISTED } from './checkpoint-suggestions.ts';
import { containerKey, isCheckpointContainer, propertiesPart } from './checkpoint-takeovers.ts';
import type { OrderedPropertyTarget } from './types.ts';

/** The properties a container's records with no change record give. */
const TARGETS: ReadonlySet<string> = new Set<OrderedPropertyTarget>([
  'tblPr',
  'tblGrid',
  'trPr',
  'tcPr',
  'sdtPr',
]);

/**
 * A pending record of properties some record changes with no change record (`untracked`: a
 * control's value, a column's lent width). The file holds only what they give, so the manifest
 * keeps each record of those properties as the room has it.
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
 * A container's properties that a record changes with no change record, by where the file writes
 * the container: its story part, its local name and its place among the part's containers of that
 * name. The file holds the properties the records give; the manifest keeps the container's own
 * and every pending record of them.
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

/** Properties as XML text, read back with the engine's bounded reader under their own name. */
const xmlOf = (root: OoxmlElement): string =>
  serializeOoxmlPart({ id: 'properties', ...propertiesPart(''), root });

/** The untracked records as the manifest lists them. */
export function untrackedJson(untracked: readonly CheckpointUntracked[]): unknown[] {
  return untracked.map(({ part, container, occurrence, target, properties, records }) => ({
    part,
    container,
    occurrence,
    target,
    ...(properties === undefined ? {} : { properties: xmlOf(properties) }),
    records: records.map(
      ({ suggestion, element, untracked, merge, recordUnchanged, tableCells }) => ({
        suggestion,
        ...(element === null ? {} : { element: xmlOf(element) }),
        ...(untracked === undefined ? {} : { untracked }),
        ...(merge === undefined ? {} : { merge }),
        ...(recordUnchanged === undefined ? {} : { recordUnchanged }),
        ...(tableCells === undefined ? {} : { tableCells }),
      })
    ),
  }));
}

/**
 * The listed untracked records, each checked: every value comes from storage. Each names a
 * suggestion of the manifest's list of `suggestions`, and a container once for each target.
 *
 * @throws Invalid when one does not read.
 */
export function readUntracked(value: unknown, suggestions: number): CheckpointUntracked[] {
  const keys = new Set<string>();

  // The manifest's byte cap bounds the list
  return list(value, 'untracked', 0, Number.MAX_SAFE_INTEGER).map((entry, k) => {
    const path = at('untracked', k);
    const f = fieldsOf(
      entry,
      path,
      new Set(['part', 'container', 'occurrence', 'target', 'properties', 'records'])
    );
    const part = str(required(f, 'part'), at(path, 'part'), MAX_FIELD_UNITS);
    const name = normalizePartName(part);
    if (!name.ok || name.partName !== part) {
      throw new Invalid('invalid-value', at(path, 'part'), 'not a part name');
    }

    const container = str(required(f, 'container'), at(path, 'container'), MAX_FIELD_UNITS);
    const target = str(required(f, 'target'), at(path, 'target'), MAX_FIELD_UNITS);
    if (!isCheckpointContainer(container) || !TARGETS.has(target)) {
      throw new Invalid('invalid-value', path, 'not a container and its properties');
    }

    const occurrence = int(required(f, 'occurrence'), at(path, 'occurrence'));
    const key = `${containerKey(part, container, occurrence)}\u0000${target}`;
    if (keys.has(key)) {
      throw new Invalid('invalid-value', path, 'named twice');
    }

    keys.add(key);
    const records = list(required(f, 'records'), at(path, 'records'), 1, MAX_LISTED);
    return {
      part,
      container,
      occurrence,
      // A name of the set above
      target: target as OrderedPropertyTarget,
      ...(f.has('properties')
        ? { properties: readProperties(f.get('properties'), at(path, 'properties'), target) }
        : {}),
      records: records.map((record, j) =>
        readRecord(record, at(at(path, 'records'), j), target, suggestions)
      ),
    };
  });
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
  const merge = f.has('merge')
    ? list(f.get('merge'), at(path, 'merge'), 0, MAX_LISTED).map((key, k) =>
        str(key, at(at(path, 'merge'), k), MAX_FIELD_UNITS)
      )
    : undefined;
  const cells = f.has('tableCells')
    ? list(f.get('tableCells'), at(path, 'tableCells'), 0, Number.MAX_SAFE_INTEGER).map((cell, k) =>
        int(cell, at(at(path, 'tableCells'), k))
      )
    : undefined;
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    element: f.has('element')
      ? readProperties(f.get('element'), at(path, 'element'), target)
      : null,
    ...(f.has('untracked') ? { untracked: bool(f.get('untracked'), at(path, 'untracked')) } : {}),
    ...(merge === undefined ? {} : { merge }),
    ...(f.has('recordUnchanged')
      ? { recordUnchanged: bool(f.get('recordUnchanged'), at(path, 'recordUnchanged')) }
      : {}),
    ...(cells === undefined ? {} : { tableCells: cells }),
  };
}

/** Properties of `target` read from their XML text; else they are not the properties named. */
function readProperties(value: unknown, path: string, target: string): OoxmlElement {
  const xml = str(value, path, Number.MAX_SAFE_INTEGER);
  const read = readOoxmlPart(xml, propertiesPart(path));
  const root = read.ok ? read.part.root : undefined;
  if (root === undefined || root.namespaceUri !== WML_NAMESPACE_URI || root.localName !== target) {
    throw new Invalid('invalid-value', path, 'not the properties it names');
  }

  // The part's own namespace declarations stay with it: the properties go into another part
  return { ...root, namespaceBindings: [] };
}
