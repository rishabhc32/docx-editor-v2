// The takeovers a checkpoint's manifest keeps (design D5, policy 2), the places of the file it
// names records by, and the properties it holds as XML.

import { serializeOoxmlPart } from '../../../store/package/ooxml-serialize.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { normalizePartName } from '../../../store/package/opc-names.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { at, fieldsOf, int, Invalid, required } from '../change/read.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { manifestList, manifestString, readNamedOnce, xmlText } from './checkpoint-suggestions.ts';

/**
 * The change records that hold records by takeover, each of the properties element it names. Word
 * keeps one for each properties element, under the latest author, so the file alone loses the
 * records a suggestion took from another: the manifest names each by where the file writes it.
 */
const RECORDS: ReadonlySet<string> = new Set([
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblGridChange',
  'tblPrExChange',
  'trPrChange',
  'tcPrChange',
]);

/** The properties element a change record names: `w:rPr` for `w:rPrChange`, and so on. */
export const recordedProperties = (record: string): string => record.slice(0, -'Change'.length);

/** A record a change record holds by takeover: its suggestion, and the properties it gives. */
export interface CheckpointTaken {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  /** The properties element with the record, without its own change record. */
  readonly properties: OoxmlElement;
}

/**
 * A change record that holds records by takeover, by where the file writes it: its story part,
 * its local name, its `w:id`, and its place among the part's records of that name and id in
 * document order (one `w:id` marks every record of a suggestion).
 */
export interface CheckpointTakeover {
  readonly part: string;
  readonly record: string;
  readonly id: string;
  readonly occurrence: number;
  /** The records it holds by takeover, in the order they are drawn. */
  readonly taken: readonly CheckpointTaken[];
}

/** The key of a change record: its part, local name, `w:id` and occurrence. */
export const takeoverKey = (part: string, record: string, id: string, occurrence: number): string =>
  [part, record, id, occurrence].join('\u0000');

/** The containers whose properties a record may change with no change record (`untracked`). */
const CONTAINERS: ReadonlySet<string> = new Set(['tbl', 'tr', 'tc', 'sdt']);

/** The key of a container: its part, local name and occurrence. */
export const containerKey = (part: string, container: string, occurrence: number): string =>
  [part, container, occurrence].join('\u0000');

/** Whether a local name is of a container {@link containerKey} names. */
export const isCheckpointContainer = (name: string): boolean => CONTAINERS.has(name);

/**
 * The places of a story part the manifest names records by, in document order, each with its key:
 * each change record that may hold records by takeover, by its `w:id` and its place among the
 * part's records of its name and id ({@link takeoverKey}), and each container by its place among
 * the part's containers of its name ({@link containerKey}).
 */
export function* checkpointPlacesOf(
  name: string,
  part: OoxmlPart
): Generator<{
  readonly element: OoxmlElement;
  readonly occurrence: number;
  readonly key: string;
}> {
  const seen = new Map<string, number>();
  const next = (counted: string): number => {
    const occurrence = seen.get(counted) ?? 0;
    seen.set(counted, occurrence + 1);
    return occurrence;
  };
  const stack: OoxmlNode[] = [part.root];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') {
      continue;
    }

    const id = attributeOf(node, 'id');
    const local = node.localName;
    if (node.namespaceUri === WML_NAMESPACE_URI && id !== undefined && RECORDS.has(local)) {
      const occurrence = next(`${local}\u0000${id}`);
      yield { element: node, occurrence, key: takeoverKey(name, local, id, occurrence) };
    } else if (node.namespaceUri === WML_NAMESPACE_URI && CONTAINERS.has(local)) {
      const occurrence = next(local);
      yield { element: node, occurrence, key: containerKey(name, local, occurrence) };
    }

    for (let k = node.children.length - 1; k >= 0; k -= 1) {
      const child = node.children[k];
      if (child !== undefined && child.kind !== 'textValue') {
        stack.push(child);
      }
    }
  }
}

/**
 * Properties as a part of their own: the manifest holds them as XML, read with the engine's
 * bounded reader. Each one is read under its own name, so its node ids are its own.
 */
const propertiesPart = (name: string) => ({
  name: `/ordered-checkpoint/${name}`,
  contentType: 'application/xml',
});

/** Properties as the manifest holds them: XML text. */
export const propertiesXml = (root: OoxmlElement): string =>
  serializeOoxmlPart({ id: 'properties', ...propertiesPart(''), root });

/**
 * Properties the manifest holds, read from their XML text; else they are not the `w:<localName>`
 * element.
 *
 * @throws Invalid when they do not read.
 */
export function readPropertiesXml(value: unknown, path: string, localName: string): OoxmlElement {
  const read = readOoxmlPart(manifestString(value, path), propertiesPart(path));
  const root = read.ok ? read.part.root : undefined;
  if (
    root === undefined ||
    root.namespaceUri !== WML_NAMESPACE_URI ||
    root.localName !== localName
  ) {
    throw new Invalid('invalid-value', path, `not w:${localName}`);
  }

  // The part's own namespace declarations stay with it: the properties go into another part
  return { ...root, namespaceBindings: [] };
}

/** The takeovers as the manifest lists them. */
export function takeoversJson(takeovers: readonly CheckpointTakeover[]): unknown[] {
  return takeovers.map(({ taken, ...takeover }) => ({
    ...takeover,
    taken: taken.map(({ suggestion, properties }) => ({
      suggestion,
      properties: propertiesXml(properties),
    })),
  }));
}

/**
 * The listed takeovers, each checked: every value comes from storage. Each names a suggestion of
 * the manifest's list of `suggestions`.
 *
 * @throws Invalid when one does not read.
 */
export function readTakeovers(takeovers: unknown, suggestions: number): CheckpointTakeover[] {
  return readNamedOnce(
    takeovers,
    'takeovers',
    (value, path) => readTakeover(value, path, suggestions),
    ({ part, record, id, occurrence }) => takeoverKey(part, record, id, occurrence)
  );
}

function readTakeover(value: unknown, path: string, suggestions: number): CheckpointTakeover {
  const f = fieldsOf(value, path, new Set(['part', 'record', 'id', 'occurrence', 'taken']));
  const part = manifestString(required(f, 'part'), at(path, 'part'));
  const name = normalizePartName(part);
  if (!name.ok || name.partName !== part) {
    throw new Invalid('invalid-value', at(path, 'part'), 'not a part name');
  }

  const record = manifestString(required(f, 'record'), at(path, 'record'));
  if (!RECORDS.has(record)) {
    throw new Invalid('invalid-value', at(path, 'record'), 'not a change record');
  }

  const taken = manifestList(required(f, 'taken'), at(path, 'taken'), 1);
  return {
    part,
    record,
    id: xmlText(required(f, 'id'), at(path, 'id')),
    occurrence: int(required(f, 'occurrence'), at(path, 'occurrence')),
    taken: taken.map((entry, k) => readTaken(entry, at(at(path, 'taken'), k), record, suggestions)),
  };
}

function readTaken(
  value: unknown,
  path: string,
  record: string,
  suggestions: number
): CheckpointTaken {
  const f = fieldsOf(value, path, new Set(['suggestion', 'properties']));
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    properties: readPropertiesXml(
      required(f, 'properties'),
      at(path, 'properties'),
      recordedProperties(record)
    ),
  };
}

/**
 * A takeover's properties as an element of their own: `properties` under the element `record`
 * names. It is never placed in a part.
 */
export const takenProperties = (record: string, properties: readonly OoxmlNode[]): OoxmlElement =>
  build('taken-properties', 'generic', recordedProperties(record), [], properties);
