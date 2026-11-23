// The takeovers a checkpoint's manifest keeps (design D5, policy 2), the elements of the file it
// names by label, and the properties it holds as XML.

import { serializeOoxmlPart } from '../../../store/package/ooxml-serialize.ts';
import { withRequiredNamespaceBindings } from '../../../store/store/tree-op-fragment.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { at, int, Invalid, plainFields, required } from '../change/read.ts';
import { isWml } from '../ooxml-nodes.ts';
import { manifestList, manifestString, readNamedOnce } from './checkpoint-suggestions.ts';

/**
 * The change records that hold records by takeover, each of the properties element it names. Word
 * keeps one for each properties element, under the latest author, so the file alone loses the
 * records a suggestion took from another: the manifest names each by the label on it.
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

/** The properties elements a change record of {@link RECORDS} names. */
const PROPERTIES: ReadonlySet<string> = new Set([...RECORDS].map(recordedProperties));

/**
 * A change record that holds records by takeover, by the label on it: one `w:id` marks every
 * record of a suggestion, so the id names no one record.
 */
export interface CheckpointTakeover {
  readonly label: number;
  /** The records it holds by takeover, in the order they are drawn. */
  readonly taken: readonly CheckpointTaken[];
}

/**
 * Whether an element is a change record of the takeover's properties: the element its label is
 * on must be one.
 */
export const holdsTakeover = (element: OoxmlElement, takeover: CheckpointTakeover): boolean =>
  RECORDS.has(element.localName) &&
  isWml(element, element.localName) &&
  takeover.taken.every(
    ({ properties }) => properties.localName === recordedProperties(element.localName)
  );

/**
 * Properties as a part of their own: the manifest holds them as XML, read with the engine's
 * bounded reader. Each one is read under its own name, so its node ids are its own.
 */
const propertiesPart = (name: string) => ({
  name: `/ordered-checkpoint/${name}`,
  contentType: 'application/xml',
});

/**
 * Properties as the manifest holds them: XML text, with each prefix they use declared on them, so
 * a read gives back the prefixes of the part they came from (`w14:checkbox`, not a new one).
 */
export const propertiesXml = (root: OoxmlElement): string => {
  const part = { id: 'properties', ...propertiesPart(''), root };
  return serializeOoxmlPart(withRequiredNamespaceBindings(part, [root]));
};

/**
 * Properties the manifest holds, read from their XML text; else they are not a `w:` element of
 * one of `localNames`.
 *
 * @throws Invalid when they do not read.
 */
export function readPropertiesXml(
  value: unknown,
  path: string,
  localNames: ReadonlySet<string>
): OoxmlElement {
  const read = readOoxmlPart(manifestString(value, path), propertiesPart(path));
  const root = read.ok ? read.part.root : undefined;
  if (
    root === undefined ||
    root.namespaceUri !== WML_NAMESPACE_URI ||
    !localNames.has(root.localName)
  ) {
    throw new Invalid('invalid-value', path, 'not the properties of a record');
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
    ({ label }) => label
  );
}

function readTakeover(value: unknown, path: string, suggestions: number): CheckpointTakeover {
  const f = plainFields(value, path);
  const taken = manifestList(required(f, 'taken'), at(path, 'taken'), 1);
  return {
    label: int(required(f, 'label'), at(path, 'label')),
    taken: taken.map((entry, k) => readTaken(entry, at(at(path, 'taken'), k), suggestions)),
  };
}

function readTaken(value: unknown, path: string, suggestions: number): CheckpointTaken {
  const f = plainFields(value, path);
  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    properties: readPropertiesXml(required(f, 'properties'), at(path, 'properties'), PROPERTIES),
  };
}

/**
 * A takeover's properties as an element of their own: `properties` under the element `record`
 * names. It is never placed in a part.
 */
export const takenProperties = (record: string, properties: readonly OoxmlNode[]): OoxmlElement =>
  build('taken-properties', 'generic', recordedProperties(record), [], properties);
