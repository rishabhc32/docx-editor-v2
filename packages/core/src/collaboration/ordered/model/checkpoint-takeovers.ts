// The takeovers a checkpoint's manifest keeps (design D5, policy 2).

import { serializeOoxmlPart } from '../../../store/package/ooxml-serialize.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlAttribute,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { normalizePartName } from '../../../store/package/opc-names.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { at, fieldsOf, int, Invalid, list, plainText, required, str } from '../change/read.ts';
import { ORDERED_CHANGE_LIMITS } from '../change/types.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { readPerson } from './person.ts';
import type { OrderedAttribution } from './types.ts';

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

/** A suggestion whose records another holds: who made it, and the `w:id`s of its own marks. */
export interface CheckpointSuggestion {
  readonly attribution: OrderedAttribution;
  /** The `w:id`s the file writes for its own marks: a load reads it from these, if any. */
  readonly revisions: readonly string[];
}

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

/** The records each change record of the file holds by takeover, and their suggestions. */
export interface CheckpointTakeovers {
  readonly takeovers: readonly CheckpointTakeover[];
  readonly suggestions: readonly CheckpointSuggestion[];
}

/** The key of a change record: its part, local name, `w:id` and occurrence. */
export const takeoverKey = (part: string, record: string, id: string, occurrence: number): string =>
  [part, record, id, occurrence].join('\u0000');

/**
 * Each change record of a part that may hold records by takeover, in document order, with its
 * `w:id` and its place among the part's records of its name and id.
 */
export function* changeRecordsOf(part: OoxmlPart): Generator<{
  readonly element: OoxmlElement;
  readonly id: string;
  readonly occurrence: number;
}> {
  const seen = new Map<string, number>();
  const stack: OoxmlNode[] = [part.root];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') {
      continue;
    }

    const id = attributeOf(node, 'id');
    if (
      id !== undefined &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      RECORDS.has(node.localName)
    ) {
      const name = `${node.localName}\u0000${id}`;
      const occurrence = seen.get(name) ?? 0;
      seen.set(name, occurrence + 1);
      yield { element: node, id, occurrence };
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
 * A takeover's properties as a part of their own: the manifest holds them as XML, read with the
 * engine's bounded reader. Each one is read under its own name, so its node ids are its own.
 */
const propertiesPart = (name: string) => ({
  name: `/ordered-checkpoint/${name}`,
  contentType: 'application/xml',
});

/** The takeovers and their suggestions as the manifest lists them. */
export function takeoversJson({ takeovers, suggestions }: CheckpointTakeovers): {
  readonly takeovers: unknown[];
  readonly suggestions: unknown[];
} {
  return {
    takeovers: takeovers.map(({ part, record, id, occurrence, taken }) => ({
      part,
      record,
      id,
      occurrence,
      taken: taken.map(({ suggestion, properties }) => ({
        suggestion,
        properties: serializeOoxmlPart({
          id: 'properties',
          ...propertiesPart(''),
          root: properties,
        }),
      })),
    })),
    suggestions: suggestions.map(({ attribution, revisions }) => ({
      author: attribution.author,
      ...(attribution.date === undefined ? {} : { date: attribution.date }),
      attributes: (attribution.attributes ?? []).map(
        ({ namespaceUri, localName, prefix, value }) => ({
          namespace: namespaceUri,
          name: localName,
          ...(prefix === undefined ? {} : { prefix }),
          value,
        })
      ),
      person: attribution.person,
      revisions,
    })),
  };
}

/** Most other attributes of one suggestion, as a load keeps them from a revision mark. */
const MAX_ATTRIBUTES = 8;

/** Most UTF-16 units of a string field: a date, an attribute's namespace or value, a name. */
const MAX_FIELD_UNITS = 256;

/** Most entries of a suggestion's revisions or a takeover's records. */
const MAX_LISTED = 256;
const NAME = /^[A-Za-z_][A-Za-z0-9._-]{0,63}$/;

/** A revision's or a comment's `w:id`, as a checkpoint names it. */
export const WORD_ID = /^[\x21-\x7E]{1,64}$/;
const TRIPLE = new Set(['id', 'author', 'date']);

/**
 * The listed takeovers and suggestions, each checked: every value comes from storage.
 *
 * @throws Invalid when one does not read.
 */
export function readTakeovers(takeovers: unknown, suggestions: unknown): CheckpointTakeovers {
  // The manifest's byte cap bounds the lists
  const read = list(suggestions, 'suggestions', 0, Number.MAX_SAFE_INTEGER).map((value, k) =>
    readSuggestion(value, at('suggestions', k))
  );
  const keys = new Set<string>();
  const out = list(takeovers, 'takeovers', 0, Number.MAX_SAFE_INTEGER).map((value, k) => {
    const takeover = readTakeover(value, at('takeovers', k), read.length);
    const key = takeoverKey(takeover.part, takeover.record, takeover.id, takeover.occurrence);
    if (keys.has(key)) {
      throw new Invalid('invalid-value', at('takeovers', k), 'named twice');
    }

    keys.add(key);
    return takeover;
  });
  return { takeovers: out, suggestions: read };
}

function readSuggestion(value: unknown, path: string): CheckpointSuggestion {
  const f = fieldsOf(value, path, new Set(['author', 'date', 'attributes', 'person', 'revisions']));
  const person = required(f, 'person');
  const attributes = list(required(f, 'attributes'), at(path, 'attributes'), 0, MAX_ATTRIBUTES);
  const revisions = list(required(f, 'revisions'), at(path, 'revisions'), 0, MAX_LISTED);
  const date = f.has('date')
    ? plainText(f.get('date'), at(path, 'date'), 1, MAX_FIELD_UNITS)
    : undefined;
  return {
    attribution: {
      author: plainText(required(f, 'author'), at(path, 'author'), 0, ORDERED_CHANGE_LIMITS.author),
      person: person === null ? null : readPerson(person, at(path, 'person')),
      ...(date === undefined ? {} : { date }),
      ...(attributes.length === 0
        ? {}
        : {
            attributes: attributes.map((a, k) => readAttribute(a, at(at(path, 'attributes'), k))),
          }),
    },
    revisions: revisions.map((id, k) =>
      matching(id, at(at(path, 'revisions'), k), WORD_ID, 'a revision id')
    ),
  };
}

/** A revision mark's attribute past its triple, such as `w16du:dateUtc`. */
function readAttribute(value: unknown, path: string): OoxmlAttribute {
  const f = fieldsOf(value, path, new Set(['namespace', 'name', 'prefix', 'value']));
  const namespaceUri = plainText(
    required(f, 'namespace'),
    at(path, 'namespace'),
    1,
    MAX_FIELD_UNITS
  );
  const localName = matching(required(f, 'name'), at(path, 'name'), NAME, 'an XML name');
  if (namespaceUri === WML_NAMESPACE_URI && TRIPLE.has(localName)) {
    throw new Invalid('invalid-value', at(path, 'name'), 'a revision triple attribute');
  }

  const prefix = f.has('prefix')
    ? matching(f.get('prefix'), at(path, 'prefix'), NAME, 'an XML name')
    : undefined;
  return {
    kind: 'genericExtension',
    namespaceUri,
    localName,
    ...(prefix === undefined ? {} : { prefix }),
    value: plainText(required(f, 'value'), at(path, 'value'), 0, MAX_FIELD_UNITS),
  };
}

function readTakeover(value: unknown, path: string, suggestions: number): CheckpointTakeover {
  const f = fieldsOf(value, path, new Set(['part', 'record', 'id', 'occurrence', 'taken']));
  const part = str(required(f, 'part'), at(path, 'part'), MAX_FIELD_UNITS);
  const name = normalizePartName(part);
  if (!name.ok || name.partName !== part) {
    throw new Invalid('invalid-value', at(path, 'part'), 'not a part name');
  }

  const record = str(required(f, 'record'), at(path, 'record'), MAX_FIELD_UNITS);
  if (!RECORDS.has(record)) {
    throw new Invalid('invalid-value', at(path, 'record'), 'not a change record');
  }

  const taken = list(required(f, 'taken'), at(path, 'taken'), 1, MAX_LISTED);
  return {
    part,
    record,
    id: matching(required(f, 'id'), at(path, 'id'), WORD_ID, 'a revision id'),
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
  const xml = str(required(f, 'properties'), at(path, 'properties'), Number.MAX_SAFE_INTEGER);
  const read = readOoxmlPart(xml, propertiesPart(path));
  const root = read.ok ? read.part.root : undefined;
  if (
    root === undefined ||
    root.namespaceUri !== WML_NAMESPACE_URI ||
    root.localName !== recordedProperties(record)
  ) {
    throw new Invalid('invalid-value', at(path, 'properties'), 'not the properties it records');
  }

  return {
    suggestion: int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1),
    properties: root,
  };
}

/** A string of at most 64 units that `pattern` matches; else it is not `what`. */
function matching(value: unknown, path: string, pattern: RegExp, what: string): string {
  const text = str(value, path, 64);
  if (!pattern.test(text)) {
    throw new Invalid('invalid-value', path, `not ${what}`);
  }

  return text;
}

/**
 * A takeover's properties as an element of their own: `properties` under the element `record`
 * names. It is never placed in a part.
 */
export const takenProperties = (record: string, properties: readonly OoxmlNode[]): OoxmlElement =>
  build('taken-properties', 'generic', recordedProperties(record), [], properties);
