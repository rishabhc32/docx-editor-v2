// The pending suggestions a checkpoint's manifest lists, and the rules every section is read by
// (design D14).

import type { OoxmlAttribute } from '../../../store/package/ooxml-tree.ts';
import { isXmlNCName } from '../../../store/package/qname.ts';
import { isValidXmlText } from '../../../store/package/sinks.ts';
import { ORDERED_NAMESPACE_PATTERN } from '../change/ids.ts';
import { at, int, Invalid, list, oneOf, plainFields, required, str } from '../change/read.ts';
import { suggestionId } from '../change/values.ts';
import { isTripleAttribute, MAX_OTHER_ATTRIBUTES } from '../load/state.ts';
import { readPerson } from './person.ts';
import type { OrderedAttribution, OrderedSuggestionStatus } from './types.ts';

/**
 * A suggestion: its id, who made it, and the `w:id`s the file writes for its marks, every role. A
 * load makes it once, under its id, and reads each of those marks as it. Pending, or decided when
 * another section names it (a decided record of properties an untracked record changes).
 */
export interface CheckpointSuggestion {
  readonly id: string;
  readonly attribution: OrderedAttribution;
  readonly revisions: readonly string[];
  /**
   * The local name of the file's change record outside the stories that the suggestion is (a
   * style's `w:rPrChange`), whose `w:id` is the one in `revisions`; absent for marks in the stories.
   */
  readonly record?: string;
  readonly status?: OrderedSuggestionStatus;
  /**
   * The room's writes that touched it (`OrderedSuggestion.version`), which a decision names; 0
   * when absent.
   */
  readonly version?: number;
}

/** A list of the manifest, of at least `min` entries: the manifest's byte cap bounds it. */
export const manifestList = (value: unknown, path: string, min = 0): readonly unknown[] =>
  list(value, path, min, Number.MAX_SAFE_INTEGER);

/** A string of the manifest, not empty: the manifest's byte cap bounds it. */
export const manifestString = (value: unknown, path: string): string =>
  str(value, path, Number.MAX_SAFE_INTEGER);

/**
 * A section of the manifest, each entry read by `readEntry`, no two under one key.
 *
 * @throws Invalid when an entry does not read, or two share a key.
 */
export function readNamedOnce<T>(
  value: unknown,
  name: string,
  readEntry: (entry: unknown, path: string) => T,
  keyOf: (entry: T) => string | number
): T[] {
  const keys = new Set<string | number>();
  return manifestList(value, name).map((entry, k) => {
    const path = at(name, k);
    const read = readEntry(entry, path);
    const key = keyOf(read);
    if (keys.has(key)) {
      throw new Invalid('invalid-value', path, 'named twice');
    }

    keys.add(key);
    return read;
  });
}

const DECIDED: readonly OrderedSuggestionStatus[] = ['accepted', 'rejected'];

/** The suggestions as the manifest lists them. */
export function suggestionsJson(suggestions: readonly CheckpointSuggestion[]): unknown[] {
  return suggestions.map(({ id, attribution, revisions, record, status, version }) => ({
    id,
    author: attribution.author,
    ...(attribution.date === undefined ? {} : { date: attribution.date }),
    ...(attribution.attributes === undefined
      ? {}
      : {
          attributes: attribution.attributes.map(({ namespaceUri, localName, prefix, value }) => ({
            namespace: namespaceUri,
            name: localName,
            ...(prefix === undefined ? {} : { prefix }),
            value,
          })),
        }),
    ...(attribution.person === null ? {} : { person: attribution.person }),
    revisions,
    ...(record === undefined ? {} : { record }),
    ...(status === undefined || status === 'pending' ? {} : { status }),
    ...(version === undefined || version === 0 ? {} : { version }),
  }));
}

/**
 * The listed suggestions, each checked: every value comes from storage. The id is as a decision
 * names a suggestion, each named once; the author, the date and the other attributes are as a
 * load keeps them from a revision mark: any text valid in XML.
 *
 * @throws Invalid when one does not read.
 */
export function readSuggestions(value: unknown): CheckpointSuggestion[] {
  return readNamedOnce(value, 'suggestions', readSuggestion, ({ id }) => id);
}

function readSuggestion(value: unknown, path: string): CheckpointSuggestion {
  const f = plainFields(value, path);
  const revisions = manifestList(required(f, 'revisions'), at(path, 'revisions'));
  return {
    // As a decision names it
    id: suggestionId(required(f, 'id'), at(path, 'id')),
    attribution: {
      author: xmlText(required(f, 'author'), at(path, 'author')),
      person: f.has('person') ? readPerson(f.get('person'), at(path, 'person')) : null,
      ...(f.has('date') ? { date: xmlText(f.get('date'), at(path, 'date')) } : {}),
      ...(f.has('attributes')
        ? { attributes: readAttributes(f.get('attributes'), at(path, 'attributes')) }
        : {}),
    },
    revisions: revisions.map((id, k) => xmlText(id, at(at(path, 'revisions'), k))),
    ...(f.has('record') ? { record: readRecord(f.get('record'), path, revisions) } : {}),
    ...(f.has('status') ? { status: oneOf(f.get('status'), at(path, 'status'), DECIDED) } : {}),
    ...(f.has('version') ? { version: int(f.get('version'), at(path, 'version')) } : {}),
  };
}

/** A change record's local name, of a suggestion with its one `w:id`. */
function readRecord(value: unknown, path: string, revisions: readonly unknown[]): string {
  if (revisions.length !== 1) {
    throw new Invalid('invalid-value', at(path, 'revisions'), 'not the one id of a record');
  }

  return xmlName(value, at(path, 'record'));
}

/** A revision mark's attributes past its triple, as many as a load keeps, each named once. */
function readAttributes(value: unknown, path: string): OoxmlAttribute[] {
  const names = new Set<string>();
  return list(value, path, 0, MAX_OTHER_ATTRIBUTES).map((entry, k) => {
    const attribute = readAttribute(entry, at(path, k));
    const name = `${attribute.namespaceUri} ${attribute.localName}`;
    if (isTripleAttribute(attribute) || names.has(name)) {
      throw new Invalid('invalid-value', at(path, k), 'a triple attribute, or one named twice');
    }

    names.add(name);
    return attribute;
  });
}

/** A revision mark's attribute past its triple, such as `w16du:dateUtc`. */
function readAttribute(value: unknown, path: string): OoxmlAttribute {
  const f = plainFields(value, path);
  const prefix = f.has('prefix') ? xmlName(f.get('prefix'), at(path, 'prefix')) : undefined;
  return {
    kind: 'genericExtension',
    namespaceUri: xmlText(required(f, 'namespace'), at(path, 'namespace')),
    localName: xmlName(required(f, 'name'), at(path, 'name')),
    ...(prefix === undefined ? {} : { prefix }),
    value: xmlText(required(f, 'value'), at(path, 'value')),
  };
}

/**
 * A string valid in XML, empty or not, as the engine's XML reader gives an attribute's value: a
 * file's `w:id`, author or date.
 */
export function xmlText(value: unknown, path: string): string {
  if (typeof value !== 'string') {
    throw new Invalid('invalid-type', path, 'not a string');
  }

  if (!isValidXmlText(value)) {
    throw new Invalid('invalid-text', path);
  }

  return value;
}

/** An XML name with no colon, as the engine's XML reader gives a local name or a prefix. */
function xmlName(value: unknown, path: string): string {
  const name = xmlText(value, path);
  if (!isXmlNCName(name)) {
    throw new Invalid('invalid-value', path, 'not an XML name');
  }

  return name;
}

/**
 * A namespace of room changes.
 *
 * @throws Invalid when it is not one.
 */
export function readNamespace(value: unknown, path: string): string {
  const namespace = str(value, path, 64);
  if (!ORDERED_NAMESPACE_PATTERN.test(namespace)) {
    throw new Invalid('invalid-namespace', path);
  }

  return namespace;
}
