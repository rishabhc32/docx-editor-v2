// The pending suggestions a checkpoint's manifest lists, and who typed their units (design D14).

import { WML_NAMESPACE_URI, type OoxmlAttribute } from '../../../store/package/ooxml-tree.ts';
import { ORDERED_NAMESPACE_PATTERN } from '../change/ids.ts';
import {
  at,
  fieldsOf,
  int,
  Invalid,
  list,
  oneOf,
  plainText,
  required,
  str,
} from '../change/read.ts';
import { ORDERED_CHANGE_LIMITS } from '../change/types.ts';
import { readPerson } from './person.ts';
import type { OrderedAttribution, OrderedSuggestionStatus } from './types.ts';

/**
 * A suggestion: who made it, and the `w:id`s the file writes for its marks, every role. A load
 * makes it once and reads each of those marks as it. Pending, but for one whose records another
 * section keeps (a decided record of properties an untracked record changes).
 */
export interface CheckpointSuggestion {
  readonly attribution: OrderedAttribution;
  readonly revisions: readonly string[];
  readonly status?: OrderedSuggestionStatus;
}

/**
 * Who typed the units a suggestion inserted, in the order the file writes them: runs of a count
 * and the namespace of the room change that typed them, or null for units no room change typed.
 */
export interface CheckpointTypers {
  /** The suggestion, by its place in the manifest's list. */
  readonly suggestion: number;
  readonly units: readonly (readonly [number, string | null])[];
}

/** Most other attributes of one suggestion, as a load keeps them from a revision mark. */
const MAX_ATTRIBUTES = 8;

/** Most UTF-16 units of a string field: a date, an attribute's namespace or value, a name. */
export const MAX_FIELD_UNITS = 256;

/** Most entries of a suggestion's revisions or a takeover's records. */
export const MAX_LISTED = 256;
const NAME = /^[A-Za-z_][A-Za-z0-9._-]{0,63}$/;

/** A revision's or a comment's `w:id`, as a checkpoint names it. */
export const WORD_ID = /^[\x21-\x7E]{1,64}$/;
const TRIPLE = new Set(['id', 'author', 'date']);
const DECIDED: readonly OrderedSuggestionStatus[] = ['accepted', 'rejected'];

/** The suggestions as the manifest lists them. */
export function suggestionsJson(suggestions: readonly CheckpointSuggestion[]): unknown[] {
  return suggestions.map(({ attribution, revisions, status }) => ({
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
    ...(status === undefined || status === 'pending' ? {} : { status }),
  }));
}

/**
 * The listed suggestions, each checked: every value comes from storage.
 *
 * @throws Invalid when one does not read.
 */
export function readSuggestions(value: unknown): CheckpointSuggestion[] {
  // The manifest's byte cap bounds the list
  return list(value, 'suggestions', 0, Number.MAX_SAFE_INTEGER).map((entry, k) =>
    readSuggestion(entry, at('suggestions', k))
  );
}

function readSuggestion(value: unknown, path: string): CheckpointSuggestion {
  const f = fieldsOf(
    value,
    path,
    new Set(['author', 'date', 'attributes', 'person', 'revisions', 'status'])
  );
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
    ...(f.has('status') ? { status: oneOf(f.get('status'), at(path, 'status'), DECIDED) } : {}),
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

/** The typers as the manifest lists them. */
export function typersJson(typers: readonly CheckpointTypers[]): unknown[] {
  return typers.map(({ suggestion, units }) => ({ suggestion, units }));
}

/**
 * The listed typers, each checked: a suggestion of the list, named once, and runs of at least one
 * unit, each typed by a namespace or by none.
 *
 * @throws Invalid when one does not read.
 */
export function readTypers(value: unknown, suggestions: number): CheckpointTypers[] {
  const named = new Set<number>();
  return list(value, 'typers', 0, Number.MAX_SAFE_INTEGER).map((entry, k) => {
    const path = at('typers', k);
    const f = fieldsOf(entry, path, new Set(['suggestion', 'units']));
    const suggestion = int(required(f, 'suggestion'), at(path, 'suggestion'), 0, suggestions - 1);
    if (named.has(suggestion)) {
      throw new Invalid('invalid-value', at(path, 'suggestion'), 'named twice');
    }

    named.add(suggestion);
    const units = list(required(f, 'units'), at(path, 'units'), 1, Number.MAX_SAFE_INTEGER);
    return { suggestion, units: units.map((run, j) => readRun(run, at(at(path, 'units'), j))) };
  });
}

/** One run of units: its count, and the namespace that typed them or null. */
function readRun(value: unknown, path: string): readonly [number, string | null] {
  const [count, namespace] = list(value, path, 2, 2);
  return [
    int(count, at(path, 0), 1),
    namespace === null
      ? null
      : matching(namespace, at(path, 1), ORDERED_NAMESPACE_PATTERN, 'a namespace'),
  ];
}

/** A string of at most 64 units that `pattern` matches; else it is not `what`. */
export function matching(value: unknown, path: string, pattern: RegExp, what: string): string {
  const text = str(value, path, 64);
  if (!pattern.test(text)) {
    throw new Invalid('invalid-value', path, `not ${what}`);
  }

  return text;
}
