// Readers for the values inside a change: addresses and places, suggestion ids, properties
// on their allowlists, a control's value, a link target. Each checks one value and throws
// `Invalid`; each returns a new value with its known fields only.

import { sanitizeHref } from '../../../store/package/sinks.ts';
import { isOrderedSuggestionId, ORDERED_NAMESPACE_PATTERN } from './ids.ts';
import { allowedChildren, type OrderedPropertyScope } from './properties.ts';
import {
  at,
  bool,
  fieldPath,
  fieldsOf,
  int,
  Invalid,
  isForbiddenKey,
  isPlainText,
  list,
  plainParagraphs,
  plainText,
  required,
  str,
  variantFields,
  type Fields,
} from './read.ts';
import {
  ORDERED_CHANGE_LIMITS as LIMITS,
  type OrderedSeenCard,
  type OrderedSeenSuggestion,
  type OrderedChangeAddress,
  type OrderedChangePlace,
  type OrderedChangeProperty,
  type OrderedChangePropertyChild,
  type OrderedControlValue,
} from './types.ts';

/** A bookmark name Word accepts: a letter or `_` first, then letters, digits and `_`. */
export const BOOKMARK_NAME = /^[\p{L}_][\p{L}\p{N}_]{0,127}$/u;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(\.\d{1,3})?Z)?$/;
const ATTRIBUTE_NAME = /^[A-Za-z][A-Za-z0-9]{0,31}$/;

export function address(value: unknown, path: string): OrderedChangeAddress {
  const fields = fieldsOf(value, path, new Set(['space', 'offset']));
  const space = str(required(fields, 'space'), at(path, 'space'), LIMITS.space);
  if (!/^[a-zA-Z]:/.test(space) || space.length < 3 || !isSpaceText(space))
    throw new Invalid('invalid-space', at(path, 'space'));
  return { space, offset: int(required(fields, 'offset'), at(path, 'offset')) };
}

/** A space name: plain text with no space character. */
function isSpaceText(space: string): boolean {
  return !space.includes(' ') && isPlainText(space);
}

export function place(value: unknown, path: string): OrderedChangePlace {
  if (typeof value === 'object' && value !== null && Object.hasOwn(value, 'mark')) {
    const fields = fieldsOf(value, path, new Set(['mark']));
    return { mark: address(fields.get('mark'), at(path, 'mark')) };
  }

  return address(value, path);
}

export function addresses(value: unknown, path: string): OrderedChangeAddress[] {
  return list(value, path, 1, LIMITS.targets).map((item, k) => address(item, at(path, k)));
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

export function suggestionId(value: unknown, path: string): string {
  const id = str(value, path, LIMITS.suggestion);
  if (!isOrderedSuggestionId(id)) throw new Invalid('invalid-suggestion', path);
  return id;
}

/** A comment's `w:id`: plain text, as a file may give any. */
export function commentId(value: unknown, path: string): string {
  return plainText(value, path, 1, LIMITS.commentId);
}

/** A comment's text, or a resolve's note: plain text in paragraphs. */
export function commentText(value: unknown, path: string): string {
  return plainParagraphs(value, path, 1, LIMITS.comment);
}

export function suggestionIds(value: unknown, path: string, min = 1): string[] {
  const ids = list(value, path, min, LIMITS.targets).map((item, k) =>
    suggestionId(item, at(path, k))
  );
  if (new Set(ids).size !== ids.length) throw new Invalid('invalid-suggestion', path, 'repeated');
  return ids;
}

/** The suggestions of a card an organiser saw, each with its version. */
function seenSuggestions(value: unknown, path: string, max: number): OrderedSeenSuggestion[] {
  return list(value, path, 1, max).map((item, k) => {
    const where = at(path, k);
    const fields = fieldsOf(item, where, new Set(['suggestion', 'version']));
    return {
      suggestion: suggestionId(required(fields, 'suggestion'), at(where, 'suggestion')),
      version: int(required(fields, 'version'), at(where, 'version')),
    };
  });
}

/**
 * The cards a decision names, as the organiser saw them: up to
 * {@link ORDERED_CHANGE_LIMITS}.decisions suggestions in all, no suggestion twice.
 */
export function seenCards(value: unknown, path: string): OrderedSeenCard[] {
  let budget = LIMITS.decisions;
  const out = list(value, path, 1, LIMITS.decisions).map((item, k) => {
    const where = at(path, k);
    const fields = fieldsOf(item, where, new Set(['suggestions']));
    const suggestions = seenSuggestions(
      required(fields, 'suggestions'),
      at(where, 'suggestions'),
      Math.max(budget, 1)
    );
    budget -= suggestions.length;
    if (budget < 0) throw new Invalid('too-many', path, `over ${LIMITS.decisions} suggestions`);
    return { suggestions };
  });
  const ids = out.flatMap((card) => card.suggestions.map((s) => s.suggestion));
  if (new Set(ids).size !== ids.length) throw new Invalid('invalid-suggestion', path, 'repeated');
  return out;
}

function attributes(value: unknown, path: string): Record<string, string> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Invalid('invalid-type', path, 'not an object');
  const keys = Object.keys(value);
  if (keys.length === 0) throw new Invalid('empty', path);
  if (keys.length > LIMITS.attributes)
    throw new Invalid('too-many', path, `over ${LIMITS.attributes}`);
  const names = new Set(keys);
  const fields = fieldsOf(value, path, names);
  const out: Record<string, string> = {};
  for (const name of [...names].sort()) {
    if (isForbiddenKey(name) || !ATTRIBUTE_NAME.test(name))
      throw new Invalid('invalid-attribute', at(path, name));
    const raw = fields.get(name);
    if (typeof raw !== 'string') throw new Invalid('invalid-type', at(path, name), 'not a string');
    out[name] = plainText(raw, at(path, name), 0, LIMITS.attributeValue);
  }

  return out;
}

function child(
  value: unknown,
  path: string,
  allowed: ReadonlySet<string>
): OrderedChangePropertyChild {
  const fields = fieldsOf(value, path, new Set(['localName', 'attributes']));
  const localName = str(required(fields, 'localName'), at(path, 'localName'), 32);
  if (!allowed.has(localName)) throw new Invalid('unknown-property', at(path, 'localName'));
  return fields.has('attributes')
    ? { localName, attributes: attributes(fields.get('attributes'), at(path, 'attributes')) }
    : { localName };
}

/**
 * Properties on a scope's allowlist, each named once. With `removals`, an entry may take its
 * property off (`remove: true`, alone): a format's, not a style's.
 */
export function properties(
  value: unknown,
  path: string,
  scope: OrderedPropertyScope,
  removals = false
): OrderedChangeProperty[] {
  const seen = new Set<string>();
  const known = new Set(['localName', 'attributes', 'children', ...(removals ? ['remove'] : [])]);
  return list(value, path, 1, LIMITS.properties).map((item, k) => {
    const where = at(path, k);
    const fields = fieldsOf(item, where, known);
    const localName = str(required(fields, 'localName'), at(where, 'localName'), 32);
    const children = allowedChildren(scope, localName);
    if (children === null) throw new Invalid('unknown-property', at(where, 'localName'));
    if (seen.has(localName)) throw new Invalid('repeated-property', at(where, 'localName'));
    seen.add(localName);
    if (fields.has('remove')) {
      if (!bool(fields.get('remove'), at(where, 'remove')))
        throw new Invalid('invalid-value', at(where, 'remove'), 'only true');
      for (const key of ['attributes', 'children'])
        if (fields.has(key)) throw new Invalid('unknown-field', at(where, key));
      return { localName, remove: true };
    }

    const property: {
      localName: string;
      attributes?: Record<string, string>;
      children?: OrderedChangePropertyChild[];
    } = { localName };
    if (fields.has('attributes'))
      property.attributes = attributes(fields.get('attributes'), at(where, 'attributes'));
    if (fields.has('children')) {
      if (children.size === 0) throw new Invalid('unknown-field', at(where, 'children'));
      property.children = list(
        fields.get('children'),
        at(where, 'children'),
        1,
        LIMITS.children
      ).map((entry, j) => child(entry, at(at(where, 'children'), j), children));
    }

    return property;
  });
}

/** Each kind of control value, and its own field. */
const CONTROL_FIELDS: Readonly<Record<OrderedControlValue['kind'], readonly string[]>> = {
  text: ['text'],
  listItem: ['value'],
  checkbox: ['checked'],
  date: ['iso'],
};

export function controlValue(value: unknown, path: string): OrderedControlValue {
  const { name: kind, fields } = variantFields(value, path, 'kind', CONTROL_FIELDS);
  switch (kind) {
    case 'text':
      return { kind, text: plainText(required(fields, 'text'), at(path, 'text'), 0, LIMITS.text) };
    case 'listItem': {
      const item = plainText(required(fields, 'value'), at(path, 'value'), 0, LIMITS.listItem);
      return { kind, value: item };
    }
    case 'checkbox':
      return { kind, checked: bool(required(fields, 'checked'), at(path, 'checked')) };
    case 'date': {
      const iso = str(required(fields, 'iso'), at(path, 'iso'), 24);
      if (!ISO_DATE.test(iso)) throw new Invalid('invalid-date', at(path, 'iso'));
      return { kind, iso };
    }
  }
}

/**
 * A link target: printable ASCII, unchanged by `sanitizeHref`, which refuses every scheme but
 * `http(s)`, `mailto`, `tel` and `ftp`. An address with other characters must be percent-encoded.
 */
export function href(value: unknown, path: string, max: number = LIMITS.href): string {
  const raw = str(value, path, max);
  const projection = sanitizeHref(raw);
  if (!projection.ok || projection.href !== raw || !/^[\x21-\x7E]+$/.test(raw))
    throw new Invalid('invalid-href', path);
  return raw;
}

/** A link's target: exactly one of an address and a bookmark name, and a tooltip. */
export function linkTarget(f: Fields): Record<string, unknown> {
  if (f.has('href') === f.has('anchor')) {
    throw new Invalid('missing-field', f.path, 'exactly one of href and anchor');
  }

  const tooltip = f.has('tooltip')
    ? { tooltip: plainText(f.get('tooltip'), fieldPath(f, 'tooltip'), 1, LIMITS.tooltip) }
    : {};
  if (f.has('href')) {
    return { href: href(f.get('href'), fieldPath(f, 'href')), ...tooltip };
  }

  const anchor = str(f.get('anchor'), fieldPath(f, 'anchor'), LIMITS.anchor);
  if (!BOOKMARK_NAME.test(anchor)) {
    throw new Invalid('invalid-value', fieldPath(f, 'anchor'));
  }

  return { anchor, ...tooltip };
}

/** An optional range `from` to `to`: both or neither. */
export function optionalRange(f: Fields): Record<string, unknown> {
  if (f.has('from') !== f.has('to')) {
    throw new Invalid('missing-field', f.path, 'both of from and to, or neither');
  }

  if (!f.has('from')) {
    return {};
  }

  return {
    from: place(f.get('from'), fieldPath(f, 'from')),
    to: place(f.get('to'), fieldPath(f, 'to')),
  };
}
