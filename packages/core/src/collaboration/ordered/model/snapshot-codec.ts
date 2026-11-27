// How a saved item list (checkpoint version 6) holds each value of the model: a codec writes a
// value as JSON, reads it back from an untrusted file, and finds the first difference between two
// values. The field table (`snapshot-fields.ts`) gives each field of each model type its codec.
// The reader reads every value in the order the writer wrote it: strings and nodes are written
// once and then named by their place (`snapshot-strings.ts`, `snapshot-nodes.ts`).

import type {
  OoxmlAttribute,
  OoxmlElement,
  OoxmlNode,
  OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { isValidXmlText } from '../../../store/package/sinks.ts';
import { at, Invalid } from '../change/read.ts';
import { readPerson, type OrderedPerson } from './person.ts';
import {
  attributesDifference,
  nodeDifference,
  type SnapshotNodeReader,
  type SnapshotNodeWriter,
} from './snapshot-nodes.ts';
import { ID_MAX } from './snapshot-strings.ts';
import type { OrderedItem, OrderedStory, OrderedSuggestion } from './types.ts';

/** What a codec writes with: the strings and nodes written so far, and the place of each item. */
export interface SnapshotWriteContext {
  readonly nodes: SnapshotNodeWriter;
  /** An item's story and place in it. */
  item(item: OrderedItem): readonly [number, number];
}

/** What a codec reads with. */
export interface SnapshotReadContext {
  readonly nodes: SnapshotNodeReader;
  /** The item a story and place name. */
  item(value: unknown, path: string): OrderedItem;
  /** The suggestions, put in the document before its stories: a tree sums by their status. */
  suggestions(suggestions: Map<string, OrderedSuggestion>): Map<string, OrderedSuggestion>;
  /** A story's tree of `items`, in order. */
  tree(items: readonly OrderedItem[]): OrderedStory['tree'];
  /** A story, indexed and put in the document. */
  story(story: Omit<OrderedStory, 'spaces'>): OrderedStory;
}

/** How a checkpoint holds one value. Methods, so a codec of a type serves its readonly forms. */
export interface SnapshotCodec<T> {
  write(value: T, w: SnapshotWriteContext): unknown;
  /** @throws Invalid when the value does not read. */
  read(json: unknown, path: string, r: SnapshotReadContext): T;
  /** The path of the first difference, or null. */
  diff(a: T, b: T, path: string): string | null;
}

/** A field the checkpoint stores, with its codec. */
export interface SnapshotStoredField<V> {
  readonly stored: SnapshotCodec<V>;
  /** The field may be absent. */
  readonly optional: boolean;
}

/** A field an open makes again from the stored ones: how. */
export interface SnapshotRebuiltField {
  readonly rebuilt: string;
}

export type SnapshotField<V> = SnapshotStoredField<V> | SnapshotRebuiltField;

/** Every key of `T`, each stored or rebuilt: a new field with no entry fails the type check. */
export type SnapshotFields<T> = {
  readonly [K in keyof T]-?: SnapshotField<Exclude<T[K], undefined>>;
};

export const stored = <V>(codec: SnapshotCodec<V>): SnapshotStoredField<V> => ({
  stored: codec,
  optional: false,
});

export const maybe = <V>(codec: SnapshotCodec<V>): SnapshotStoredField<V> => ({
  stored: codec,
  optional: true,
});

export const rebuilt = (how: string): SnapshotRebuiltField => ({ rebuilt: how });

/** A value JSON holds as it is: written and compared as it is, read by `read`. */
function plain<T>(read: (json: unknown, path: string) => T): SnapshotCodec<T> {
  return {
    write: (value) => value,
    read: (json, path) => read(json, path),
    diff: (a, b, path) => (a === b ? null : path),
  };
}

/**
 * A string of `min` to `max` units that XML can hold.
 *
 * @throws Invalid when it is not one.
 */
export function readXmlString(json: unknown, path: string, min: number, max: number): string {
  if (typeof json !== 'string') {
    throw new Invalid('invalid-type', path, 'not a string');
  }

  if (json.length < min || json.length > max || !isValidXmlText(json)) {
    throw new Invalid('invalid-text', path);
  }

  return json;
}

/** A string of `min` to `max` units that XML can hold, written once and then by its place. */
export const xmlString = (min: number, max: number): SnapshotCodec<string> => ({
  write: (value, w) => w.nodes.strings.string(value),
  read: (json, path, r) => {
    const value = r.nodes.strings.string(json, path);
    if (value.length < min || value.length > max) {
      throw new Invalid('invalid-text', path);
    }

    return value;
  },
  diff: (a, b, path) => (a === b ? null : path),
});

/** An id or a name. */
export const id = xmlString(1, ID_MAX);

/** Text of any length the file holds, written as it is. */
export const text = plain((json, path) => readXmlString(json, path, 0, Number.MAX_SAFE_INTEGER));

/** A safe integer from `min` to `max`. */
export const int = (min: number, max = Number.MAX_SAFE_INTEGER): SnapshotCodec<number> =>
  plain((json, path) => {
    if (typeof json !== 'number' || !Number.isSafeInteger(json) || json < min || json > max) {
      throw new Invalid('invalid-number', path);
    }

    return json;
  });

/** A count: a safe integer from 0. */
export const count = int(0);

/** A finite number. */
export const num: SnapshotCodec<number> = plain((json, path) => {
  if (typeof json !== 'number' || !Number.isFinite(json)) {
    throw new Invalid('invalid-number', path);
  }

  return json;
});

export const bool: SnapshotCodec<boolean> = plain((json, path) => {
  if (typeof json !== 'boolean') {
    throw new Invalid('invalid-type', path, 'not a boolean');
  }

  return json;
});

/** One of `values`, by its place in them. */
export function choice<T extends string>(values: readonly T[]): SnapshotCodec<T> {
  return {
    write: (value) => {
      const place = values.indexOf(value);
      if (place === -1) {
        throw new RangeError(`The model holds ${value}, not one of ${values.join(', ')}.`);
      }

      return place;
    },
    read: (json, path) => {
      const value =
        typeof json === 'number' && Number.isSafeInteger(json) ? values[json] : undefined;
      if (value === undefined) {
        throw new Invalid('invalid-value', path, `not one of ${values.join(', ')}`);
      }

      return value;
    },
    diff: (a, b, path) => (a === b ? null : path),
  };
}

/** A value or null. */
export function nullable<T>(codec: SnapshotCodec<T>): SnapshotCodec<T | null> {
  return {
    write: (value, w) => (value === null ? null : codec.write(value, w)),
    read: (json, path, r) => (json === null ? null : codec.read(json, path, r)),
    diff: (a, b, path) => {
      if (a === null || b === null) {
        return a === b ? null : path;
      }

      return codec.diff(a, b, path);
    },
  };
}

function arrayOf(json: unknown, path: string): readonly unknown[] {
  if (!Array.isArray(json)) {
    throw new Invalid('invalid-type', path, 'not a list');
  }

  return json;
}

/** The first difference between two lists, entry by entry, or null. */
function listDifference<T>(
  a: readonly T[],
  b: readonly T[],
  path: string,
  same: (x: T, y: T, path: string) => string | null
): string | null {
  if (a.length !== b.length) {
    return path;
  }

  for (let k = 0; k < a.length; k += 1) {
    const found = same(a[k]!, b[k]!, at(path, k));
    if (found !== null) {
      return found;
    }
  }

  return null;
}

/** Shared empty list: the model replaces its lists, never pushes to them. */
const NONE: readonly never[] = Object.freeze([]);

/** A list, in order. */
export function list<T>(codec: SnapshotCodec<T>): SnapshotCodec<readonly T[]> {
  return {
    write: (value, w) => value.map((entry) => codec.write(entry, w)),
    read: (json, path, r) => {
      const entries = arrayOf(json, path);
      return entries.length === 0 ? NONE : entries.map((e, k) => codec.read(e, at(path, k), r));
    },
    diff: (a, b, path) => listDifference(a, b, path, (x, y, p) => codec.diff(x, y, p)),
  };
}

/** A set of strings, in its order: a repeated entry does not read. */
export function set(codec: SnapshotCodec<string>): SnapshotCodec<Set<string>> {
  return {
    write: (value, w) => [...value].map((entry) => codec.write(entry, w)),
    read: (json, path, r) => {
      const out = new Set<string>();
      arrayOf(json, path).forEach((entry, k) => {
        const value = codec.read(entry, at(path, k), r);
        if (out.has(value)) {
          throw new Invalid('repeated-id', at(path, k));
        }

        out.add(value);
      });
      return out;
    },
    diff: (a, b, path) => listDifference([...a], [...b], path, (x, y, p) => codec.diff(x, y, p)),
  };
}

/** A map, in its order, as a list of pairs: a repeated key does not read. */
export function map<K, V>(
  key: SnapshotCodec<K>,
  value: SnapshotCodec<V>
): SnapshotCodec<Map<K, V>> {
  return {
    write: (m, w) => [...m].map(([k, v]) => [key.write(k, w), value.write(v, w)]),
    read: (json, path, r) => {
      const out = new Map<K, V>();
      arrayOf(json, path).forEach((entry, k) => {
        const pair = arrayOf(entry, at(path, k));
        if (pair.length !== 2) {
          throw new Invalid('invalid-value', at(path, k), 'not a pair');
        }

        const read = key.read(pair[0], at(at(path, k), 0), r);
        if (out.has(read)) {
          throw new Invalid('repeated-id', at(path, k));
        }

        out.set(read, value.read(pair[1], at(at(path, k), 1), r));
      });
      return out;
    },
    diff: (a, b, path) =>
      listDifference([...a], [...b], path, ([ka, va], [kb, vb], p) => {
        return key.diff(ka, kb, at(p, 0)) ?? value.diff(va, vb, at(p, 1));
      }),
  };
}

/** Entries in the order of their keys. */
const byKey = <V>(m: Map<string, V>): Map<string, V> =>
  new Map([...m].sort(([a], [b]) => (a < b ? -1 : 1)));

/**
 * A map whose order carries no meaning, written in the order of its keys: a key out of that
 * order, or repeated, does not read.
 */
export function sortedMap<V>(
  key: SnapshotCodec<string>,
  value: SnapshotCodec<V>
): SnapshotCodec<Map<string, V>> {
  const inner = map(key, value);
  return {
    write: (m, w) => inner.write(byKey(m), w),
    read: (json, path, r) => {
      const out = inner.read(json, path, r);
      const keys = [...out.keys()];
      const k = keys.findIndex((name, n) => n > 0 && keys[n - 1]! > name);
      if (k !== -1) {
        throw new Invalid('invalid-value', at(path, k), 'not in key order');
      }

      return out;
    },
    diff: (a, b, path) => inner.diff(byKey(a), byKey(b), path),
  };
}

/** A node of Word XML. */
export const node: SnapshotCodec<OoxmlNode> = {
  write: (value, w) => w.nodes.ref(value),
  read: (json, path, r) => r.nodes.node(json, path),
  diff: (a, b, path) => nodeDifference(a, b, path),
};

/** An element of Word XML. */
export const element: SnapshotCodec<OoxmlElement> = {
  write: (value, w) => w.nodes.ref(value),
  read: (json, path, r) => r.nodes.element(json, path),
  diff: (a, b, path) => nodeDifference(a, b, path),
};

/** A part: the base package's by its name, any other by value. */
export const part: SnapshotCodec<OoxmlPart> = {
  write: (value, w) => w.nodes.part(value),
  read: (json, path, r) => r.nodes.part(json, path),
  diff: (a, b, path) =>
    a.id === b.id && a.name === b.name && a.contentType === b.contentType
      ? nodeDifference(a.root, b.root, at(path, 'root'))
      : path,
};

/** Attributes of Word XML. */
export const attributes: SnapshotCodec<readonly OoxmlAttribute[]> = {
  write: (value, w) => w.nodes.attributes(value),
  read: (json, path, r) => r.nodes.attributes(json, path),
  diff: (a, b, path) => attributesDifference(a, b, path),
};

/** An item of a story, by its story and place. */
export const item: SnapshotCodec<OrderedItem> = {
  write: (value, w) => w.item(value),
  read: (json, path, r) => r.item(json, path),
  diff: (a, b, path) =>
    a.space === b.space && a.offset === b.offset && a.kind === b.kind ? null : path,
};

/** A person the room stamped. */
export const person: SnapshotCodec<OrderedPerson> = {
  write: ({ id: personId, name, org }) => ({ id: personId, name, org }),
  read: (json, path) => readPerson(json, path),
  diff: (a, b, path) => (a.id === b.id && a.name === b.name && a.org === b.org ? null : path),
};

/** The most fields a record holds: one bit each in its mask, a safe integer. */
const RECORD_MAX_FIELDS = 53;

/**
 * An object whose fields the table gives, as `[mask, ...values]`: the values of the fields it
 * holds, in the table's order, and the mask with one bit for each stored field (the first field
 * is 1) set when its value follows. Absent and undefined read alike. A field that holds its value
 * in `defaults` is left out, and an absent one reads as that value.
 */
export function record<T>(fields: SnapshotFields<T>, defaults: Partial<T> = {}): SnapshotCodec<T> {
  const entries = Object.entries(fields as Record<string, SnapshotField<unknown>>).flatMap(
    ([key, field]) => ('stored' in field ? [[key, field] as const] : [])
  );
  if (entries.length > RECORD_MAX_FIELDS) {
    throw new RangeError(`A record holds at most ${RECORD_MAX_FIELDS} stored fields.`);
  }

  const limit = 2 ** entries.length;
  const fallback: Readonly<Record<string, unknown>> = defaults;
  return {
    write: (value, w) => {
      const out: unknown[] = [0];
      let mask = 0;
      let bit = 1;
      for (const [key, field] of entries) {
        const v = (value as Record<string, unknown>)[key];
        if (v === undefined) {
          if (!field.optional) {
            throw new RangeError(`The model has no ${key}.`);
          }
        } else if (
          !Object.hasOwn(fallback, key) ||
          field.stored.diff(v, fallback[key], key) !== null
        ) {
          out.push(field.stored.write(v, w));
          mask += bit;
        }

        bit *= 2;
      }

      out[0] = mask;
      return out;
    },
    read: (json, path, r) => {
      const values = arrayOf(json, path);
      const first = values[0];
      if (
        typeof first !== 'number' ||
        !Number.isSafeInteger(first) ||
        first < 0 ||
        first >= limit
      ) {
        throw new Invalid('invalid-value', path, 'not a mask of its fields');
      }

      let mask = first;
      const out: Record<string, unknown> = {};
      let next = 1;
      for (const [key, field] of entries) {
        const held = mask % 2 === 1;
        mask = Math.floor(mask / 2);
        if (held) {
          if (next >= values.length) {
            throw new Invalid('missing-field', at(path, key));
          }

          out[key] = field.stored.read(values[next], at(path, key), r);
          next += 1;
        } else if (Object.hasOwn(fallback, key)) {
          out[key] = fallback[key];
        } else if (!field.optional) {
          throw new Invalid('missing-field', at(path, key));
        }
      }

      if (next !== values.length) {
        throw new Invalid('invalid-value', path, 'a value of no field');
      }

      return out as T;
    },
    diff: (a, b, path) => {
      for (const [key, field] of entries) {
        const va = (a as Record<string, unknown>)[key];
        const vb = (b as Record<string, unknown>)[key];
        if (va === undefined || vb === undefined) {
          if (va !== vb) {
            return at(path, key);
          }

          continue;
        }

        const found = field.stored.diff(va, vb, at(path, key));
        if (found !== null) {
          return found;
        }
      }

      return null;
    },
  };
}

/** A codec whose read goes on with `finish`. */
export function after<T>(
  codec: SnapshotCodec<T>,
  finish: (value: T, r: SnapshotReadContext) => T
): SnapshotCodec<T> {
  return { ...codec, read: (json, path, r) => finish(codec.read(json, path, r), r) };
}

/**
 * One of several objects, by the value of their field `key`; each variant's own table. Written as
 * `[place, record]`: the variant's place among the tables, and its record without `key`.
 */
export function variants<T extends object, K extends keyof T & string>(
  key: K,
  tables: { readonly [N in T[K] & string]: SnapshotFields<Extract<T, Record<K, N>>> }
): SnapshotCodec<T> {
  const names = Object.keys(tables);
  const codecs = Object.entries(tables).map(([name, fields]) =>
    record(fields as SnapshotFields<T>, { [key]: name } as Partial<T>)
  );
  const placeOf = (value: T) => names.indexOf(String(value[key]));
  return {
    write: (value, w) => {
      const place = placeOf(value);
      return [place, codecs[place]!.write(value, w)];
    },
    read: (json, path, r) => {
      const [place, fields, ...rest] = arrayOf(json, path);
      const codec =
        typeof place === 'number' && Number.isSafeInteger(place) ? codecs[place] : undefined;
      if (codec === undefined || rest.length > 0) {
        throw new Invalid('invalid-value', at(path, key), `not one of ${names.join(', ')}`);
      }

      return codec.read(fields, path, r);
    },
    diff: (a, b, path) =>
      a[key] === b[key] ? codecs[placeOf(a)]!.diff(a, b, path) : at(path, key),
  };
}

/** A fixed value: the field that names a variant. */
export const exactly = <T extends string | number | boolean>(value: T): SnapshotCodec<T> =>
  plain((json, path) => {
    if (json !== value) {
      throw new Invalid('invalid-value', path);
    }

    return value;
  });
