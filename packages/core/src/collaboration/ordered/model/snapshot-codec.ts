// How a saved item list (checkpoint version 6) holds each value of the model: a codec writes a
// value as JSON, reads it back from an untrusted file, and finds the first difference between two
// values. The field table (`snapshot-fields.ts`) gives each field of each model type its codec.

import type {
  OoxmlAttribute,
  OoxmlElement,
  OoxmlNode,
  OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { isValidXmlText } from '../../../store/package/sinks.ts';
import { at, fieldsOf, Invalid, oneOf, plainFields, required } from '../change/read.ts';
import { readPerson, type OrderedPerson } from './person.ts';
import {
  attributesDifference,
  nodeDifference,
  type SnapshotNodeReader,
  type SnapshotNodeWriter,
} from './snapshot-nodes.ts';
import type { OrderedItem, OrderedStory, OrderedSuggestion } from './types.ts';

/** What a codec writes with: the node list, and the place of each item. */
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

/** UTF-16 units of an id or a name the model holds. */
export const ID_MAX = 4096;

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

/** A string of `min` to `max` units that XML can hold. */
export const xmlString = (min: number, max: number): SnapshotCodec<string> =>
  plain((json, path) => readXmlString(json, path, min, max));

/** An id or a name. */
export const id = xmlString(1, ID_MAX);

/** Text of any length the file holds. */
export const text = xmlString(0, Number.MAX_SAFE_INTEGER);

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

/** One of `values`. */
export const choice = <T extends string>(values: readonly T[]): SnapshotCodec<T> =>
  plain((json, path) => oneOf(json, path, values));

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

/** An object whose fields the table gives. Absent and undefined read alike. */
export function record<T>(fields: SnapshotFields<T>): SnapshotCodec<T> {
  const entries = Object.entries(fields as Record<string, SnapshotField<unknown>>).flatMap(
    ([key, field]) => ('stored' in field ? [[key, field] as const] : [])
  );
  const known = new Set(entries.map(([key]) => key));
  return {
    write: (value, w) => {
      const out: Record<string, unknown> = {};
      for (const [key, field] of entries) {
        const v = (value as Record<string, unknown>)[key];
        if (v !== undefined) {
          out[key] = field.stored.write(v, w);
        } else if (!field.optional) {
          throw new RangeError(`The model has no ${key}.`);
        }
      }

      return out;
    },
    read: (json, path, r) => {
      const f = fieldsOf(json, path, known);
      const out: Record<string, unknown> = {};
      for (const [key, field] of entries) {
        if (field.optional && !f.has(key)) {
          continue;
        }

        out[key] = field.stored.read(required(f, key), at(path, key), r);
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

/** One of several objects, by the value of their field `key`; each variant's own table. */
export function variants<T extends object, K extends keyof T & string>(
  key: K,
  tables: { readonly [N in T[K] & string]: SnapshotFields<Extract<T, Record<K, N>>> }
): SnapshotCodec<T> {
  const codecs = new Map<string, SnapshotCodec<T>>(
    Object.entries(tables).map(([name, fields]) => [name, record(fields as SnapshotFields<T>)])
  );
  const names = [...codecs.keys()];
  const codecOf = (value: T) => codecs.get(String(value[key]))!;
  return {
    write: (value, w) => codecOf(value).write(value, w),
    read: (json, path, r) => {
      const name = oneOf(plainFields(json, path).get(key), at(path, key), names);
      return codecs.get(name)!.read(json, path, r);
    },
    diff: (a, b, path) => (a[key] === b[key] ? codecOf(a).diff(a, b, path) : at(path, key)),
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
