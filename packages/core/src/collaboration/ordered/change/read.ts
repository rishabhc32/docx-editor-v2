// Strict readers for untrusted JSON. Each one checks one value and throws `Invalid` with a
// stable code and the value's path; the validator turns the throw into a result. Readers never
// merge input into an object, read only own data properties (never a getter, never the
// prototype), and refuse `__proto__`, `constructor` and `prototype` as keys anywhere.

import { isValidXmlText } from '../../../store/package/sinks.ts';

/**
 * Why a change or a message was refused. Stable: a caller branches on it.
 *
 * - `invalid-json`, `frame-too-large`: a frame that is not JSON, or over the frame cap.
 * - `invalid-type`: a value of the wrong JSON type.
 * - `missing-field`, `unknown-field`: a required field that is absent; a field the shape does
 *   not define.
 * - `forbidden-key`: `__proto__`, `constructor` or `prototype` as a key.
 * - `server-field`: `n`, `by` or `at` in a change from a client.
 * - `invalid-version`: a version other than 1.
 * - `unknown-kind`, `unknown-type`: a change kind or a message type the format does not have.
 * - `invalid-id`, `foreign-id`, `repeated-id`: a change id that does not match the pattern, is
 *   outside the sender's namespace, or comes twice in one batch.
 * - `invalid-part`, `unknown-part`: a part name that is malformed, or not a story of the room.
 * - `invalid-space`, `invalid-number`, `invalid-text`, `invalid-value`, `invalid-date`,
 *   `invalid-suggestion`, `invalid-namespace`: a malformed value.
 * - `too-long`, `too-many`, `empty`: a string or a list over its cap, or empty.
 * - `unknown-property`, `repeated-property`, `invalid-attribute`: a property not on the
 *   allowlist, named twice, or with a malformed attribute.
 * - `invalid-href`: a link target that `sanitizeHref` refuses or changes.
 *
 * @public
 */
export type OrderedChangeErrorCode =
  | 'invalid-json'
  | 'frame-too-large'
  | 'invalid-type'
  | 'missing-field'
  | 'unknown-field'
  | 'forbidden-key'
  | 'server-field'
  | 'invalid-version'
  | 'unknown-kind'
  | 'unknown-type'
  | 'invalid-id'
  | 'foreign-id'
  | 'repeated-id'
  | 'invalid-part'
  | 'unknown-part'
  | 'invalid-space'
  | 'invalid-number'
  | 'invalid-text'
  | 'invalid-value'
  | 'invalid-date'
  | 'invalid-suggestion'
  | 'invalid-namespace'
  | 'too-long'
  | 'too-many'
  | 'empty'
  | 'unknown-property'
  | 'repeated-property'
  | 'invalid-attribute'
  | 'invalid-href';

/** A refusal of untrusted input: its code, where it is, and a short note for logs. @public */
export interface OrderedChangeError {
  readonly code: OrderedChangeErrorCode;
  /** Where the value is, such as `[2].properties[0].attributes.val`. */
  readonly path: string;
  readonly detail?: string;
}

/** A checked value, or the reason it was refused. @public */
export type OrderedChangeResult<T> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly error: OrderedChangeError };

/** The throw a reader makes. Caught by the validator, never seen by a caller. */
export class Invalid extends Error {
  readonly code: OrderedChangeErrorCode;
  readonly path: string;
  readonly detail: string | undefined;

  constructor(code: OrderedChangeErrorCode, path: string, detail?: string) {
    super(code);
    this.code = code;
    this.path = path;
    this.detail = detail;
  }
}

/** Run a reader; a refusal comes back as a result, any other throw goes on. */
export function attempt<T>(read: () => T): OrderedChangeResult<T> {
  try {
    return { ok: true, value: read() };
  } catch (error) {
    if (!(error instanceof Invalid)) throw error;
    const base = { code: error.code, path: error.path };
    return {
      ok: false,
      error: error.detail === undefined ? base : { ...base, detail: error.detail },
    };
  }
}

const FORBIDDEN_KEYS = new Set(['__proto__', 'constructor', 'prototype']);

/** Whether a key may never be used: it reaches an object's prototype. */
export function isForbiddenKey(key: string): boolean {
  return FORBIDDEN_KEYS.has(key);
}

/** A key as it shows in a path: at most 64 characters, printable ASCII only. */
export function keyLabel(key: string): string {
  const short = key.length > 64 ? `${key.slice(0, 64)}...` : key;
  return short.replace(/[^\x20-\x7E]/g, '?');
}

/** A child path: a key is labelled at every depth, the top one too. */
export function at(path: string, key: string | number): string {
  if (typeof key === 'number') {
    return `${path}[${key}]`;
  }

  const label = keyLabel(key);
  return path === '' ? label : `${path}.${label}`;
}

/** A plain object's own data fields, read without its prototype. */
export interface Fields {
  readonly path: string;
  has(key: string): boolean;
  get(key: string): unknown;
}

function isPlainObject(value: unknown): value is object {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const proto: unknown = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

/**
 * The fields of a plain object whose keys are all in `known`. `server` names fields only the
 * room may set: a client that sends one gets `server-field`.
 */
export function fieldsOf(
  value: unknown,
  path: string,
  known: ReadonlySet<string>,
  server: ReadonlySet<string> = new Set()
): Fields {
  if (!isPlainObject(value)) throw new Invalid('invalid-type', path, 'not an object');
  const keys = Object.keys(value);
  for (const key of keys)
    if (isForbiddenKey(key)) throw new Invalid('forbidden-key', at(path, key));
  for (const key of keys) if (server.has(key)) throw new Invalid('server-field', at(path, key));
  for (const key of keys) if (!known.has(key)) throw new Invalid('unknown-field', at(path, key));
  return plainFields(value, path);
}

/**
 * The fields of a plain object, every key allowed: for a checkpoint's manifest, whose reader skips
 * a field a newer engine added and refuses a key that reaches a prototype as it parses.
 */
export function plainFields(value: unknown, path: string): Fields {
  if (!isPlainObject(value)) {
    throw new Invalid('invalid-type', path, 'not an object');
  }

  return {
    path,
    has: (key) => Object.hasOwn(value, key),
    get: (key) => {
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (descriptor === undefined) {
        return undefined;
      }

      if (!('value' in descriptor)) {
        throw new Invalid('invalid-type', at(path, key), 'accessor');
      }

      return descriptor.value;
    },
  };
}

/** The path of a field. */
export function fieldPath(fields: Fields, key: string): string {
  return at(fields.path, key);
}

/** A required field. */
export function required(fields: Fields, key: string): unknown {
  if (!fields.has(key)) {
    throw new Invalid('missing-field', fieldPath(fields, key));
  }

  return fields.get(key);
}

/**
 * A plain object of one of several shapes, named by its `key` field: the fields of every shape
 * are checked first, then the shape's name, then exactly that shape's fields. `variants` maps
 * each name to its fields past `key`.
 */
export function variantFields<Name extends string>(
  value: unknown,
  path: string,
  key: string,
  variants: Readonly<Record<Name, readonly string[]>>
): { readonly name: Name; readonly fields: Fields } {
  const names = Object.keys(variants) as Name[];
  const head = fieldsOf(value, path, new Set([key, ...names.flatMap((name) => variants[name])]));
  const name = oneOf(required(head, key), at(path, key), names);
  return { name, fields: fieldsOf(value, path, new Set([key, ...variants[name]])) };
}

/** A string of 1 to `max` UTF-16 units. */
export function str(value: unknown, path: string, max: number): string {
  if (typeof value !== 'string') throw new Invalid('invalid-type', path, 'not a string');
  if (value.length === 0) throw new Invalid('empty', path);
  if (value.length > max) throw new Invalid('too-long', path, `over ${max}`);
  return value;
}

/** Control characters: C0 (with tab, line feed and return), DEL and C1. */
const CONTROL = /[\u0000-\u001F\u007F-\u009F]/;

/** Whether a string is plain text: no control character, valid in XML 1.0, pairs intact. */
export function isPlainText(value: string): boolean {
  return !CONTROL.test(value) && isValidXmlText(value);
}

/** A string of `min` to `max` UTF-16 units. */
function sized(value: unknown, path: string, min: number, max: number): string {
  if (typeof value !== 'string') throw new Invalid('invalid-type', path, 'not a string');
  if (value.length < min) throw new Invalid('empty', path);
  if (value.length > max) throw new Invalid('too-long', path, `over ${max}`);
  return value;
}

/** Plain text of `min` to `max` UTF-16 units. */
export function plainText(value: unknown, path: string, min: number, max: number): string {
  const text = sized(value, path, min, max);
  if (!isPlainText(text)) {
    throw new Invalid('invalid-text', path);
  }

  return text;
}

/** Plain text of `min` to `max` UTF-16 units in paragraphs: `\n` ends each but the last. */
export function plainParagraphs(value: unknown, path: string, min: number, max: number): string {
  const text = sized(value, path, min, max);
  if (!text.split('\n').every((line) => isPlainText(line))) {
    throw new Invalid('invalid-text', path);
  }

  return text;
}

/** A safe integer from `min` to `max`. */
export function int(value: unknown, path: string, min = 0, max = Number.MAX_SAFE_INTEGER): number {
  if (typeof value !== 'number') throw new Invalid('invalid-type', path, 'not a number');
  if (!Number.isSafeInteger(value) || value < min || value > max)
    throw new Invalid('invalid-number', path, `not a whole number from ${min} to ${max}`);

  // -0 reads as 0
  return value + 0;
}

/** A boolean. */
export function bool(value: unknown, path: string): boolean {
  if (typeof value !== 'boolean') throw new Invalid('invalid-type', path, 'not a boolean');
  return value;
}

/** One of `allowed`. */
export function oneOf<T extends string>(value: unknown, path: string, allowed: readonly T[]): T {
  if (typeof value !== 'string') throw new Invalid('invalid-type', path, 'not a string');
  const found = allowed.find((option) => option === value);
  if (found === undefined)
    throw new Invalid('invalid-value', path, `not one of ${allowed.join(', ')}`);
  return found;
}

/** An array of `min` to `max` items. */
export function list(value: unknown, path: string, min: number, max: number): readonly unknown[] {
  if (!Array.isArray(value)) throw new Invalid('invalid-type', path, 'not an array');
  if (value.length < min) throw new Invalid('empty', path);
  if (value.length > max) throw new Invalid('too-many', path, `over ${max}`);
  return value;
}
