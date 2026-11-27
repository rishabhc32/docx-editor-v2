// The strings of a saved item list (checkpoint version 6, design D14), as one stream that the
// writer and the reader walk in the same order. Each string is written once, at its first use;
// a later use names it by a short number:
// - an XML name (a kind, a namespace, a local name, a prefix): its place in the names table, or
//   the name itself at its first use;
// - any other string: `d`, the id of the base node `d` after the last base node named; `[d]`,
//   the string `d` after the last string named; the string itself at its first use; or `[a, b]`
//   at its first use, the join of two parts, each of the first three forms (`b:1:` and a base id).

import { isValidXmlText } from '../../../store/package/sinks.ts';
import { XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { Invalid } from '../change/read.ts';

/** UTF-16 units of an id or a name the model holds. */
export const ID_MAX = 4096;

/** Where a string may be cut for a join: after one of these. */
const CUTS = new Set([':', '@', '!', '/', '#', '.']);

/**
 * A place in a list, written as the difference from the last place of the same list, so places
 * near each other are short. The writer and the reader move it at each place they write or read.
 */
export class SnapshotChain {
  #last = 0;

  write(place: number): number {
    const difference = place - this.#last;
    this.#last = place;
    return difference;
  }

  /**
   * The place `json` names in a list of `size`.
   *
   * @throws Invalid when it is not a whole number or names no place.
   */
  read(json: unknown, size: number, path: string): number {
    const place = typeof json === 'number' ? this.#last + json : Number.NaN;
    if (!Number.isSafeInteger(json) || place < 0 || place >= size) {
      throw new Invalid('invalid-number', path, 'no such place');
    }

    this.#last = place;
    return place;
  }

  /** A new entry at `place` was written or read. */
  move(place: number): void {
    this.#last = place;
  }
}

/** Writes names and strings at their first use, and a short number at each later one. */
export class SnapshotStringWriter {
  readonly #base: (value: string) => number | undefined;
  readonly #baseChain: SnapshotChain;
  readonly #names = new Map<string, number>();
  readonly #strings = new Map<string, number>();
  readonly #chain = new SnapshotChain();

  /** `base` gives the number of the base node whose id is a string, on `baseChain`. */
  constructor(base: (value: string) => number | undefined, baseChain: SnapshotChain) {
    this.#base = base;
    this.#baseChain = baseChain;
  }

  name(value: string): number | string {
    const known = this.#names.get(value);
    if (known !== undefined) {
      return known;
    }

    this.#names.set(value, this.#names.size);
    return value;
  }

  string(value: string): unknown {
    const known = this.#known(value);
    if (known !== undefined) {
      return known;
    }

    const join = this.#join(value);
    this.#define(value);
    return join ?? value;
  }

  /** A base id or a string written before, or undefined. */
  #known(value: string): unknown {
    const base = this.#base(value);
    if (base !== undefined) {
      return this.#baseChain.write(base);
    }

    const place = this.#strings.get(value);
    return place === undefined ? undefined : [this.#chain.write(place)];
  }

  /** A part of a join: known, or written here. */
  #part(value: string): unknown {
    const known = this.#known(value);
    if (known !== undefined) {
      return known;
    }

    this.#define(value);
    return value;
  }

  /** `value` as a head and a known tail, cut at the first place that gives one, or undefined. */
  #join(value: string): unknown[] | undefined {
    if (value.length > ID_MAX) {
      return undefined;
    }

    for (let cut = 1; cut < value.length; cut += 1) {
      if (CUTS.has(value[cut - 1]!) && this.#isKnown(value.slice(cut))) {
        const head = this.#part(value.slice(0, cut));
        return [head, this.#known(value.slice(cut))];
      }
    }

    return undefined;
  }

  #isKnown(value: string): boolean {
    return this.#strings.has(value) || this.#base(value) !== undefined;
  }

  #define(value: string): void {
    const place = this.#strings.size;
    this.#strings.set(value, place);
    this.#chain.move(place);
  }
}

/** Reads names and strings as {@link SnapshotStringWriter} wrote them. */
export class SnapshotStringReader {
  readonly #base: readonly { readonly id: string }[];
  readonly #baseChain: SnapshotChain;
  readonly #names: string[] = [];
  readonly #strings: string[] = [];
  readonly #chain = new SnapshotChain();
  /** UTF-16 units the joins may still make: a join makes a new string. */
  #joinBudget = XML_HARD_MAX_BYTES;

  /** `base` holds the base nodes by number, which `baseChain` names. */
  constructor(base: readonly { readonly id: string }[], baseChain: SnapshotChain) {
    this.#base = base;
    this.#baseChain = baseChain;
  }

  /**
   * A name: XML text of at most {@link ID_MAX} units.
   *
   * @throws Invalid when it names no name or is not one.
   */
  name(json: unknown, path: string): string {
    if (typeof json === 'number') {
      const name = Number.isSafeInteger(json) ? this.#names[json] : undefined;
      if (name === undefined) {
        throw new Invalid('invalid-number', path, 'no such name');
      }

      return name;
    }

    const name = text(json, path, ID_MAX);
    this.#names.push(name);
    return name;
  }

  /**
   * A string: XML text.
   *
   * @throws Invalid when it names no string or is not one.
   */
  string(json: unknown, path: string): string {
    if (!Array.isArray(json) || json.length !== 2) {
      return this.#part(json, path);
    }

    const head = this.#part(json[0], path);
    const tail = this.#part(json[1], path);
    const value = head + tail;
    this.#joinBudget -= value.length;
    if (value.length > ID_MAX || this.#joinBudget < 0) {
      throw new Invalid('too-many', path, 'joined strings too long');
    }

    return this.#define(value);
  }

  /** A base id, a string read before, or a new string. */
  #part(json: unknown, path: string): string {
    if (typeof json === 'number') {
      return this.#base[this.#baseChain.read(json, this.#base.length, path)]!.id;
    }

    if (Array.isArray(json) && json.length === 1) {
      return this.#strings[this.#chain.read(json[0], this.#strings.length, path)]!;
    }

    return this.#define(text(json, path, XML_HARD_MAX_BYTES));
  }

  #define(value: string): string {
    this.#chain.move(this.#strings.push(value) - 1);
    return value;
  }
}

/**
 * XML text of at most `max` units.
 *
 * @throws Invalid when it is not.
 */
function text(json: unknown, path: string, max: number): string {
  if (typeof json !== 'string') {
    throw new Invalid('invalid-type', path, 'not a string');
  }

  if (json.length > max || !isValidXmlText(json)) {
    throw new Invalid('invalid-text', path);
  }

  return json;
}
