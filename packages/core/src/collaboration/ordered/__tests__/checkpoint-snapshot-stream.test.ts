// The stream of a checkpoint of version 6 (design D14): strings, names and nodes written once and
// then named by a short number, and records as their values with a mask. A file is untrusted, so
// each form that is not as the writer writes it is refused by name where the reader meets it.

import { describe, expect, test } from 'bun:test';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { attempt } from '../change/read.ts';
import {
  choice,
  count,
  maybe,
  record,
  stored,
  variants,
  type SnapshotReadContext,
} from '../model/snapshot-codec.ts';
import { SnapshotNodeReader } from '../model/snapshot-nodes.ts';
import { ID_MAX } from '../model/snapshot-strings.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

const BASE = packageOf(docx(p(t('one'))));

/** The base nodes by number, as the stream numbers them: parts by name, then document order. */
const NODES: readonly OoxmlNode[] = (() => {
  const out: OoxmlNode[] = [];
  const walk = (node: OoxmlNode) => {
    out.push(node);
    if (node.kind !== 'textValue') {
      node.children.forEach(walk);
    }
  };
  [...BASE.parts.values()]
    .sort((a, b) => (a.name < b.name ? -1 : 1))
    .forEach((part) => {
      walk(part.root);
    });
  return out;
})();

const TEXT = NODES.findIndex((node) => node.kind === 'textValue');
const ELEMENT = NODES.findIndex((node) => node.kind !== 'textValue' && node.children.length > 0);

/** The code a read is refused with, or 'reads'. */
function codeOf(read: (reader: SnapshotNodeReader) => unknown): string {
  const reader = new SnapshotNodeReader(BASE);
  const result = attempt(() => read(reader));
  return result.ok ? 'reads' : result.error.code;
}

/** A context for codecs that name no string, node or item. */
const NO_CONTEXT = {} as SnapshotReadContext;

const codecCode = (read: () => unknown): string => {
  const result = attempt(read);
  return result.ok ? 'reads' : result.error.code;
};

describe('the strings of a checkpoint', () => {
  const cases: readonly (readonly [string, (reader: SnapshotNodeReader) => unknown, string])[] = [
    ['a base id by its number', (r) => r.strings.string(TEXT, 's'), 'reads'],
    [
      'a string, then the same by its place',
      (r) => [r.strings.string('a', 's'), r.strings.string([0], 's')],
      'reads',
    ],
    ['a join of a new head and a base id', (r) => r.strings.string(['b:1:', TEXT], 's'), 'reads'],
    ['a base number past the base', (r) => r.strings.string(NODES.length, 's'), 'invalid-number'],
    ['a base number in parts', (r) => r.strings.string(0.5, 's'), 'invalid-number'],
    ['a string not yet written', (r) => r.strings.string([0], 's'), 'invalid-number'],
    ['a string that is not XML text', (r) => r.strings.string('\u0001', 's'), 'invalid-text'],
    ['a join of three parts', (r) => r.strings.string(['a', 'b', 'c'], 's'), 'invalid-type'],
    ['a join inside a join', (r) => r.strings.string([['a', 'b'], 'c'], 's'), 'invalid-type'],
    [
      'a join past the id limit',
      (r) => r.strings.string(['x'.repeat(ID_MAX), 'y'], 's'),
      'too-many',
    ],
    [
      'a name, then the same by its place',
      (r) => [r.strings.name('w', 'n'), r.strings.name(0, 'n')],
      'reads',
    ],
    ['a name not yet written', (r) => r.strings.name(0, 'n'), 'invalid-number'],
    [
      'a name past the id limit',
      (r) => r.strings.name('x'.repeat(ID_MAX + 1), 'n'),
      'invalid-text',
    ],
  ];
  for (const [name, read, code] of cases) {
    test(`${name}: ${code}`, () => {
      expect(codeOf(read)).toBe(code);
    });
  }

  test('refuses joins that make more than the limit of one part: too-many', () => {
    // Each join makes a new string of the id limit from one string written once
    const joins = (reader: SnapshotNodeReader, count: number) => {
      reader.strings.string('x'.repeat(ID_MAX / 2), 's');
      for (let k = 1; k <= count; k += 1) {
        reader.strings.string([[1 - k], [0]], 's');
      }
    };
    const limit = (64 * 1024 * 1024) / ID_MAX;
    expect(codeOf((reader) => joins(reader, limit))).toBe('reads');
    expect(codeOf((reader) => joins(reader, limit + 1))).toBe('too-many');
  });
});

describe('the nodes of a checkpoint', () => {
  const deep = (depth: number): unknown =>
    depth === 0 ? ELEMENT : [depth === 1 ? ELEMENT : 0, [deep(depth - 1)]];
  const cases: readonly (readonly [string, unknown, string])[] = [
    ['a base node by its number', ELEMENT, 'reads'],
    ['a base element with other children', [ELEMENT, [TEXT - ELEMENT]], 'reads'],
    ['a new text node', ['t:1', 'one'], 'reads'],
    ['a base number past the base', NODES.length, 'invalid-number'],
    ['a listed node not yet read', [0], 'invalid-number'],
    ['a base element of a text node', [TEXT, []], 'invalid-id'],
    ['a base element by a listed node', [[0], []], 'invalid-id'],
    ['a node of another length', [ELEMENT, [], []], 'invalid-value'],
    ['a text node that is not XML text', ['t:1', '\u0001'], 'invalid-text'],
    ['nodes past the depth limit', deep(258), 'too-many'],
    [
      'an attribute of an unknown kind',
      ['e:1', 'generic', 'urn:x', 'e', null, [], [[9, 'urn:x', 'a', null, 'v']], []],
      'invalid-value',
    ],
    [
      'an element named as a text node',
      ['e:1', 'textValue', 'urn:x', 'e', null, [], [], []],
      'invalid-value',
    ],
  ];
  for (const [name, json, code] of cases) {
    test(`${name}: ${code}`, () => {
      expect(codeOf((reader) => reader.node(json, 'n'))).toBe(code);
    });
  }

  test('a listed node read twice is one node', () => {
    const reader = new SnapshotNodeReader(BASE);
    const first = reader.node(['t:1', 'one'], 'n');
    expect(reader.node([0], 'n')).toBe(first);
  });
});

describe('the records of a checkpoint', () => {
  const PAIR = record<{ a: number; b?: number }>({ a: stored(count), b: maybe(count) });
  const KIND = choice(['x', 'y']);
  const SHAPE = variants<{ kind: 'x'; n: number } | { kind: 'y' }, 'kind'>('kind', {
    x: { kind: stored(choice(['x'])), n: stored(count) },
    y: { kind: stored(choice(['y'])) },
  });
  const cases: readonly (readonly [string, () => unknown, string])[] = [
    ['a record with its mask', () => PAIR.read([3, 1, 2], 'r', NO_CONTEXT), 'reads'],
    ['a mask in parts', () => PAIR.read([1.5, 1], 'r', NO_CONTEXT), 'invalid-value'],
    ['a mask past its fields', () => PAIR.read([4, 1], 'r', NO_CONTEXT), 'invalid-value'],
    ['a field it must hold left out', () => PAIR.read([2, 1], 'r', NO_CONTEXT), 'missing-field'],
    ['fewer values than its mask', () => PAIR.read([3, 1], 'r', NO_CONTEXT), 'missing-field'],
    ['a value of no field', () => PAIR.read([1, 1, 2], 'r', NO_CONTEXT), 'invalid-value'],
    ['a record that is not a list', () => PAIR.read({ a: 1 }, 'r', NO_CONTEXT), 'invalid-type'],
    ['a choice by its place', () => KIND.read(1, 'c', NO_CONTEXT), 'reads'],
    ['a choice past its values', () => KIND.read(2, 'c', NO_CONTEXT), 'invalid-value'],
    ['a choice by its name', () => KIND.read('x', 'c', NO_CONTEXT), 'invalid-value'],
    ['a variant by its place', () => SHAPE.read([0, [2, 5]], 'v', NO_CONTEXT), 'reads'],
    ['a variant past the variants', () => SHAPE.read([2, [0]], 'v', NO_CONTEXT), 'invalid-value'],
    ['a variant with its own key', () => SHAPE.read([1, [1, 0]], 'v', NO_CONTEXT), 'reads'],
    ['a variant with another key', () => SHAPE.read([1, [1, 1]], 'v', NO_CONTEXT), 'invalid-value'],
    ['a variant with more parts', () => SHAPE.read([1, [0], 0], 'v', NO_CONTEXT), 'invalid-value'],
  ];
  for (const [name, read, code] of cases) {
    test(`${name}: ${code}`, () => {
      expect(codecCode(read)).toBe(code);
    });
  }

  test('a variant leaves out its key, and reads it back', () => {
    const value = { kind: 'x', n: 5 } as const;
    expect(SHAPE.write(value, {} as never)).toEqual([0, [2, 5]]);
    expect(SHAPE.read([0, [2, 5]], 'v', NO_CONTEXT)).toEqual(value);
  });
});
