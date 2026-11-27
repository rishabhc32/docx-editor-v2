// A checkpoint of version 6 is untrusted (design D14): a file that is not what the writer wrote is
// refused with its code, whatever field is wrong.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { createOrderedItem } from '../model/document.ts';
import { ITEM } from '../model/snapshot-fields.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { readSnapshotJson } from '../model/snapshot-read.ts';
import { headerText, type SnapshotHeader } from '../model/snapshot-write.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { ROOM_FACTS } from './checkpoint-fixture.ts';
import { resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkpointBytes } from './fold-check.ts';
import { documentView, itemJson, itemViews, type RecordView } from './snapshot-json.ts';

const ENTRY = 'ordered/checkpoint.json';

/** The room facts' checkpoint, as its zip entries. */
async function entries(): Promise<Record<string, Uint8Array>> {
  const { document: doc, base } = openOrderedWordFile(ROOM_FACTS.source, { file: 1 });
  ROOM_FACTS.edit(doc);
  return unzipSync(await checkpointBytes(doc, base, resourceBlobs(base)));
}

const FILE = entries();

type Json = Record<string, unknown>;

/**
 * The checkpoint with its JSON changed by `change` and its body's text by `edit`, the checksum
 * made again unless `raw`.
 */
async function changed(
  change: (header: Json, body: Json) => void,
  raw = false,
  edit = (text: string) => text
) {
  const files = { ...(await FILE) };
  const text = strFromU8(files[ENTRY]!);
  const value = JSON.parse(text) as Json;
  const body = value.body as Json;
  const header = { ...value };
  delete header.body;
  change(header, body);
  const bodyText = edit(JSON.stringify(body));
  const sha256 = raw ? header.sha256 : orderedBlobHash(strToU8(bodyText));
  const head = headerText({ ...(header as unknown as SnapshotHeader), sha256: String(sha256) });
  files[ENTRY] = strToU8(`${head}${bodyText}}`);
  return zipSync(files);
}

/** Why the checkpoint does not open, or 'opens'. */
function refusal(bytes: Uint8Array): string {
  try {
    openOrderedCheckpoint(bytes, { file: 2 });
    return 'opens';
  } catch (error) {
    if (error instanceof OrderedWordFileError) {
      return error.reason;
    }

    throw error;
  }
}

/** The checkpoint of the room facts, written from their document changed by `change`. */
async function written(change: (doc: OrderedDocument) => void): Promise<Uint8Array> {
  const { document: doc, base } = openOrderedWordFile(ROOM_FACTS.source, { file: 1 });
  ROOM_FACTS.edit(doc);
  change(doc);
  return checkpointBytes(doc, base, resourceBlobs(base));
}

/** The first item of the main story. */
const firstItem = (body: Json): RecordView => itemViews(body)[0]!;

/** The first plain object inside a value of the body. */
function firstObjectIn(value: unknown): object {
  const found = (json: unknown): object | undefined => {
    if (Array.isArray(json)) {
      return json.map(found).find(Boolean);
    }

    return typeof json === 'object' && json !== null ? json : undefined;
  };
  return found(value)!;
}

/** Two items of one space, in order, in the main story. */
function itemsOfOneSpace(doc: OrderedDocument): readonly [OrderedItem, OrderedItem] {
  const [story] = doc.stories.values();
  const items = [...story!.spaces.values()].find((list) => list.length > 1)!;
  return [items[0]!, items[1]!];
}

/** A case: its name, how it changes the file, its refusal, and whether its checksum stays. */
type Case = readonly [string, (header: Json, body: Json) => void, string, boolean?];

/** Last counters that do not read. */
const LAST_COUNTERS: readonly (readonly [string, [string, number][]])[] = [
  [
    'out of namespace order',
    [
      ['b', 1],
      ['a', 2],
    ],
  ],
  [
    'a repeated namespace',
    [
      ['a', 1],
      ['a', 2],
    ],
  ],
  ['a namespace past its pattern', [['a:b', 1]]],
  ['a counter of 0', [['a', 0]]],
  ['a counter past 15 digits', [['a', 10 ** 15]]],
];

describe('a checkpoint of version 6 that is not as written', () => {
  test('opens when only its checksum was made again', async () => {
    expect(refusal(await changed(() => undefined))).toBe('opens');
  });

  // Each change is the first value the reader meets that is not as written
  const cases: readonly Case[] = [
    [
      'a body that does not match its checksum',
      (_, body) => documentView(body).set('counters', [0]),
      'checkpoint-manifest',
      true,
    ],
    ['another format', (header) => (header.format = 'other'), 'checkpoint-format'],
    ['a version a later reader needs', (header) => (header.reads = 7), 'checkpoint-version'],
    ['a negative order number', (header) => (header.order = -1), 'checkpoint-manifest'],
    ['an order number in parts', (header) => (header.order = 1.5), 'checkpoint-manifest'],
    ['a missing field', (_, body) => documentView(body).delete('counters'), 'checkpoint-manifest'],
    [
      'an unknown field',
      (_, body) => {
        const document = body.document as number[];
        document[0] = document[0]! + 2 ** 30;
      },
      'checkpoint-manifest',
    ],
    [
      'a value of no field',
      (_, body) => (body.document as unknown[]).push(0),
      'checkpoint-manifest',
    ],
    ['a wrong type', (_, body) => firstItem(body).set('offset', '0'), 'checkpoint-manifest'],
    ['a negative number', (_, body) => firstItem(body).set('offset', -1), 'checkpoint-manifest'],
    [
      'an unsafe number',
      (_, body) => firstItem(body).set('length', 2 ** 60),
      'checkpoint-manifest',
    ],
    ['an unknown kind', (_, body) => firstItem(body).set('kind', 99), 'checkpoint-manifest'],
    ['a kind by its name', (_, body) => firstItem(body).set('kind', 'text'), 'checkpoint-manifest'],
    [
      'a base node out of range',
      (_, body) => firstItem(body).set('node', 99_999),
      'checkpoint-manifest',
    ],
    [
      'a listed node out of range',
      (_, body) => firstItem(body).set('node', [99_999]),
      'checkpoint-manifest',
    ],
    [
      'a repeated suggestion',
      (_, body) => {
        const suggestions = documentView(body).get('suggestions') as unknown[];
        suggestions.push(suggestions[0]);
      },
      'checkpoint-manifest',
    ],
    ...LAST_COUNTERS.map(
      ([name, counters]): Case => [
        `last counters ${name}`,
        (_, body) => documentView(body).set('lastCounters', counters),
        'checkpoint-manifest',
      ]
    ),
  ];
  for (const [name, change, reason, raw] of cases) {
    test(`refuses ${name}: ${reason}`, async () => {
      expect(refusal(await changed(change, raw))).toBe(reason);
    });
  }

  /** A document that breaks one of its rules: the writer writes it, the reader refuses it. */
  const rules: readonly (readonly [string, (doc: OrderedDocument) => void])[] = [
    [
      'a label of no suggestion',
      (doc) => {
        const [story] = doc.stories.values();
        story!.tree.toArray()[0]!.by = 'S:none';
      },
    ],
    [
      'two items over one unit',
      (doc) => {
        const [first, second] = itemsOfOneSpace(doc);
        (second as { offset: number }).offset = first.offset;
      },
    ],
    // A card never shows one name for another person
    [
      "a suggestion under another person's name",
      (doc) => {
        const suggestion = doc.suggestions.get('S:quin:1.1')!;
        suggestion.attribution = { ...suggestion.attribution, author: 'Pia' };
      },
    ],
    [
      "a comment under another person's name",
      (doc) => {
        const comment = doc.comments.get('1001645124')!;
        doc.comments.set('1001645124', { ...comment, author: 'Quin' });
      },
    ],
    [
      "an attribution before a change under another person's name",
      (doc) => {
        const { attribution } = doc.suggestions.get('S:quin:1.1')!;
        doc.revisionIds.reattributed.set('S:quin:1.1', { ...attribution, author: 'Pia' });
      },
    ],
  ];
  for (const [name, change] of rules) {
    test(`refuses ${name}: checkpoint-manifest`, async () => {
      expect(refusal(await written(() => undefined))).toBe('opens');
      expect(refusal(await written(change))).toBe('checkpoint-manifest');
    });
  }

  test("leaves out each item field that holds a new item's value, and reads it given", async () => {
    const { body } = JSON.parse(strFromU8((await FILE)[ENTRY]!)) as { body: Json };
    const { document: doc } = openOrderedCheckpoint(zipSync(await FILE), { file: 2 });
    const fresh = createOrderedItem({ space: 'b:1:x', kind: 'text' });
    const keys = Object.entries(ITEM).flatMap(([key, field]) => ('stored' in field ? [key] : []));
    const held = (item: OrderedItem) =>
      keys.filter(
        (key) =>
          key === 'space' ||
          key === 'kind' ||
          !Bun.deepEquals(item[key as keyof OrderedItem], fresh[key as keyof OrderedItem])
      ).length;
    const items = [...doc.stories.values()].flatMap((story) => story.tree.toArray());
    expect(itemJson(body).map((item) => item.length - 1)).toEqual(items.map(held));
    expect(refusal(await changed((_, b) => firstItem(b).set('dels', [])))).toBe('opens');
  });

  test('keeps the last counters of each namespace, in namespace order', async () => {
    const counters: [string, number][] = [
      ['a', 10 ** 15 - 1],
      ['b', 1],
    ];
    const bytes = await changed((_, body) => documentView(body).set('lastCounters', counters));
    const { document: doc } = openOrderedCheckpoint(bytes, { file: 2 });
    expect([...doc.lastCounters]).toEqual(counters);
  });

  test('refuses a prototype key anywhere', async () => {
    for (const key of ['__proto__', 'constructor', 'prototype']) {
      // The body, and an object inside it (a property's attributes), whose reader takes any other
      // key
      for (const place of ['body', 'inside'] as const) {
        const bytes = await changed((_, body) => {
          const target = place === 'body' ? body : firstObjectIn(body.document);
          Object.defineProperty(target, key, { value: '', enumerable: true });
        });
        expect({ key, place, reason: refusal(bytes) }).toEqual({
          key,
          place,
          reason: 'checkpoint-manifest',
        });
      }
    }
  });

  test('refuses a list past the limit of one part before it reads it', async () => {
    const { base } = openOrderedWordFile(ROOM_FACTS.source, { file: 1 });
    expect(readSnapshotJson(new Uint8Array(64 * 1024 * 1024 + 1), base)).toEqual({
      ok: false,
      reason: 'checkpoint-manifest',
    });
  });

  test('refuses a file with another entry', async () => {
    const files = { ...(await FILE), 'ordered/other.bin': new Uint8Array([1]) };
    expect(refusal(zipSync(files))).toBe('checkpoint-manifest');
  });
});
