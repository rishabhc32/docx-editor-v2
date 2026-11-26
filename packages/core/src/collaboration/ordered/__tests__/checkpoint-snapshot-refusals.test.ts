// A checkpoint of version 6 is untrusted (design D14): a file that is not what the writer wrote is
// refused with its code, whatever field is wrong.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { createOrderedItem } from '../model/document.ts';
import { readSnapshotJson } from '../model/snapshot-read.ts';
import { headerText, type SnapshotHeader } from '../model/snapshot-write.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { ROOM_FACTS } from './checkpoint-fixture.ts';
import { resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkpointBytes } from './fold-check.ts';

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

/** The first item of the main story. */
const firstItem = (body: Json): Json => {
  const stories = (body.document as Json).stories as [string, Json][];
  return (stories[0]![1].tree as Json[])[0]!;
};

/** The value of a map entry of the document, by its key. */
const entryOf = (body: Json, map: string, key: string): Json => {
  const entries = (map.split('.').reduce((o, k) => (o as Json)[k], body.document) ?? []) as [
    string,
    Json,
  ][];
  return entries.find(([k]) => k === key)![1];
};

/** The first node the list holds as a base element's shell (its id alone). */
const firstShell = (body: Json): unknown[] =>
  (body.nodes as unknown[][]).find((entry) => entry.length === 1)!;

/** The id of a text node of the base's main part. */
const BASE_TEXT_ID = (() => {
  const { base } = openOrderedWordFile(ROOM_FACTS.source, { file: 1 });
  const find = (node: OoxmlNode): string | null =>
    node.kind === 'textValue' ? node.id : (node.children.map(find).find(Boolean) ?? null);
  return find(base.parts.get('/word/document.xml')!.root)!;
})();

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

  const cases: readonly Case[] = [
    [
      'a body that does not match its checksum',
      (_, body) => (body.namespaces = []),
      'checkpoint-manifest',
      true,
    ],
    ['another format', (header) => (header.format = 'other'), 'checkpoint-format'],
    ['a version a later reader needs', (header) => (header.reads = 7), 'checkpoint-version'],
    ['a negative order number', (header) => (header.order = -1), 'checkpoint-manifest'],
    ['an order number in parts', (header) => (header.order = 1.5), 'checkpoint-manifest'],
    [
      'a missing field',
      (_, body) => delete (body.document as Json).counters,
      'checkpoint-manifest',
    ],
    ['an unknown field', (_, body) => ((body.document as Json).extra = 1), 'checkpoint-manifest'],
    ['a wrong type', (_, body) => (firstItem(body).offset = '0'), 'checkpoint-manifest'],
    ['a negative number', (_, body) => (firstItem(body).offset = -1), 'checkpoint-manifest'],
    ['an unsafe number', (_, body) => (firstItem(body).length = 2 ** 60), 'checkpoint-manifest'],
    ['an unknown kind', (_, body) => (firstItem(body).kind = 'word'), 'checkpoint-manifest'],
    [
      'a label of no suggestion',
      (_, body) => (firstItem(body).by = 'S:none'),
      'checkpoint-manifest',
    ],
    ['a node out of range', (_, body) => (firstItem(body).node = 99_999), 'checkpoint-manifest'],
    [
      'a base node it does not hold',
      (_, body) => (firstItem(body).node = '/word/document.xml#9'),
      'checkpoint-manifest',
    ],
    [
      'a shell of no base element',
      (_, body) => (firstShell(body)[0] = '/word/document.xml#9'),
      'checkpoint-manifest',
    ],
    [
      'a shell of a text node',
      (_, body) => (firstShell(body)[0] = BASE_TEXT_ID),
      'checkpoint-manifest',
    ],
    ['a shell by a listed node', (_, body) => (firstShell(body)[0] = 0), 'checkpoint-manifest'],
    [
      'a repeated suggestion',
      (_, body) => {
        const suggestions = (body.document as Json).suggestions as unknown[];
        suggestions.push(suggestions[0]);
      },
      'checkpoint-manifest',
    ],
    [
      'two items over one unit',
      (_, body) => {
        const stories = (body.document as Json).stories as [string, Json][];
        const tree = stories[0]![1].tree as Json[];
        tree.splice(1, 0, { ...tree[0]! });
      },
      'checkpoint-manifest',
    ],
    // A card never shows one name for another person
    [
      "a suggestion under another person's name",
      (_, body) =>
        ((entryOf(body, 'suggestions', 'S:quin:1.1').attribution as Json).author = 'Pia'),
      'checkpoint-manifest',
    ],
    [
      "a comment under another person's name",
      (_, body) => (entryOf(body, 'comments', '1001645124').author = 'Quin'),
      'checkpoint-manifest',
    ],
    [
      "an attribution before a change under another person's name",
      (_, body) => {
        const { attribution } = entryOf(body, 'suggestions', 'S:quin:1.1');
        const revisionIds = (body.document as Json).revisionIds as Json;
        revisionIds.reattributed = [['S:quin:1.1', { ...(attribution as Json), author: 'Pia' }]];
      },
      'checkpoint-manifest',
    ],
    ...LAST_COUNTERS.map(
      ([name, counters]): Case => [
        `last counters ${name}`,
        (_, body) => ((body.document as Json).lastCounters = counters),
        'checkpoint-manifest',
      ]
    ),
  ];
  for (const [name, change, reason, raw] of cases) {
    test(`refuses ${name}: ${reason}`, async () => {
      expect(refusal(await changed(change, raw))).toBe(reason);
    });
  }

  test("leaves out each item field that holds a new item's value, and reads it given", async () => {
    const { body } = JSON.parse(strFromU8((await FILE)[ENTRY]!)) as { body: Json };
    const stories = (body.document as Json).stories as [string, Json][];
    const items = stories.flatMap(([, story]) => story.tree as Json[]);
    const fresh = createOrderedItem({ space: 'b:1:x', kind: 'text' }) as unknown as Json;
    const given = items.flatMap((item) =>
      Object.keys(item).filter((key) => key !== 'kind' && Bun.deepEquals(item[key], fresh[key]))
    );
    expect(given).toEqual([]);
    expect(refusal(await changed((_, b) => Object.assign(firstItem(b), { dels: [] })))).toBe(
      'opens'
    );
  });

  test('keeps the last counters of each namespace, in namespace order', async () => {
    const counters: [string, number][] = [
      ['a', 10 ** 15 - 1],
      ['b', 1],
    ];
    const bytes = await changed((_, body) => ((body.document as Json).lastCounters = counters));
    const { document: doc } = openOrderedCheckpoint(bytes, { file: 2 });
    expect([...doc.lastCounters]).toEqual(counters);
  });

  test('refuses a prototype key anywhere', async () => {
    for (const key of ['__proto__', 'constructor', 'prototype']) {
      // A record, and a property's attributes, whose reader takes any other key
      for (const place of ['"document":{', '"attributes":{']) {
        const edit = (text: string) => text.replace(place, `${place}"${key}":"",`);
        const bytes = await changed(() => undefined, false, edit);
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
