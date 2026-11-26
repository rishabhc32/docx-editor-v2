// A checkpoint of version 6 is untrusted (design D6): a file that is not what the writer wrote is
// refused with its code, whatever field is wrong.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
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
