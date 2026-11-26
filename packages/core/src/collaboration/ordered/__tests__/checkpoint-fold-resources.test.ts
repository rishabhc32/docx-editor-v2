// What a change adds outside the story text, across a fold (design D14): a list level, a link's
// address, an embedded object, a chart's part and a list's picture bullet. A checkpoint writes
// each one as the file's, so the open gives each back to its suggestion: a decision after a fold
// writes it or leaves it out as the room does live. The manifest is untrusted: an entry that
// does not fit the file refuses the whole checkpoint.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { writeOrderedCheckpointVersion5 } from '../compose/checkpoint.ts';
import type { CheckpointRoom } from './checkpoint-fixture.ts';
import {
  CHART_ROOM,
  LINK_ROOM,
  LIST_LEVEL_ROOM,
  OBJECT_ROOM,
  RESOURCE_ROOMS,
  resourceBlobs,
} from './checkpoint-resources-fixture.ts';
import { checkFold, checkpointBytes } from './fold-check.ts';

/** A room's checkpoint of version 5. */
async function checkpointOf(room: CheckpointRoom): Promise<Uint8Array> {
  const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
  room.edit(doc);
  return checkpointBytes(doc, base, resourceBlobs(base), writeOrderedCheckpointVersion5);
}

/** A room's checkpoint with its manifest's text edited. */
async function crafted(room: CheckpointRoom, edit: (json: string) => string): Promise<Uint8Array> {
  const entries = unzipSync(await checkpointOf(room));
  entries['ordered-checkpoint.json'] = strToU8(
    edit(strFromU8(entries['ordered-checkpoint.json']!))
  );
  return zipSync(entries);
}

/** A room's checkpoint with its manifest's `resources` edited. */
const withResources = (
  room: CheckpointRoom,
  edit: (resources: Record<string, unknown>[]) => unknown
): Promise<Uint8Array> =>
  crafted(room, (json) => {
    const manifest = JSON.parse(json) as { resources: Record<string, unknown>[] };
    return JSON.stringify({ ...manifest, resources: edit(manifest.resources) });
  });

/** The reason a checkpoint does not open, or `opens`. */
function openReason(bytes: Uint8Array): string {
  try {
    openOrderedCheckpoint(bytes, { file: 2 });
    return 'opens';
  } catch (error) {
    return error instanceof OrderedWordFileError ? error.reason : String(error);
  }
}

describe('a fold keeps what a change added outside the story text with its suggestion', () => {
  for (const room of RESOURCE_ROOMS) {
    test(room.name, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      await checkFold(doc, base, resourceBlobs(base), room.name);
    });
  }
});

describe("a version 5 checkpoint's resources are untrusted", () => {
  // An earlier writer listed no resources, and no list with picture bullets
  test('a checkpoint written before the section opens them as the file holds them', async () => {
    for (const room of RESOURCE_ROOMS) {
      const bytes = await crafted(room, (json) => {
        const { resources: _none, ...rest } = JSON.parse(json) as Record<string, unknown>;
        return JSON.stringify({ ...rest, lists: [] });
      });
      expect({ room: room.name, reason: openReason(bytes) }).toEqual({
        room: room.name,
        reason: 'opens',
      });
    }
  });

  test('each forged entry refuses the whole checkpoint', async () => {
    const forged: [string, Promise<Uint8Array>][] = [
      [
        'a part the file does not hold',
        withResources(CHART_ROOM, (list) =>
          list.map((r) => (r.kind === 'part' ? { ...r, key: 'p:ana:9.9' } : r))
        ),
      ],
      [
        'an owner that steps out of the package',
        withResources(LINK_ROOM, (list) =>
          list.map((r) => ({ ...r, owner: '/word/../word/document.xml' }))
        ),
      ],
      ['an entry named twice', withResources(LINK_ROOM, (list) => [...list, ...list])],
      [
        'one key under two suggestions',
        withResources(LINK_ROOM, (list) => [
          ...list,
          ...list.map((r) => ({ ...r, suggestion: 1 })),
        ]),
      ],
      [
        'a level the engine declares otherwise',
        withResources(LIST_LEVEL_ROOM, (list) => list.map((r) => ({ ...r, format: 'bullet' }))),
      ],
      [
        'media with no reference',
        withResources(OBJECT_ROOM, (list) =>
          list.map((r) => (r.kind === 'media' ? { ...r, key: 'p:ana:9.9' } : r))
        ),
      ],
      [
        'a relationship to a part no entry lists',
        withResources(OBJECT_ROOM, (list) => list.filter((r) => r.kind !== 'media')),
      ],
      [
        'a key that reaches a prototype',
        crafted(LINK_ROOM, (json) =>
          json.replace('"kind":"relationship"', '"__proto__":{},"kind":"relationship"')
        ),
      ],
    ];
    for (const [name, bytes] of forged) {
      expect({ name, reason: openReason(await bytes) }).toEqual({
        name,
        reason: 'checkpoint-manifest',
      });
    }
  });
});
