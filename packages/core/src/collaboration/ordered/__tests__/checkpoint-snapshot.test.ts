// A checkpoint of version 6 saves the room's item list (design D2 to D6): an open gives back every
// stored field, the same Word file and the same bytes on the next save.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-serialize.ts';
import { openOrderedRoomFile, openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { snapshotDifference } from '../model/snapshot-compare.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  PASTED_LIST,
  ROOM_FACTS,
  SPLIT_RUN,
  STYLE_RECORDS,
  TABLE_RECORDS,
  TYPED_BARE,
  TYPED_RUNS,
  type CheckpointRoom,
} from './checkpoint-fixture.ts';
import { RESOURCE_ROOMS, resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkpointBytes } from './fold-check.ts';

const ROOMS: readonly CheckpointRoom[] = [
  ROOM_FACTS,
  STYLE_RECORDS,
  TYPED_RUNS,
  TYPED_BARE,
  SPLIT_RUN,
  PASTED_LIST,
  TABLE_RECORDS,
  ...RESOURCE_ROOMS,
];

/**
 * A Word file by what it says: each XML part by the engine's semantic oracle (prefixes and where
 * namespaces are declared are the writer's choice), each other part by its bytes.
 */
function meaning(bytes: Uint8Array): Record<string, string> {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) {
    throw new Error(read.reason);
  }

  const { parts, partBytes } = read.package;
  return Object.fromEntries(
    [...partBytes].map(([name, data]) => {
      const part = parts.get(name);
      return [name, part === undefined ? orderedBlobHash(data) : canonicalOoxmlFingerprint(part)];
    })
  );
}

describe('a checkpoint of version 6', () => {
  for (const room of ROOMS) {
    test(`${room.name}: an open gives back every stored field and the same file`, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      const blobs = resourceBlobs(base);
      const bytes = await checkpointBytes(doc, base, blobs);
      const same = openOrderedRoomFile(bytes, { file: 1 });
      expect(snapshotDifference(doc, same.document)).toBeNull();

      // Before any compose: each one moves the counter of the node ids it mints
      expect(await checkpointBytes(same.document, same.base, blobs)).toEqual(bytes);
      const next = openOrderedRoomFile(bytes, { file: 2 });
      expect(snapshotDifference(doc, next.document)).toBeNull();
      const live = orderedPackageBytes(doc, base, { blobs });
      const opened = orderedPackageBytes(next.document, next.base, { blobs });
      expect(meaning(opened)).toEqual(meaning(live));

      // The first save writes the base as the package writer writes it; from then on, the same
      const again = openOrderedRoomFile(await checkpointBytes(next.document, next.base, blobs), {
        file: 3,
      });
      expect(orderedPackageBytes(again.document, again.base, { blobs })).toEqual(opened);
    });
  }
});
