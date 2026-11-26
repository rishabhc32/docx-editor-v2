// A checkpoint of version 6 saves the room's item list (design D2 to D6): an open gives back every
// stored field, the same Word file and the same bytes on the next save.

import { describe, expect, test } from 'bun:test';
import { openOrderedRoomFile, openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
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
import { checkpointBytes, folded, wordMeaning } from './fold-check.ts';

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

describe('a checkpoint of version 6', () => {
  for (const room of ROOMS) {
    test(`${room.name}: an open gives back every stored field and the same file`, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      const blobs = resourceBlobs(base);

      // Under the room's next file number: the document keeps the number of its addresses
      const next = await folded(doc, base, 2, blobs);
      const live = orderedPackageBytes(doc, base, { blobs });
      const opened = orderedPackageBytes(next.document, next.base, { blobs });
      expect(wordMeaning(opened)).toEqual(wordMeaning(live));

      // The first save writes the base as the package writer writes it; from then on, the same
      const again = openOrderedRoomFile(await checkpointBytes(next.document, next.base, blobs), {
        file: 3,
      });
      expect(orderedPackageBytes(again.document, again.base, { blobs })).toEqual(opened);
    });
  }
});
