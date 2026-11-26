// Golden checkpoints (design D14): checkpoints of version 6 (`fixtures/checkpoints-v6`, the item
// list as data), each from a room of `checkpoint-fixture.ts` over an engine test document, with the
// fold print of each one opened. Every later engine must read them to the same print. A file is
// written when it is missing: delete it, run `ORDERED_GOLDEN_WRITE=1 bun test <this file>`, then
// format the prints. A print changes only when what it shows changes (an engine that composes the
// Word file another way), never because a checkpoint reads otherwise.

import { describe, expect, test } from 'bun:test';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import {
  PASTED_LIST,
  ROOM_FACTS,
  roomBlobs,
  SPLIT_RUN,
  STYLE_RECORDS,
  TABLE_RECORDS,
  TYPED_BARE,
  TYPED_RUNS,
  type CheckpointRoom,
} from './checkpoint-fixture.ts';
import { checkpointBytes, foldPrint } from './fold-check.ts';

const GOLDEN: readonly CheckpointRoom[] = [
  ROOM_FACTS,
  STYLE_RECORDS,
  TYPED_RUNS,
  TYPED_BARE,
  SPLIT_RUN,
  PASTED_LIST,
  TABLE_RECORDS,
];

const WRITE = process.env.ORDERED_GOLDEN_WRITE === '1';

/** A golden file of a room, by the room's name. */
const goldenUrl = (room: CheckpointRoom, extension: string): URL =>
  new URL(
    `./fixtures/checkpoints-v6/${room.name.replaceAll(' ', '-')}.${extension}`,
    import.meta.url
  );

/** A golden checkpoint opened, and its print; the room's Word file gives the media bytes. */
function printOf(room: CheckpointRoom, bytes: Uint8Array): unknown {
  const blobs = roomBlobs(openOrderedWordFile(room.source, { file: 1 }).base);
  const opened = openOrderedCheckpoint(bytes, { file: 2 });
  return JSON.parse(JSON.stringify(foldPrint(opened.document, opened.base, blobs)));
}

describe('a golden checkpoint', () => {
  for (const room of GOLDEN) {
    test.if(WRITE)(`${room.name}: written when missing`, async () => {
      const file = goldenUrl(room, 'docx');
      if (!existsSync(file)) {
        const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
        room.edit(doc);
        writeFileSync(file, await checkpointBytes(doc, base, roomBlobs(base)));
      }

      const print = printOf(room, new Uint8Array(readFileSync(file)));
      writeFileSync(goldenUrl(room, 'print.json'), JSON.stringify(print));
    });

    test(`${room.name}: opens to its print`, () => {
      const bytes = new Uint8Array(readFileSync(goldenUrl(room, 'docx')));
      const json = readFileSync(goldenUrl(room, 'print.json'), 'utf8');
      expect(printOf(room, bytes)).toEqual(JSON.parse(json));
    });
  }
});
