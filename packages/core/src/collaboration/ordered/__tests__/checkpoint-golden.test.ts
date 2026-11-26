// Golden checkpoints (design D14): checkpoints this engine wrote when it shipped version 5 of the
// format (`fixtures/checkpoints`), and version 6 (`fixtures/checkpoints-v6`, the item list as
// data), each from a room of `checkpoint-fixture.ts` over an engine test document, with the fold
// print of each one opened. Every later engine must read them to the same print. No engine writes
// version 5 now: its files are never written again. A version 6 file is written when it is
// missing: delete it, run `ORDERED_GOLDEN_WRITE=1 bun test <this file>`, then format the prints.
// A print changes only when what it shows changes (an engine that composes the Word file another
// way), never because a checkpoint reads otherwise.

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

/**
 * The golden files of each version, by folder: version 5's are never written again; the current
 * version's are written when missing, by this engine's writer.
 */
const VERSIONS = [
  { folder: 'checkpoints', written: false },
  { folder: 'checkpoints-v6', written: true },
] as const;

/** A golden file of a room, by the room's name. */
const goldenUrl = (folder: string, room: CheckpointRoom, extension: string): URL =>
  new URL(`./fixtures/${folder}/${room.name.replaceAll(' ', '-')}.${extension}`, import.meta.url);

/** A golden checkpoint opened, and its print; the room's Word file gives the media bytes. */
function printOf(room: CheckpointRoom, bytes: Uint8Array): unknown {
  const blobs = roomBlobs(openOrderedWordFile(room.source, { file: 1 }).base);
  const opened = openOrderedCheckpoint(bytes, { file: 2 });
  return JSON.parse(JSON.stringify(foldPrint(opened.document, opened.base, blobs)));
}

describe('a golden checkpoint', () => {
  for (const { folder, written } of VERSIONS) {
    for (const room of GOLDEN) {
      test.if(WRITE && written)(`${folder} ${room.name}: written when missing`, async () => {
        const file = goldenUrl(folder, room, 'docx');
        if (!existsSync(file)) {
          const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
          room.edit(doc);
          writeFileSync(file, await checkpointBytes(doc, base, roomBlobs(base)));
        }

        const print = printOf(room, new Uint8Array(readFileSync(file)));
        writeFileSync(goldenUrl(folder, room, 'print.json'), JSON.stringify(print));
      });

      test(`${folder} ${room.name}: opens to its print`, () => {
        const bytes = new Uint8Array(readFileSync(goldenUrl(folder, room, 'docx')));
        const json = readFileSync(goldenUrl(folder, room, 'print.json'), 'utf8');
        expect(printOf(room, bytes)).toEqual(JSON.parse(json));
      });
    }
  }
});
