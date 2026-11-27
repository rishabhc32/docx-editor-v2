// A room's close (design D14): the room cleaned, written as a checkpoint, opened back and compared
// field by field before the host deletes anything. The room stays as it was whatever happens, so
// the old file and the log stay the truth until the host moves the room to the new file. The
// room here is a random session's: its Word file, then the log of each round.

import { afterEach, describe, expect, mock, spyOn, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { applyOrderedChange } from '../change/apply.ts';
import * as compact from '../change/compact.ts';
import type { OrderedStampedChange } from '../change/types.ts';
import {
  closeOrderedRoom,
  OrderedRoomCloseError,
  type OrderedRoomClose,
  type OrderedRoomCloseOptions,
} from '../client/close.ts';
import * as open from '../client/open.ts';
import * as checkpoint from '../compose/checkpoint.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { snapshotDifference } from '../model/snapshot-compare.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { mediaStore, wordMeaning } from './fold-check.ts';
import { SESSION_FILE } from './sessions/fixture.ts';
import { itemViews } from './snapshot-json.ts';
import { runSession, type SessionResult } from './sessions/session.ts';

/** The session's log, and where each of its rounds ends. */
let session: SessionResult | undefined;
const sessionLog = (): SessionResult => (session ??= runSession(5, ['Olga', 'Ana'], 240, 3));

/** The session's Word file, under the file number its log's addresses name. */
const wordFile = () => open.openOrderedWordFile(SESSION_FILE, { file: 0 });

const BLOBS = mediaStore(wordFile().base);

const STOP = new Error('stopped');

/** The changes of one file: the log from `from` to `to`, numbered from 1 again. */
const fileLog = (from: number, to: number): OrderedStampedChange[] =>
  sessionLog()
    .log.slice(from, to)
    .map((change) => ({ ...change, n: change.n - from }));

/** A room on a document, rebuilt from its file's log. */
const roomOn = (document: OrderedDocument, log: readonly OrderedStampedChange[]): OrderedRoom =>
  new OrderedRoom({
    document,
    applyChange: applyOrderedChange,
    policy: suggestionsOnlyPolicy(),
    now: () => 0,
    log,
  });

/** The room of the first round on the Word file, as a host opens it from the file and its log. */
function firstRound() {
  const { document, base } = wordFile();
  return { room: roomOn(document, fileLog(0, sessionLog().rounds[0]!)), base };
}

/** The close, with the backup it wrote. */
async function closed(
  room: OrderedRoom,
  base: Parameters<typeof closeOrderedRoom>[1],
  backup: OrderedRoomCloseOptions['backup'] = async () => undefined
): Promise<OrderedRoomClose & { readonly backup: Uint8Array }> {
  const chunks: Uint8Array[] = [];
  const close = await closeOrderedRoom(room, base, {
    media: async (sha256) => BLOBS(sha256) ?? new Uint8Array(),
    backup: async (chunk) => {
      await backup(chunk);
      chunks.push(chunk);
    },
  });
  return { ...close, backup: new Uint8Array(Buffer.concat(chunks)) };
}

/** `bytes` with one letter of the first item text changed, and its checksum made again if `rehash`. */
function oneLetterOff(bytes: Uint8Array, rehash: boolean): Uint8Array {
  const entries = unzipSync(bytes);
  const json = strFromU8(entries['ordered/checkpoint.json']!);
  const start = json.indexOf('"body":') + '"body":'.length;
  const body = JSON.parse(json.slice(start, -1)) as Record<string, unknown>;
  const item = itemViews(body).find((view) => /^[A-Za-z]/.test(String(view.get('text'))))!;
  const text = item.get('text') as string;
  item.set('text', `${text[0] === 'x' ? 'y' : 'x'}${text.slice(1)}`);
  const changed = JSON.stringify(body);
  const head = rehash
    ? json
        .slice(0, start)
        .replace(/"sha256":"[0-9a-f]{64}"/, `"sha256":"${orderedBlobHash(strToU8(changed))}"`)
    : json.slice(0, start);
  return zipSync({ ...entries, 'ordered/checkpoint.json': strToU8(`${head}${changed}}`) });
}

afterEach(() => {
  mock.restore();
});

describe("a room's close", () => {
  test('gives the cleaned file, its checksum, its order number and a backup', async () => {
    const { room, base } = firstRound();
    const close = await closed(room, base);
    expect(close.sha256).toBe(orderedBlobHash(close.bytes));
    expect(close.order).toBe(room.order);
    expect(close.compaction.items).toBeGreaterThan(0);

    // The room is as the Word file and the log give it: the cleanup went back
    expect(snapshotDifference(firstRound().room.document, room.document)).toBeNull();
    const opened = open.openOrderedCheckpoint(close.bytes, { file: 1 });
    expect(opened.order).toBe(room.order);

    // The backup is the room's Word file: the cleanup changes nothing a person sees
    const live = orderedPackageBytes(room.document, base, { blobs: BLOBS });
    expect(wordMeaning(close.backup)).toEqual(wordMeaning(live));
  });

  test("the room's file and the log after it give the room the whole log gives", async () => {
    const { rounds } = sessionLog();
    const whole = wordFile();
    let file = wordFile();
    let from = 0;
    for (const to of rounds) {
      const room = roomOn(file.document, fileLog(from, to));
      file = open.openOrderedCheckpoint((await closed(room, file.base)).bytes, { file: 1 });

      // The whole log, with the cleanup where each close ran it
      roomOn(whole.document, fileLog(from, to));
      compact.compactOrdered(whole.document);
      from = to;
    }

    expect(snapshotDifference(whole.document, file.document)).toBeNull();

    // With no cleanup at all, the same Word file
    const never = wordFile();
    roomOn(never.document, fileLog(0, from));
    expect(wordMeaning(orderedPackageBytes(file.document, file.base, { blobs: BLOBS }))).toEqual(
      wordMeaning(orderedPackageBytes(never.document, never.base, { blobs: BLOBS }))
    );
  });

  /** Where a close stops, and the backup writer that stops it there, if any. */
  type Stop = () => OrderedRoomCloseOptions['backup'] | undefined;
  const stops: readonly (readonly [string, Stop])[] = [
    [
      'before the cleanup ends',
      () => {
        const cleanup = compact.compactOrdered;
        spyOn(compact, 'compactOrdered').mockImplementation((doc) => {
          cleanup(doc);
          throw STOP;
        });
        return undefined;
      },
    ],
    [
      'after the write',
      () => {
        spyOn(open, 'openOrderedCheckpoint').mockImplementation(() => {
          throw STOP;
        });
        return undefined;
      },
    ],
    [
      'after the read-back',
      () => async () => {
        throw STOP;
      },
    ],
  ];
  for (const [name, stop] of stops) {
    test(`a stop ${name} leaves the room as the old file and the log give it`, async () => {
      const { room, base } = firstRound();
      await expect(closed(room, base, stop())).rejects.toBe(STOP);
      mock.restore();
      expect(snapshotDifference(firstRound().room.document, room.document)).toBeNull();

      // The next close gives the file a close with no stop gives
      const next = firstRound();
      expect((await closed(room, base)).bytes).toEqual((await closed(next.room, next.base)).bytes);
    });
  }

  const faults = [
    ['the checksum does not cover', false, 'unreadable', 'checkpoint-manifest'],
    ['the checksum covers', true, 'different', expect.stringMatching(/^document\.stories/)],
  ] as const;
  for (const [name, rehash, code, detail] of faults) {
    test(`refuses a file whose writer changed one byte ${name}`, async () => {
      const write = checkpoint.orderedCheckpointBytes;
      spyOn(checkpoint, 'orderedCheckpointBytes').mockImplementation((...args) =>
        oneLetterOff(write(...args), rehash)
      );
      const { room, base } = firstRound();
      const refusal = await closed(room, base).catch((error: unknown) => error);
      expect(refusal).toBeInstanceOf(OrderedRoomCloseError);
      expect(refusal).toMatchObject({ code, detail });
      expect(snapshotDifference(firstRound().room.document, room.document)).toBeNull();
    });
  }
});
