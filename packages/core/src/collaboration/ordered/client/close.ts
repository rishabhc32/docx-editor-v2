// A room's close (design D14): the engine's part of it. The host owns the storage: it uploads the
// file this gives, checks the upload, points the room to the new file, and only then deletes the
// old file and the log.

import type { OoxmlPackage, OoxmlPackageLimits } from '../../../store/package/ooxml-package.ts';
import { compactOrdered, type OrderedCompaction } from '../change/compact.ts';
import { orderedCheckpointBytes } from '../compose/checkpoint.ts';
import { writeOrderedPackage, type OrderedPackageWriteOptions } from '../compose/package.ts';
import { snapshotDifference } from '../model/snapshot-compare.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import type { OrderedRoom } from '../room/room.ts';
import { openOrderedCheckpoint, OrderedWordFileError, type OrderedWordFile } from './open.ts';

/** How {@link closeOrderedRoom} reads media and hands the backup on. @public */
export interface OrderedRoomCloseOptions {
  /** The bytes of media by SHA-256, for the backup, as {@link writeOrderedPackage} reads them. */
  readonly media: OrderedPackageWriteOptions['media'];
  /** Takes the next bytes of the backup. The close waits for it before it goes on. */
  readonly backup: OrderedPackageWriteOptions['write'];
  /** Caps on the read-back. A host gives the caps its next open uses. */
  readonly limits?: OoxmlPackageLimits;
}

/** What a room's close gives its host. @public */
export interface OrderedRoomClose {
  /**
   * The room's next file: a checkpoint of version 6, cleaned (`compactOrdered`), which opens to
   * the room's document field by field.
   */
  readonly bytes: Uint8Array;
  /** The SHA-256 of `bytes`, lower-case hex: the host checks its upload against it. */
  readonly sha256: string;
  /** The room's last order number, which the file holds. */
  readonly order: number;
  /** What the cleanup removed. */
  readonly compaction: OrderedCompaction;
}

/**
 * Why a room's close gave no file: the file it wrote does not open to the room. Stable.
 *
 * - `unreadable`: the open refused the file; `detail` is the open's reason.
 * - `different`: the file opens to another document; `detail` names the first field that differs.
 *
 * @public
 */
export type OrderedRoomCloseCode = 'unreadable' | 'different';

/**
 * A close whose read-back failed. The room is as it was: the host keeps the old file and the
 * log, reports the error, and tries again at the next close.
 *
 * @public
 */
export class OrderedRoomCloseError extends Error {
  /** Why the file does not stand for the room. */
  readonly code: OrderedRoomCloseCode;
  /** The open's reason, or the path of the first field that differs. */
  readonly detail: string;

  constructor(code: OrderedRoomCloseCode, detail: string) {
    super(`The room's next file does not open to the room: ${code} ${detail}`);
    this.name = 'OrderedRoomCloseError';
    this.code = code;
    this.detail = detail;
  }
}

/**
 * The engine's part of a room's close, run when no client is connected:
 *
 * 1. It cleans the room's document (`compactOrdered`).
 * 2. It writes the next file: a checkpoint of version 6 with the room's last order number.
 * 3. It opens those bytes with the open the next room runs (`openOrderedCheckpoint`) and compares
 *    the result with the cleaned document, field by field.
 * 4. It writes the backup through `options.backup`: the opened document as a Word file
 *    ({@link writeOrderedPackage}), which still holds the document when a later open refuses the
 *    room's file.
 *
 * Steps 1 to 3 run in one go, and then the cleanup goes back: the room is never seen cleaned, and
 * it is as it was after any outcome. The host opens the next room from the file. It deletes the
 * old file and the log only when the upload of `bytes` gives `sha256` and the room points to the
 * new file. The old file and the log stay the truth until then: a stop at any step loses nothing.
 *
 * @throws OrderedRoomCloseError when the written file does not open to the room.
 * @throws OrderedRefusal as `writeOrderedCheckpoint` and `writeOrderedPackage` do.
 * @public
 */
export async function closeOrderedRoom(
  room: OrderedRoom,
  base: OoxmlPackage,
  options: OrderedRoomCloseOptions
): Promise<OrderedRoomClose> {
  const doc = room.document;
  const { order } = room;
  const { value, record } = withUndoRecord(doc, () => {
    const compaction = compactOrdered(doc);
    const bytes = orderedCheckpointBytes(doc, base, order);
    const opened = readBack(bytes, doc.file, options.limits);
    const difference = opened.order === order ? snapshotDifference(doc, opened.document) : 'order';
    if (difference !== null) {
      throw new OrderedRoomCloseError('different', difference);
    }

    return { compaction, bytes, opened };
  });
  takeBack(doc, record);
  const { compaction, bytes, opened } = value;
  await writeOrderedPackage(opened.document, opened.base, {
    media: options.media,
    write: options.backup,
  });
  return { bytes, sha256: orderedBlobHash(bytes), order, compaction };
}

/** The file as the next room opens it. */
function readBack(
  bytes: Uint8Array,
  file: number,
  limits: OoxmlPackageLimits | undefined
): OrderedWordFile {
  try {
    return openOrderedCheckpoint(bytes, { file, ...(limits === undefined ? {} : { limits }) });
  } catch (error) {
    if (error instanceof OrderedWordFileError) {
      throw new OrderedRoomCloseError('unreadable', error.reason);
    }

    throw error;
  }
}
