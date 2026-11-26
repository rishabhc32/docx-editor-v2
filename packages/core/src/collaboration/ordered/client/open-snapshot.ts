// Opens a room checkpoint of version 6 (design D5, D6): the item list as data, and the base
// package its Word XML names. Every value comes from a file.

import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { DEFAULT_ZIP_LIMITS, readZip, zipHoldsEntry } from '../../../store/package/zip.ts';
import { composeStoryPart } from '../compose/compose.ts';
import { recordLoadedPackage } from '../load/load.ts';
import { holdsCheckpointManifest, mediaFits } from '../model/checkpoint.ts';
import { recordFileMedia } from '../model/file-media.ts';
import { readSnapshotJson } from '../model/snapshot-read.ts';
import { SNAPSHOT_BASE_ENTRY, SNAPSHOT_ENTRY } from '../model/snapshot-write.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  normalizedRoomPackage,
  OrderedWordFileError,
  type OrderedWordFile,
  type OrderedWordFileOptions,
} from './open.ts';

/** Whether a room's file is a checkpoint of version 6: its zip holds the item list. */
export const holdsSnapshot = (bytes: Uint8Array): boolean => zipHoldsEntry(bytes, SNAPSHOT_ENTRY);

/**
 * A checkpoint of version 6: its base package, read by the engine's bounded reader, and the
 * document its item list holds, checked field by field and composed once. The document keeps the
 * Word file number its list was loaded under (`file`), so every address stays as it was; the
 * room's own number of the file (`options.file`) moves at each close. The document's
 * `resources.fileMedia` lists the media the base names by reference.
 *
 * @throws OrderedWordFileError as `openOrderedCheckpoint` does.
 */
export function openSnapshot(bytes: Uint8Array, options: OrderedWordFileOptions): OrderedWordFile {
  const zip = readZip(bytes, options.limits?.zip ?? DEFAULT_ZIP_LIMITS);
  if (!zip.ok) {
    throw new OrderedWordFileError(zip.reason);
  }

  const json = zip.entries.get(SNAPSHOT_ENTRY);
  const base = basePackage(zip.entries.get(SNAPSHOT_BASE_ENTRY), options);
  if (json === undefined || base === null || zip.entries.size !== 2) {
    throw new OrderedWordFileError('checkpoint-manifest');
  }

  const read = readSnapshotJson(json, base);
  if (!read.ok) {
    throw new OrderedWordFileError(read.reason);
  }

  const { document, media, order } = read;
  if (!mediaFits(base, media)) {
    throw new OrderedWordFileError('checkpoint-manifest');
  }

  recordFileMedia(document, media);
  recordLoadedPackage(document, base);
  if (!composesEveryStory(document)) {
    throw new OrderedWordFileError('checkpoint-manifest');
  }

  return { base, document, order };
}

/**
 * The base package: a Word package, which holds no checkpoint of its own, normalized as the open
 * of a room's file normalizes it. Its entries are the bytes the base was read from, so these are
 * the trees the room held.
 */
function basePackage(
  bytes: Uint8Array | undefined,
  options: OrderedWordFileOptions
): OoxmlPackage | null {
  const read = bytes === undefined ? null : readOoxmlPackage(bytes, options.limits ?? {});
  if (read?.ok !== true || holdsCheckpointManifest(read.package)) {
    return null;
  }

  return normalizedRoomPackage(read.package).pkg;
}

/**
 * Whether compose draws every story without a refusal. The check leaves the document as it was:
 * compose mints node ids from `composeIds`, which goes back.
 */
function composesEveryStory(doc: OrderedDocument): boolean {
  const next = doc.composeIds.next;
  try {
    for (const name of doc.stories.keys()) {
      // A text box's story is drawn into its object, in its host's story
      if (!isTextBoxStoryName(name)) {
        composeStoryPart(doc, name);
      }
    }

    return true;
  } catch {
    // Any failure: the list came from a file
    return false;
  } finally {
    doc.composeIds.next = next;
  }
}
