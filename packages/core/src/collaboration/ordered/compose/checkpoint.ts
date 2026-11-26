// Writes a room checkpoint (design D14): the item list as data (version 6).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { writeZip } from '../../../store/package/zip.ts';
import { CHECKPOINT_MANIFEST, checkpointManifestBytes } from '../model/checkpoint.ts';
import { fileMediaAt, partReferences, referencedMediaType } from '../model/file-media.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedFileMedia, OrderedMediaReference } from '../model/resources.ts';
import { SNAPSHOT_BASE_ENTRY, SNAPSHOT_ENTRY, snapshotJsonBytes } from '../model/snapshot-write.ts';
import type { OrderedDocument } from '../model/types.ts';
import { checkpointFacts } from './checkpoint-facts.ts';
import { withFileMediaPlaces } from './file-media.ts';
import type { RoomNumber } from './context.ts';
import {
  composedWithoutMedia,
  orderedMediaOf,
  writeComposedEntries,
  type OrderedPackageWriteOptions,
} from './package.ts';

/** A media part of a package, as a checkpoint names it, with its bytes. @public */
export interface OrderedPackageMedia extends OrderedMediaReference {
  readonly bytes: Uint8Array;
}

/**
 * The media parts of a package that a checkpoint names by reference, in the package's order:
 * each part of a type the room stores (`OrderedMediaType`) that the package holds as bytes, with
 * at least one byte. Every other part stays inside a checkpoint: fonts, the thumbnail (and any
 * part the package root relates), an ink part (the package reader reads it as XML), and a type
 * the room does not store. The bytes are not checked: a file's own pictures are kept as they
 * are. A host stores each one under its SHA-256 once, when a room is made from the Word file,
 * before it writes the first checkpoint. A package opened from a checkpoint holds none of them.
 *
 * @public
 */
export function orderedPackageMedia(pkg: OoxmlPackage): OrderedPackageMedia[] {
  const referenceOf = partReferences(pkg);
  const out: OrderedPackageMedia[] = [];
  for (const [part, bytes] of pkg.partBytes) {
    const reference = referenceOf(part, bytes);
    if (reference !== null) {
      out.push({ ...reference, bytes });
    }
  }

  return out;
}

/** How {@link writeOrderedCheckpoint} writes a checkpoint. @public */
export interface OrderedCheckpointWriteOptions extends OrderedPackageWriteOptions {
  /**
   * The room's last order number, which the checkpoint holds: the room's log after it starts at
   * 1 again. Default 0, as for a room's first file, which no change came before.
   */
  readonly order?: number;
}

/**
 * The whole document as a room checkpoint, handed to `options.write` front to back: one zip that
 * holds the room's item list as data (`ordered/checkpoint.json`, every stored field of the
 * document) and the Word package it was first opened from (`ordered/base.docx`), whose parts the
 * list's Word XML names by node id. The base package holds no bytes of its media of a type the
 * room stores: `checkpoint.json` names each of them by reference (part name, SHA-256, content
 * type, size, place), and the room's storage holds the bytes, as it holds those of media a change
 * added. So the file carries no media bytes, and `options.media` is not read. Word does not open
 * a checkpoint; {@link openOrderedCheckpoint} does, to the same document. A room's close writes
 * one with {@link closeOrderedRoom}, which checks it too.
 *
 * @throws OrderedRefusal `too-many-items` when `checkpoint.json` is past the engine's limit of
 *   one part (64 MiB), which no open reads. The bytes written before a refusal are not a whole
 *   file: the host keeps its last checkpoint and the room's log.
 * @throws RangeError when `options.order` is not a whole number from 0.
 * @public
 */
export async function writeOrderedCheckpoint(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedCheckpointWriteOptions
): Promise<void> {
  await options.write(orderedCheckpointBytes(doc, base, options.order ?? 0));
}

/** The bytes of `writeOrderedCheckpoint`, at once: a checkpoint holds no media bytes. */
export function orderedCheckpointBytes(
  doc: OrderedDocument,
  base: OoxmlPackage,
  order: number
): Uint8Array {
  if (!Number.isSafeInteger(order) || order < 0) {
    throw new RangeError('A checkpoint needs an order number from 0.');
  }

  const { bytes, media } = baseFileBytes(doc, base);
  return writeZip(
    new Map([
      [SNAPSHOT_ENTRY, snapshotJsonBytes(doc, base, media, order)],
      [SNAPSHOT_BASE_ENTRY, bytes],
    ])
  );
}

/**
 * `ordered/base.docx`: the base package's entries as the bytes it was read from (the open of a room
 * file normalizes them again to the same trees), each media part of a type the room stores by
 * reference, at its place among the entries: the file's own that the package holds no bytes for
 * (from a checkpoint), and those whose bytes it holds (from a Word file), which the host stored
 * when it made the room.
 */
function baseFileBytes(
  doc: OrderedDocument,
  base: OoxmlPackage
): { readonly bytes: Uint8Array; readonly media: readonly OrderedFileMedia[] } {
  const pkg = withFileMediaPlaces(base, doc);
  const referenceOf = partReferences(pkg);
  const media: OrderedFileMedia[] = [];
  const entries = new Map<string, Uint8Array>();
  let entry = 0;
  for (const [name, data] of pkg.partBytes) {
    const reference = fileMediaAt(doc, name) ?? referenceOf(name, data);
    if (reference === null) {
      entries.set(name, data);
    } else {
      const { part, contentType, sha256, size } = reference;
      media.push({ part, contentType, sha256, size, entry });
    }

    entry += 1;
  }

  return { bytes: writeZip(entries), media };
}

/**
 * The whole document as a room checkpoint, handed to `options.write` front to back: the parts
 * of the Word file that {@link writeOrderedPackage} writes for the same document (with its
 * comments: a checkpoint is the room's own file), in the same order, without the media parts
 * that {@link orderedPackageMedia} lists (the file's own, and those a change added), and a
 * manifest part that names each of them by reference (part name, SHA-256, content type, size),
 * with the room facts Word has no place for: each pending suggestion whole, with its person and
 * the revision `w:id`s of its marks, the person of each room comment, the records a suggestion
 * took from another, the person of each namespace that typed units of an insertion, and the
 * records Word writes no change record for. Its parts carry what the manifest names them by: a
 * label on each element a fact names, and the typer on each run a room change typed inside a
 * pending insertion. The content types and the relationships stay as they are. The
 * room's storage holds the media bytes: `options.media` is read only for media a change added
 * that stays inside (an ink part), never for a picture. Word does not open a checkpoint;
 * {@link openOrderedCheckpoint} does. Version 6 replaced it ({@link writeOrderedCheckpoint}).
 *
 * @throws OrderedRefusal as `writeOrderedPackage` does, `too-many-items` when the manifest is
 *   past the engine's limit of one part (64 MiB), which no open reads, and `engine-refused` when
 *   the manifest cannot keep a record made over a pending one. The bytes written before a
 *   refusal are not a whole file: the host keeps its last checkpoint and the room's log.
 * @internal Kept until every room has closed once under version 6.
 */
export async function writeOrderedCheckpointVersion5(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedPackageWriteOptions
): Promise<void> {
  const typers = new Map<string, OrderedPerson>();
  const runs = new Map<string, RoomNumber>();
  const composed = composedWithoutMedia(doc, base, { typers, runs });
  const { media } = composed;
  const { pkg, facts } = checkpointFacts(doc, composed.pkg, composed.takeovers, typers, runs);
  const referenceOf = partReferences(pkg);
  const references: OrderedFileMedia[] = [];
  await writeComposedEntries(
    pkg,
    async (name, data, entry) => {
      const held = media.get(partNameKey(name));
      const reference = held === undefined ? referenceOf(name, data) : heldReference(held);
      if (reference !== null) {
        references.push({ ...reference, part: name, entry });
        return null;
      }

      return held === undefined ? data : orderedMediaOf(held, options.media);
    },
    options.write,
    () => [CHECKPOINT_MANIFEST, checkpointManifestBytes(references, facts)]
  );
}

/** Media the composed document holds by hash, or null when a checkpoint keeps it (an ink part). */
function heldReference(media: OrderedMediaReference): OrderedMediaReference | null {
  return referencedMediaType(media.contentType) === null ? null : media;
}
