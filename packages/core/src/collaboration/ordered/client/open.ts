import {
  readOoxmlPackage,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
  type OoxmlPackageLimits,
} from '../../../store/package/ooxml-package.ts';
import { normalizeParagraphIdentity } from '../../../store/package/para-id.ts';
import { textboxFallbackExportPackage } from '../../../store/package/textbox-fallback-export.ts';
import { materializeGlossaryPlaceholders } from '../../../store/store/placeholder-materialize.ts';
import { loadOrderedPackage, orderedStoryParts, type OrderedLoadOptions } from '../load/load.ts';
import { holdsCheckpointManifest } from '../model/checkpoint.ts';
import type { OrderedDocument } from '../model/types.ts';
import { holdsSnapshot, openSnapshot } from './open-snapshot.ts';

/** How to open a room's Word file. @public */
export interface OrderedWordFileOptions {
  /** The room's Word file number: the room and every client open one file with one number. */
  readonly file: number;
  /** Caps on the package read. A caller may lower the engine's defaults. */
  readonly limits?: OoxmlPackageLimits;
  /** Caps on the load. */
  readonly load?: Omit<
    OrderedLoadOptions,
    'file' | 'placeholders' | 'suggestions' | 'commentPersons'
  >;
}

/**
 * An opened Word file or room checkpoint: the normalized package and the ordered document loaded
 * from it.
 *
 * @public
 */
export interface OrderedWordFile {
  /**
   * The package the document was loaded from. Compose and save keep its other parts. From a
   * checkpoint, it holds no bytes for the media the checkpoint names by reference: the
   * document's `resources.fileMedia` lists them.
   */
  readonly base: OoxmlPackage;
  readonly document: OrderedDocument;
  /**
   * The room's last order number a checkpoint holds: the room's log after it starts at 1. None
   * for a Word file.
   */
  readonly order?: number;
}

/** Why a Word file or a room checkpoint did not open. @public */
export class OrderedWordFileError extends Error {
  /**
   * The package reader's reason (`zip-limit`, `missing-part`, ...) or `no-main-document`. A Word
   * file that is a checkpoint: `checkpoint`. A checkpoint: `not-checkpoint` (it holds no item
   * list), `checkpoint-format` or `checkpoint-version` (it names another format, or a version this
   * reader does not read: every open refuses a room file of version 5 or older, a Word file with
   * the manifest beside it, by this one), or `checkpoint-manifest` (it does not read, or does not
   * fit the file).
   */
  readonly reason: string;

  constructor(reason: string) {
    super(`The room's Word file cannot be opened: ${reason}`);
    this.name = 'OrderedWordFileError';
    this.reason = reason;
  }
}

/**
 * Open a room's Word file as the room server and every client must, so every copy holds the same
 * items under the same addresses (design D8): the engine's bounded read (the trust boundary), the
 * editor's normalization when it opens a document (a valid `w14:paraId` on every paragraph, an
 * empty control's glossary prompt) and when it saves one (a text box's VML fallback copy in step
 * with its DrawingML copy), and the ordered load under the file's number. A room checkpoint does
 * not open as a Word file: {@link openOrderedCheckpoint} opens it.
 *
 * @throws OrderedWordFileError when the package does not read or has no main document, when it is
 *   a checkpoint (`checkpoint`), or a room file of version 5 or older (`checkpoint-version`).
 * @throws OrderedRefusal when a story is past the load limits.
 * @public
 */
export function openOrderedWordFile(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  if (holdsSnapshot(bytes)) {
    throw new OrderedWordFileError('checkpoint');
  }

  return openWordPackage(readPackage(bytes, options), options);
}

/**
 * Open a room checkpoint (`writeOrderedCheckpoint`): the document the room saved, field by field,
 * with every address as saved. Its `file` stays the number its loaded spaces hold; the room's own
 * number of the file (`options.file`) moves at each close.
 * No media bytes are read. The document's `resources.fileMedia` lists the media the checkpoint
 * names by reference; a client reads each one through its `media` reader, and
 * `composeOrderedPackage` and `writeOrderedPackage` put the bytes back in their places.
 *
 * @throws OrderedWordFileError when the package does not read, or when it is not a checkpoint of
 *   this engine's format and version (`not-checkpoint`, `checkpoint-format`, `checkpoint-version`,
 *   `checkpoint-manifest`).
 * @throws OrderedRefusal when a story is past the load limits.
 * @public
 */
export function openOrderedCheckpoint(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  if (holdsSnapshot(bytes)) {
    return openSnapshot(bytes, options);
  }

  // Read for its refusals first: a room file of version 5 or older is refused by name
  readPackage(bytes, options);
  throw new OrderedWordFileError('not-checkpoint');
}

/** A room's file as a client reads it: a checkpoint by its item list, else a Word file. */
export function openOrderedRoomFile(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  return holdsSnapshot(bytes)
    ? openSnapshot(bytes, options)
    : openWordPackage(readPackage(bytes, options), options);
}

/**
 * The engine's bounded read of a room's file that holds no item list. A room file of version 5 or
 * older, the Word file with the room's facts beside it in a manifest, does not open: no engine
 * reads it now (design D14).
 */
function readPackage(bytes: Uint8Array, options: OrderedWordFileOptions): OoxmlPackage {
  const read = readOoxmlPackage(bytes, options.limits ?? {});
  if (!read.ok) {
    throw new OrderedWordFileError(read.reason);
  }

  if (holdsCheckpointManifest(read.package)) {
    throw new OrderedWordFileError('checkpoint-version');
  }

  return read.package;
}

/**
 * A read Word file, opened from its normalized package as written and read back. A checkpoint keeps
 * these bytes as its base, so every later open reads the trees the room loaded, with the same node
 * ids, and finds nothing left to normalize.
 */
function openWordPackage(pkg: OoxmlPackage, options: OrderedWordFileOptions): OrderedWordFile {
  const written = readOoxmlPackage(
    writeOoxmlPackage(normalizedRoomPackage(pkg)),
    options.limits ?? {}
  );
  if (!written.ok) {
    throw new OrderedWordFileError(written.reason);
  }

  const base = normalizedRoomPackage(written.package);
  return { base, document: loadOrderedPackage(base, { ...options.load, file: options.file }) };
}

/**
 * A read package as every open of a room's file normalizes it, the same bytes to the same trees:
 * the editor's normalization when it opens a document (a valid `w14:paraId` on every paragraph, an
 * empty control's glossary prompt) and when it saves one (a text box's VML fallback copy in step
 * with its DrawingML copy).
 *
 * @throws OrderedWordFileError `no-main-document`.
 */
export function normalizedRoomPackage(pkg: OoxmlPackage): OoxmlPackage {
  let base = pkg;
  const main = base.parts.get(base.mainDocumentPart);
  if (!main) throw new OrderedWordFileError('no-main-document');
  base = withPart(base, materializeGlossaryPlaceholders(base, normalizeParagraphIdentity(main)));
  for (const [name, part] of orderedStoryParts(base))
    if (name !== base.mainDocumentPart) base = withPart(base, normalizeParagraphIdentity(part));

  // As a save writes them, so the file the room saves reopens to the same paragraph ids
  return textboxFallbackExportPackage(base);
}
