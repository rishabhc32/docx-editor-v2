import {
  readOoxmlPackage,
  withPart,
  type OoxmlPackage,
  type OoxmlPackageLimits,
} from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { normalizeParagraphIdentity } from '../../../store/package/para-id.ts';
import { textboxFallbackExportPackage } from '../../../store/package/textbox-fallback-export.ts';
import { materializeGlossaryPlaceholders } from '../../../store/store/placeholder-materialize.ts';
import { loadOrderedPackage, orderedStoryParts, type OrderedLoadOptions } from '../load/load.ts';
import {
  holdsCheckpointManifest,
  readCheckpoint,
  type CheckpointFacts,
} from '../model/checkpoint.ts';
import { checkpointPlacesOf } from '../model/checkpoint-takeovers.ts';
import { recordFileMedia } from '../model/file-media.ts';
import { orderedAuthorName, type OrderedPerson } from '../model/person.ts';
import type { OrderedFileMedia } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withCheckpointTakeovers } from './takeovers.ts';
import { withCheckpointTypers } from './typers.ts';
import { withCheckpointUntracked } from './untracked.ts';

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
}

/** Why a Word file or a room checkpoint did not open. @public */
export class OrderedWordFileError extends Error {
  /**
   * The package reader's reason (`zip-limit`, `missing-part`, ...) or `no-main-document`. A Word
   * file that holds a checkpoint's manifest: `checkpoint`. A checkpoint: `not-checkpoint` (it
   * holds no manifest), `checkpoint-format` or `checkpoint-version` (the manifest names another
   * format or version of it), or `checkpoint-manifest` (the manifest does not read, or does not
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
 * @throws OrderedWordFileError when the package does not read, has no main document, or holds
 *   a checkpoint's manifest (`checkpoint`).
 * @throws OrderedRefusal when a story is past the load limits.
 * @public
 */
export function openOrderedWordFile(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  const pkg = readPackage(bytes, options);
  if (holdsCheckpointManifest(pkg)) {
    throw new OrderedWordFileError('checkpoint');
  }

  return openRead(pkg, [], options);
}

/**
 * Open a room checkpoint (`writeOrderedCheckpoint`) as {@link openOrderedWordFile} opens the
 * Word file it stands for: the same document, under the same addresses. No media bytes are
 * read. The document's `resources.fileMedia` lists the media the checkpoint names by reference;
 * a client reads each one through its `media` reader, and `composeOrderedPackage` and
 * `writeOrderedPackage` put the bytes back in their places. Whatever the file number it opens
 * under, each pending suggestion is one suggestion again, with its person, however many Word
 * changes its marks are; each record a suggestion took from another (policy 2) is that record
 * again, with `replaced`; each record Word writes no change record for (a control's value, a
 * column's lent width) is that record again, over the properties it was made over; and
 * `orderedUnitTyper` names who typed each unit of an insertion as the room did.
 *
 * @throws OrderedWordFileError when the package does not read or has no main document, or when
 *   it is not a checkpoint of this engine's format and version (`not-checkpoint`,
 *   `checkpoint-format`, `checkpoint-version`, `checkpoint-manifest`).
 * @throws OrderedRefusal when a story is past the load limits.
 * @public
 */
export function openOrderedCheckpoint(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  return openCheckpointRead(readPackage(bytes, options), options);
}

/**
 * A room's file as a client reads it: a checkpoint when it holds the manifest, else a Word file.
 */
export function openOrderedRoomFile(
  bytes: Uint8Array,
  options: OrderedWordFileOptions
): OrderedWordFile {
  const pkg = readPackage(bytes, options);
  return holdsCheckpointManifest(pkg)
    ? openCheckpointRead(pkg, options)
    : openRead(pkg, [], options);
}

/** The engine's bounded read of a room's file. */
function readPackage(bytes: Uint8Array, options: OrderedWordFileOptions): OoxmlPackage {
  const read = readOoxmlPackage(bytes, options.limits ?? {});
  if (!read.ok) throw new OrderedWordFileError(read.reason);
  return read.package;
}

/**
 * A read checkpoint package, opened without its manifest and with the media and the room facts it
 * names: each by the suggestions the manifest lists, whose ids the load keeps in its order.
 */
function openCheckpointRead(pkg: OoxmlPackage, options: OrderedWordFileOptions): OrderedWordFile {
  const checkpoint = readCheckpoint(pkg);
  if (!checkpoint.ok) {
    throw new OrderedWordFileError(checkpoint.reason);
  }

  const { facts } = checkpoint;
  const opened = openRead(checkpoint.base, checkpoint.media, options, facts);
  const doc = opened.document;
  const listedIds = [...doc.revisionIds.listed.keys()];

  // Only the takeovers and the untracked records name places of the file
  const places =
    facts.takeovers.length + facts.untracked.length === 0 ? new Map() : placesOf(opened.base);
  if (
    !withCheckpointTakeovers(doc, places, facts.takeovers, listedIds) ||
    !withCheckpointUntracked(doc, places, facts.untracked, listedIds) ||
    !withCheckpointTypers(doc, facts.typers) ||
    !personsFit(doc)
  ) {
    throw new OrderedWordFileError('checkpoint-manifest');
  }

  return opened;
}

/** The places of a package's stories the manifest names records by, each by its key. */
function placesOf(pkg: OoxmlPackage): Map<string, OoxmlElement> {
  const places = new Map<string, OoxmlElement>();
  for (const [name, part] of orderedStoryParts(pkg)) {
    for (const { element, key } of checkpointPlacesOf(name, part)) {
      places.set(key, element);
    }
  }

  return places;
}

/** Whether every item and comment the manifest gives a person carries that person's Word name. */
function personsFit(doc: OrderedDocument): boolean {
  const fits = ({ author, person }: { author: string; person: OrderedPerson | null }) =>
    person === null || author === orderedAuthorName(person);
  for (const { attribution } of doc.suggestions.values()) {
    if (!fits(attribution)) {
      return false;
    }
  }

  for (const comment of doc.comments.values()) {
    if (!fits(comment)) {
      return false;
    }
  }

  return true;
}

/**
 * A read package, normalized and loaded, with the file's media, and the suggestions and comment
 * persons a checkpoint names.
 */
function openRead(
  pkg: OoxmlPackage,
  fileMedia: readonly OrderedFileMedia[],
  options: OrderedWordFileOptions,
  facts?: CheckpointFacts
): OrderedWordFile {
  let base = pkg;
  const main = base.parts.get(base.mainDocumentPart);
  if (!main) throw new OrderedWordFileError('no-main-document');
  base = withPart(base, materializeGlossaryPlaceholders(base, normalizeParagraphIdentity(main)));
  for (const [name, part] of orderedStoryParts(base))
    if (name !== base.mainDocumentPart) base = withPart(base, normalizeParagraphIdentity(part));

  // As a save writes them, so the file the room saves reopens to the same paragraph ids
  base = textboxFallbackExportPackage(base);
  const document = loadOrderedPackage(base, {
    ...options.load,
    file: options.file,
    ...(facts === undefined
      ? {}
      : { suggestions: facts.suggestions, commentPersons: facts.commentPersons }),
  });
  recordFileMedia(document, fileMedia);
  return { base, document };
}
