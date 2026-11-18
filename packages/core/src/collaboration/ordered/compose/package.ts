// Whole-document save (design D3, D10's close step): every story composed into the package the
// document was loaded from.

import type { OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  ooxmlPackageEntriesSteps,
  withPart,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '../../../store/package/ooxml-package.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { zipPartWriter } from '../../../store/package/zip.ts';
import { withRetainedPrompts } from '../../../store/store/content-control-prompt-glossary.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedMediaReference } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withoutUnusedLoadedMedia } from './unused-media.ts';
import { withOrderedComments } from './comments.ts';
import { composeStoryPart } from './compose.ts';
import { withoutDroppedFurniture } from './furniture.ts';
import { withDecidedPartRevisions } from './part-revisions.ts';
import { withUniqueIds } from './unique-ids.ts';
import { MEDIA_PLACE, withFileMediaBytes, withFileMediaPlaces } from './file-media.ts';
import { withoutComments } from './without-comments.ts';
import type { HeldTakeovers, OrderedRetainedPrompt } from './context.ts';
import {
  isAddedStory,
  mediaBytesOf,
  withOrderedResources,
  type OrderedMediaBytes,
  type OrderedResourceOptions,
} from './resources.ts';

/**
 * Whether the Word file holds comments.
 *
 * @public
 */
export interface OrderedPackageCommentOptions {
  /**
   * `keep` (the default) writes the file's comments and the comments the room added.
   * `remove` writes no comment at all: no comment marker in any part, no comments,
   * commentsExtended, commentsIds or commentsExtensible part with their relationships and
   * content types, and no part that only removed parts name, in turn (a picture in a comment,
   * a chart's workbook). Everything else stays as `keep` writes it, tracked changes too. The
   * document does not change.
   */
  readonly comments?: 'keep' | 'remove';
}

/**
 * The whole document as a package: every story part composed, the other parts of `base` kept as
 * the file had them, and what the session added outside the stories written where Word reads
 * it. `base` is the package the document was loaded from, and is not changed. `options.blobs`
 * gives the bytes of added media, and of the file's media a checkpoint names by reference.
 * `options.comments` keeps or removes the comments.
 *
 * @throws OrderedRefusal when the package cannot hold what the session added (a story whose
 *   relationships the package keeps as bytes, a link id the file already uses), when the bytes
 *   of used media are missing (`missing-blob`, unless `options.missing` is `skip`), or when a
 *   comment part cannot leave the package (`engine-refused`).
 * @public
 */
export function composeOrderedPackage(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedResourceOptions & OrderedPackageCommentOptions = {}
): OoxmlPackage {
  return composeWithMedia(doc, base, mediaBytesOf(options), options).pkg;
}

/** The takeovers the change records of each story part hold: by part name, then by node id. */
export type OrderedPartTakeovers = ReadonlyMap<string, ReadonlyMap<string, HeldTakeovers>>;

/**
 * {@link composeOrderedPackage}, with the bytes of added media from `bytes`, and the takeovers
 * the change records of its story parts hold.
 */
function composeWithMedia(
  doc: OrderedDocument,
  base: OoxmlPackage,
  bytes: OrderedMediaBytes,
  options: OrderedPackageCommentOptions
): { readonly pkg: OoxmlPackage; readonly takeovers: OrderedPartTakeovers } {
  const removesComments = options.comments === 'remove';
  let pkg = withFileMediaPlaces(base, doc);
  const prompts: OrderedRetainedPrompt[] = [];
  const added = new Map<string, OoxmlPart>();
  const takeovers = new Map<string, ReadonlyMap<string, HeldTakeovers>>();
  for (const name of doc.stories.keys()) {
    // A text box's story is drawn into its object, in its host's part
    if (isTextBoxStoryName(name)) continue;
    const { part: composedPart, composed, takeovers: held } = composeStoryPart(doc, name);
    takeovers.set(name, held);

    // Each node id once in a part: the package writer finds nodes by id
    const part = withUniqueIds(composedPart);
    prompts.push(...composed.prompts);

    // A story a change added is written only while something uses it
    if (isAddedStory(doc, name)) {
      added.set(name, part);
      continue;
    }

    pkg = withPart(pkg, part);
  }

  // The prompts a suggestion hid, kept as placeholder blocks in the glossary
  pkg = withRetainedPrompts(pkg, prompts, new Set());
  pkg = withoutUnusedLoadedMedia(pkg, doc);

  // The file's tracked changes outside the stories as the decisions leave them, before the
  // resources: the styles they add go after the file's
  pkg = withDecidedPartRevisions(pkg, doc);

  // The comment parts as the model gives them, then every resource a change added that shows and
  // is used (media, new story parts, relationships, list definitions, styles and settings)
  const commented = removesComments ? pkg : withOrderedComments(pkg, doc);
  pkg = withoutDroppedFurniture(withOrderedResources(commented, doc, added, bytes), doc);

  // No comment, before the media bytes: a picture only a comment showed is not read
  if (removesComments) {
    pkg = withoutComments(pkg);
  }

  // Last, the bytes of the file's media a checkpoint names that the document still uses
  return { pkg: withFileMediaBytes(pkg, doc, bytes), takeovers };
}

/**
 * The whole document as `.docx` bytes, through the engine's package writer.
 *
 * @public
 */
export function orderedPackageBytes(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedResourceOptions & OrderedPackageCommentOptions = {}
): Uint8Array {
  return writeOoxmlPackage(composeOrderedPackage(doc, base, options));
}

/**
 * How {@link writeOrderedPackage} and `writeOrderedCheckpoint` read media and hand the file on.
 *
 * @public
 */
export interface OrderedPackageWriteOptions {
  /**
   * The bytes of media by SHA-256: media a change added, and the file's media a checkpoint names
   * by reference. A rejection, or bytes of another size, refuses the file (`missing-blob`).
   */
  readonly media: (sha256: string) => Promise<Uint8Array>;
  /** Takes the next bytes of the file. The writer waits for it before it goes on. */
  readonly write: (chunk: Uint8Array) => Promise<void>;
}

/** A package composed without the bytes of media, and that media by part name key. */
interface OrderedComposedWithoutMedia {
  readonly pkg: OoxmlPackage;
  readonly media: ReadonlyMap<string, OrderedMediaReference>;
  readonly takeovers: OrderedPartTakeovers;
}

/**
 * The document composed with each media part it uses empty, for a writer that reads the bytes
 * of each one when its entry is written: media a change added, and the file's media a
 * checkpoint names. The composed package never holds their bytes.
 */
export function composedWithoutMedia(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedPackageCommentOptions = {}
): OrderedComposedWithoutMedia {
  const media = new Map<string, OrderedMediaReference>();
  const { pkg, takeovers } = composeWithMedia(
    doc,
    base,
    (item) => {
      media.set(partNameKey(item.part), item);
      return MEDIA_PLACE;
    },
    options
  );
  return { pkg, media, takeovers };
}

/**
 * The whole document as `.docx`, the same bytes as {@link orderedPackageBytes} gives for the
 * same media, handed to `options.write` front to back: each part when it is written, the
 * central directory last. Each media part is read through `options.media` when its turn
 * comes, one at a time, and is not kept after it is written: media a change added, and the
 * file's media a checkpoint names by reference. `options.comments` keeps or removes the
 * comments: two calls on one document, one with each, give the two files in either order.
 *
 * @throws OrderedRefusal as `orderedPackageBytes` does, and `missing-blob` when media cannot
 *   be read. The bytes written before a refusal are not a whole file.
 * @public
 */
export async function writeOrderedPackage(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedPackageWriteOptions & OrderedPackageCommentOptions
): Promise<void> {
  const { pkg, media } = composedWithoutMedia(doc, base, options);
  await writeComposedEntries(
    pkg,
    async (name, data) => {
      const held = media.get(partNameKey(name));
      return held === undefined ? data : orderedMediaOf(held, options.media);
    },
    options.write
  );
}

/** An entry's bytes as a writer writes them, or null to leave the entry out. */
type OrderedEntryBytes = (
  name: string,
  data: Uint8Array,
  entry: number
) => Promise<Uint8Array | null>;

/**
 * The entries of a composed package as a zip, handed to `write` front to back through the
 * engine's zip writer: each with the bytes `entryBytes` gives (`entry` is its place, from 0),
 * then the entry `last` gives, then the central directory. One entry at a time: an entry's
 * bytes are not kept once it is written.
 */
export async function writeComposedEntries(
  pkg: OoxmlPackage,
  entryBytes: OrderedEntryBytes,
  write: OrderedPackageWriteOptions['write'],
  last?: () => readonly [string, Uint8Array]
): Promise<void> {
  const zip = zipPartWriter();
  const writePart = async (name: string, bytes: Uint8Array): Promise<void> => {
    for (const chunk of finished(zip.part(name, bytes))) {
      await write(chunk);
    }
  };
  let entry = 0;
  for (const [name, data] of finished(ooxmlPackageEntriesSteps(pkg))) {
    const bytes = await entryBytes(name, data, entry);
    entry += 1;
    if (bytes !== null) {
      await writePart(name, bytes);
    }
  }

  if (last !== undefined) {
    await writePart(...last());
  }

  await write(zip.end());
}

/** The bytes of media, or `missing-blob` when `read` rejects or gives another size. */
export async function orderedMediaOf(
  media: OrderedMediaReference,
  read: OrderedPackageWriteOptions['media']
): Promise<Uint8Array> {
  let data: Uint8Array;
  try {
    data = await read(media.sha256);
  } catch {
    throw new OrderedRefusal('missing-blob', media.sha256);
  }

  if (data.byteLength !== media.size) {
    throw new OrderedRefusal('missing-blob', media.sha256);
  }

  return data;
}

/** Run `steps` to the end. */
function finished<T>(steps: Generator<void, T>): T {
  for (;;) {
    const next = steps.next();
    if (next.done) {
      return next.value;
    }
  }
}
