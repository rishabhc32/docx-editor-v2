// Writes a room checkpoint (design D14).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { CHECKPOINT_MANIFEST, checkpointManifestBytes } from '../model/checkpoint.ts';
import { partReferences, referencedMediaType } from '../model/file-media.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedFileMedia, OrderedMediaReference } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import { checkpointFacts } from './checkpoint-facts.ts';
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
 * {@link openOrderedCheckpoint} does.
 *
 * @throws OrderedRefusal as `writeOrderedPackage` does, and `too-many-items` when the manifest is
 *   past the engine's limit of one part (64 MiB), which no open reads. The bytes written before a
 *   refusal are not a whole file: the host keeps its last checkpoint and the room's log.
 * @public
 */
export async function writeOrderedCheckpoint(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedPackageWriteOptions
): Promise<void> {
  const typers = new Map<string, OrderedPerson>();
  const composed = composedWithoutMedia(doc, base, { typers });
  const { media } = composed;
  const { pkg, facts } = checkpointFacts(doc, composed.pkg, composed.takeovers, typers);
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
