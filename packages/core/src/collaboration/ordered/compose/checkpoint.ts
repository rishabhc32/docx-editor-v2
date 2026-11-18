// Writes a room checkpoint (design D14).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { textboxFallbackExportPackage } from '../../../store/package/textbox-fallback-export.ts';
import { orderedStoryParts } from '../load/load.ts';
import {
  CHECKPOINT_MANIFEST,
  checkpointManifestBytes,
  type CheckpointPerson,
} from '../model/checkpoint.ts';
import {
  changeRecordsOf,
  takenProperties,
  type CheckpointSuggestion,
  type CheckpointTakeover,
  type CheckpointTakeovers,
} from '../model/checkpoint-takeovers.ts';
import { partReferences, referencedMediaType } from '../model/file-media.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedFileMedia, OrderedMediaReference } from '../model/resources.ts';
import { revisionIdsOf } from '../model/revision-ids.ts';
import type { OrderedAttribution, OrderedDocument } from '../model/types.ts';
import { attributeOf, eachElement } from '../ooxml-nodes.ts';
import {
  composedWithoutMedia,
  orderedMediaOf,
  writeComposedEntries,
  type OrderedPackageWriteOptions,
  type OrderedPartTakeovers,
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
 * with the person of each pending room item by the revision `w:id`s of its marks and of each
 * room comment by its `w:id`. The content types and the relationships stay as they are. The
 * room's storage holds the media bytes: `options.media` is read only for media a change added
 * that stays inside (an ink part), never for a picture. Word does not open a checkpoint;
 * {@link openOrderedCheckpoint} does.
 *
 * @throws OrderedRefusal as `writeOrderedPackage` does.
 * @public
 */
export async function writeOrderedCheckpoint(
  doc: OrderedDocument,
  base: OoxmlPackage,
  options: OrderedPackageWriteOptions
): Promise<void> {
  const { pkg, media, takeovers } = composedWithoutMedia(doc, base);
  const written = writtenRevisionIds(pkg);
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
    () => [
      CHECKPOINT_MANIFEST,
      checkpointManifestBytes(
        references,
        checkpointPersons(doc, written),
        checkpointTakeovers(doc, pkg, takeovers, written)
      ),
    ]
  );
}

/**
 * The person of each pending room item, with the revision `w:id`s the composed stories write for
 * its marks, and of each room comment, with its `w:id`: a load of the checkpoint reads each mark's
 * and each comment's person back by its `w:id`.
 */
function checkpointPersons(doc: OrderedDocument, written: ReadonlySet<string>): CheckpointPerson[] {
  const byPerson = new Map<string, PersonEntry>();
  const named = new Set<string>();
  for (const suggestion of doc.suggestions.values()) {
    const person = suggestion.attribution.person;
    if (person === null || suggestion.status !== 'pending') {
      continue;
    }

    // An id names one person: two suggestions share one only when two minted ids meet
    for (const id of revisionIdsOf(doc, suggestion.id)) {
      if (written.has(id) && !named.has(id)) {
        named.add(id);
        personEntry(byPerson, person).revisions.push(id);
      }
    }
  }

  for (const comment of doc.comments.values()) {
    if (comment.person !== null) {
      personEntry(byPerson, comment.person).comments.push(comment.id);
    }
  }

  return [...byPerson.values()];
}

/**
 * The records each change record of the composed stories holds by takeover, by where the file
 * writes it (a model address holds the file number), and the suggestions they were taken from,
 * each with the `w:id`s the file writes for its own marks.
 */
function checkpointTakeovers(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  held: OrderedPartTakeovers,
  written: ReadonlySet<string>
): CheckpointTakeovers {
  const takeovers: CheckpointTakeover[] = [];
  const suggestions: CheckpointSuggestion[] = [];
  if (![...held.values()].some((records) => records.size > 0)) {
    return { takeovers, suggestions };
  }

  const listed = new Map<string, number>();
  const indexOf = (id: string, attribution: OrderedAttribution): number => {
    let k = listed.get(id);
    if (k === undefined) {
      k = suggestions.length;
      listed.set(id, k);
      const revisions = [...revisionIdsOf(doc, id)].filter((r) => written.has(r));
      suggestions.push({ attribution, revisions });
    }

    return k;
  };

  // As a load reads the file: with each text box's legacy copy in step
  for (const [part, story] of orderedStoryParts(textboxFallbackExportPackage(pkg))) {
    const records = held.get(part);
    if (records === undefined || records.size === 0) {
      continue;
    }

    for (const { element, id, occurrence } of changeRecordsOf(story)) {
      const found = records.get(element.id);
      if (found === undefined) {
        continue;
      }

      const taken = found.taken.flatMap(({ replaced, after }) => {
        const suggestion = doc.suggestions.get(replaced);
        if (suggestion === undefined) {
          return [];
        }

        return [
          {
            suggestion: indexOf(replaced, suggestion.attribution),
            properties: takenProperties(found.record, after),
          },
        ];
      });
      takeovers.push({ part, record: element.localName, id, occurrence, taken });
    }
  }

  return { takeovers, suggestions };
}

/** A person's entry in the manifest while it is built. */
interface PersonEntry {
  readonly person: OrderedPerson;
  readonly revisions: string[];
  readonly comments: string[];
}

/** The entry of `person` in `byPerson`, made at its first use. */
function personEntry(byPerson: Map<string, PersonEntry>, person: OrderedPerson): PersonEntry {
  const key = JSON.stringify([person.id, person.name, person.org]);
  let entry = byPerson.get(key);
  if (entry === undefined) {
    entry = { person, revisions: [], comments: [] };
    byPerson.set(key, entry);
  }

  return entry;
}

/** The `w:id` of every revision mark in a package's stories. */
function writtenRevisionIds(pkg: OoxmlPackage): Set<string> {
  const ids = new Set<string>();
  for (const part of orderedStoryParts(pkg).values()) {
    eachElement(part.root, (element) => {
      const id = attributeOf(element, 'id');
      if (id !== undefined && attributeOf(element, 'author') !== undefined) {
        ids.add(id);
      }
    });
  }

  return ids;
}

/** Media the composed document holds by hash, or null when a checkpoint keeps it (an ink part). */
function heldReference(media: OrderedMediaReference): OrderedMediaReference | null {
  return referencedMediaType(media.contentType) === null ? null : media;
}
