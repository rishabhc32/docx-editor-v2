// Reads and writes a checkpoint's manifest (design D14).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { normalizePartName, partNameKey } from '../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { DEFAULT_ZIP_LIMITS } from '../../../store/package/zip.ts';
import { attempt } from '../change/read.ts';
import { sha256 } from '../change/resources.ts';
import {
  readTakeovers,
  takeoversJson,
  WORD_ID,
  type CheckpointTakeovers,
} from './checkpoint-takeovers.ts';
import { referencedMediaType } from './file-media.ts';
import { readPerson, type OrderedPerson } from './person.ts';
import type { OrderedFileMedia } from './resources.ts';

/** The manifest's part name. A Word file that holds this part is not opened as one. */
export const CHECKPOINT_MANIFEST = '/ordered-checkpoint.json';

const FORMAT = 'docx-editor.dev/ordered-checkpoint';
const VERSION = 3;
const MAX_MANIFEST_BYTES = 4 * 1024 * 1024;

/**
 * A person, the revision `w:id`s of their items' marks and the `w:id`s of their comments, as the
 * manifest lists them.
 */
export interface CheckpointPerson {
  readonly person: OrderedPerson;
  readonly revisions: readonly string[];
  readonly comments: readonly string[];
}

/** The persons a checkpoint names: by revision `w:id`, and by comment `w:id`. */
export interface CheckpointPersons {
  readonly revisions: ReadonlyMap<string, OrderedPerson>;
  readonly comments: ReadonlyMap<string, OrderedPerson>;
}

/**
 * Why a checkpoint does not open. Stable.
 *
 * - `not-checkpoint`: the file holds no manifest.
 * - `checkpoint-format`: the manifest names another format.
 * - `checkpoint-version`: the manifest names another version of the format.
 * - `checkpoint-manifest`: the manifest does not read, or does not fit the file.
 */
export type CheckpointRefusal =
  | 'not-checkpoint'
  | 'checkpoint-format'
  | 'checkpoint-version'
  | 'checkpoint-manifest';

/**
 * The manifest's bytes: one JSON part with the format, its version, each media part the
 * checkpoint holds by reference, the person of each room item by its revision `w:id`s and of
 * each room comment by its comment `w:id`, and the records each change record holds by takeover.
 */
export function checkpointManifestBytes(
  media: readonly OrderedFileMedia[],
  persons: readonly CheckpointPerson[],
  takeovers: CheckpointTakeovers
): Uint8Array {
  const listed = media.map(({ part, sha256, contentType, size, entry }) => ({
    part,
    sha256,
    contentType,
    size,
    entry,
  }));
  return new TextEncoder().encode(
    JSON.stringify({
      format: FORMAT,
      version: VERSION,
      media: listed,
      persons,
      ...takeoversJson(takeovers),
    })
  );
}

/** The part under the manifest's name, in any case as OPC compares names, and its bytes. */
function manifestEntry(pkg: OoxmlPackage): readonly [string, Uint8Array] | undefined {
  const key = partNameKey(CHECKPOINT_MANIFEST);
  for (const entry of pkg.partBytes) {
    if (partNameKey(entry[0]) === key) {
      return entry;
    }
  }

  return undefined;
}

/** Whether a package holds a part under the manifest's name. */
export function holdsCheckpointManifest(pkg: OoxmlPackage): boolean {
  return manifestEntry(pkg) !== undefined;
}

/**
 * What a checkpoint's package holds: the package without the manifest, its media, the person of
 * each room item by revision `w:id` and of each room comment by comment `w:id`, and the takeovers.
 */
export type CheckpointRead =
  | {
      readonly ok: true;
      readonly base: OoxmlPackage;
      readonly media: readonly OrderedFileMedia[];
      readonly persons: CheckpointPersons;
      readonly takeovers: CheckpointTakeovers;
    }
  | { readonly ok: false; readonly reason: CheckpointRefusal };

/**
 * A read checkpoint package as the package without the manifest and the media it names. Each
 * medium is a part the package does not hold, of the type the package's content types give it,
 * at its own place among the entries. Every value of the manifest comes from a file.
 */
export function readCheckpoint(pkg: OoxmlPackage): CheckpointRead {
  const entry = manifestEntry(pkg);
  if (entry === undefined) {
    return { ok: false, reason: 'not-checkpoint' };
  }

  const [name, bytes] = entry;
  const manifest = manifestOf(bytes);
  if (!manifest.ok) {
    return manifest;
  }

  // A manifest that reads as JSON is no XML part: the package holds it as bytes only
  const partBytes = new Map(pkg.partBytes);
  partBytes.delete(name);
  const base: OoxmlPackage = Object.freeze({ ...pkg, partBytes });
  const media = mediaIn(base, manifest.media);
  const persons = personsIn(manifest.persons);
  const takeovers = attempt(() => readTakeovers(manifest.takeovers, manifest.suggestions));
  if (media === null || persons === null || !takeovers.ok) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  return { ok: true, base, media, persons, takeovers: takeovers.value };
}

type Manifest =
  | {
      readonly ok: true;
      readonly media: readonly unknown[];
      readonly persons: readonly unknown[];
      readonly takeovers: unknown;
      readonly suggestions: unknown;
    }
  | { readonly ok: false; readonly reason: CheckpointRefusal };

/** The manifest's lists, after its format and version. */
function manifestOf(bytes: Uint8Array): Manifest {
  if (bytes.byteLength > MAX_MANIFEST_BYTES) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  if (!isRecord(value)) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  if (value.format !== FORMAT) {
    return { ok: false, reason: 'checkpoint-format' };
  }

  if (value.version !== VERSION) {
    return { ok: false, reason: 'checkpoint-version' };
  }

  const { media, persons } = value;
  if (!Array.isArray(media) || media.length > DEFAULT_ZIP_LIMITS.maxEntries) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  if (!Array.isArray(persons)) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  return { ok: true, media, persons, takeovers: value.takeovers, suggestions: value.suggestions };
}

/**
 * The listed persons by revision `w:id` and by comment `w:id`, or null when one does not read or
 * an id is named twice in its list. The manifest's byte cap bounds the lists.
 */
function personsIn(listed: readonly unknown[]): CheckpointPersons | null {
  const revisions = new Map<string, OrderedPerson>();
  const comments = new Map<string, OrderedPerson>();
  for (const value of listed) {
    if (!isRecord(value)) {
      return null;
    }

    const read = attempt(() => readPerson(value.person, 'person'));
    if (
      !read.ok ||
      !namedIds(value.revisions, read.value, revisions) ||
      !namedIds(value.comments, read.value, comments)
    ) {
      return null;
    }
  }

  return { revisions, comments };
}

/** Name a person by each id of a list; false when the list does not read or names an id twice. */
function namedIds(ids: unknown, person: OrderedPerson, byId: Map<string, OrderedPerson>): boolean {
  if (!Array.isArray(ids)) {
    return false;
  }

  for (const id of ids) {
    if (typeof id !== 'string' || !WORD_ID.test(id) || byId.has(id)) {
      return false;
    }

    byId.set(id, person);
  }

  return true;
}

/** The listed media, each checked against `base`, or null when one does not fit. */
function mediaIn(base: OoxmlPackage, listed: readonly unknown[]): OrderedFileMedia[] | null {
  const taken = new Set<string>();
  for (const name of base.partBytes.keys()) {
    taken.add(partNameKey(name));
  }

  const entries = new Set<number>();
  const total = base.partBytes.size + listed.length;
  const media: OrderedFileMedia[] = [];
  for (const value of listed) {
    const item = mediaItem(value);
    if (item === null || item.entry >= total || entries.has(item.entry)) {
      return null;
    }

    // Another part's name, or a type the content types do not give it
    const key = partNameKey(item.part);
    if (taken.has(key) || resolveContentTypeOf(base, item.part) !== item.contentType) {
      return null;
    }

    taken.add(key);
    entries.add(item.entry);
    media.push(item);
  }

  return media;
}

/** One listed medium, or null when a field is missing or out of range. */
function mediaItem(value: unknown): OrderedFileMedia | null {
  if (!isRecord(value)) {
    return null;
  }

  const { part, contentType, size, entry } = value;
  if (typeof part !== 'string' || typeof contentType !== 'string') {
    return null;
  }

  const name = normalizePartName(part);
  const type = referencedMediaType(contentType);
  const hash = attempt(() => sha256(value.sha256, 'sha256'));
  if (!name.ok || name.partName !== part || !hash.ok || type === null) {
    return null;
  }

  if (!isCount(size) || size === 0 || !isCount(entry)) {
    return null;
  }

  return { part, sha256: hash.value, contentType: type, size, entry };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isCount(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}
