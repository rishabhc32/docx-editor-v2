// Reads and writes a checkpoint's manifest (design D14).

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { normalizePartName, partNameKey } from '../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { DEFAULT_ZIP_LIMITS } from '../../../store/package/zip.ts';
import { at, attempt, fieldsOf, Invalid, required } from '../change/read.ts';
import { sha256 } from '../change/resources.ts';
import {
  manifestList,
  readSuggestions,
  readTypers,
  suggestionsJson,
  typersJson,
  matching,
  WORD_ID,
  type CheckpointSuggestion,
  type CheckpointTypers,
} from './checkpoint-suggestions.ts';
import { readTakeovers, takeoversJson, type CheckpointTakeover } from './checkpoint-takeovers.ts';
import { readUntracked, untrackedJson, type CheckpointUntracked } from './checkpoint-untracked.ts';
import { referencedMediaType } from './file-media.ts';
import { readPerson, type OrderedPerson } from './person.ts';
import type { OrderedFileMedia } from './resources.ts';

/** The manifest's part name. A Word file that holds this part is not opened as one. */
export const CHECKPOINT_MANIFEST = '/ordered-checkpoint.json';

const FORMAT = 'docx-editor.dev/ordered-checkpoint';
const VERSION = 4;
const MAX_MANIFEST_BYTES = 4 * 1024 * 1024;

/**
 * The room facts a checkpoint's manifest keeps, which Word has no place for: each pending
 * suggestion whole, the person of each room comment by its `w:id`, the records each change
 * record holds by takeover, who typed the units of each insertion, and the records Word writes
 * no change record for. Every section names a suggestion by its place in `suggestions`.
 */
export interface CheckpointFacts {
  readonly suggestions: readonly CheckpointSuggestion[];
  readonly comments: ReadonlyMap<string, OrderedPerson>;
  readonly takeovers: readonly CheckpointTakeover[];
  readonly typers: readonly CheckpointTypers[];
  readonly untracked: readonly CheckpointUntracked[];
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
 * checkpoint holds by reference, and the room's facts.
 */
export function checkpointManifestBytes(
  media: readonly OrderedFileMedia[],
  facts: CheckpointFacts
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
      suggestions: suggestionsJson(facts.suggestions),
      comments: [...facts.comments].map(([id, person]) => ({ id, person })),
      takeovers: takeoversJson(facts.takeovers),
      typers: typersJson(facts.typers),
      untracked: untrackedJson(facts.untracked),
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

/** What a checkpoint's package holds: the package without the manifest, its media and facts. */
export type CheckpointRead =
  | {
      readonly ok: true;
      readonly base: OoxmlPackage;
      readonly media: readonly OrderedFileMedia[];
      readonly facts: CheckpointFacts;
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
  const facts = attempt(() => factsOf(manifest.value));
  if (media === null || !facts.ok) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  return { ok: true, base, media, facts: facts.value };
}

/**
 * The manifest's facts, each section checked.
 *
 * @throws Invalid when one does not read.
 */
function factsOf(manifest: Record<string, unknown>): CheckpointFacts {
  const suggestions = readSuggestions(manifest.suggestions);
  const count = suggestions.length;
  return {
    suggestions,
    comments: readComments(manifest.comments),
    takeovers: readTakeovers(manifest.takeovers, count),
    typers: readTypers(manifest.typers, count),
    untracked: readUntracked(manifest.untracked, count),
  };
}

type Manifest =
  | {
      readonly ok: true;
      readonly media: readonly unknown[];
      readonly value: Record<string, unknown>;
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

  const { media } = value;
  if (!Array.isArray(media) || media.length > DEFAULT_ZIP_LIMITS.maxEntries) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  return { ok: true, media, value };
}

/**
 * The person of each listed comment by its `w:id`, each id named once.
 *
 * @throws Invalid when one does not read.
 */
function readComments(value: unknown): Map<string, OrderedPerson> {
  const persons = new Map<string, OrderedPerson>();
  for (const [k, entry] of manifestList(value, 'comments').entries()) {
    const path = at('comments', k);
    const f = fieldsOf(entry, path, new Set(['id', 'person']));
    const id = matching(required(f, 'id'), at(path, 'id'), WORD_ID, 'a comment id');
    if (persons.has(id)) {
      throw new Invalid('invalid-value', at(path, 'id'), 'named twice');
    }

    persons.set(id, readPerson(required(f, 'person'), at(path, 'person')));
  }

  return persons;
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
