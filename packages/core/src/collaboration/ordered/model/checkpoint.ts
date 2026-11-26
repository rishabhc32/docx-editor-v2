// Reads the manifest of a checkpoint of version 5 (design D14), which no engine writes now.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { normalizePartName, partNameKey } from '../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { DEFAULT_ZIP_LIMITS } from '../../../store/package/zip.ts';
import { at, attempt, Invalid, isForbiddenKey, plainFields, required } from '../change/read.ts';
import { sha256 } from '../change/resources.ts';
import {
  readLists,
  readResources,
  readStyles,
  type CheckpointList,
  type CheckpointResource,
  type CheckpointStyle,
} from './checkpoint-lists.ts';
import { readMadeOver, type CheckpointMadeOver } from './checkpoint-over.ts';
import {
  readNamedOnce,
  readNamespace,
  readSuggestions,
  xmlText,
  type CheckpointSuggestion,
} from './checkpoint-suggestions.ts';
import { readTakeovers, type CheckpointTakeover } from './checkpoint-takeovers.ts';
import {
  readChangeRecords,
  readUntracked,
  type CheckpointChangeRecords,
  type CheckpointUntracked,
} from './checkpoint-untracked.ts';
import { referencedMediaType } from './file-media.ts';
import { readPerson, type OrderedPerson } from './person.ts';
import type { OrderedFileMedia } from './resources.ts';

/** The manifest's part name. A Word file that holds this part is not opened as one. */
const CHECKPOINT_MANIFEST = '/ordered-checkpoint.json';

const FORMAT = 'docx-editor.dev/ordered-checkpoint';

/**
 * The format's version. A version adds sections and fields, and a reader skips those it does not
 * know: the version moves only for a change an older reader would misread.
 */
const VERSION = 5;

/**
 * The namespace of the attributes only a checkpoint's parts hold: `oc:label` on each element a
 * manifest entry names (a number unique in the checkpoint); `oc:typer` on each run, and each
 * unit not in a run, whose units a room change typed inside a pending insertion (its namespace);
 * and `oc:run` on each run drawn from units and on each of its text elements, a number that names
 * the run the room draws them in (`run`, face and format records) and the room's text element,
 * which a pending suggestion may split into pieces. An open takes them off before the load, from
 * a Word file too.
 */
export const CHECKPOINT_NAMESPACE = 'https://docx-editor.dev/ordered-checkpoint';

/**
 * The room facts a checkpoint's manifest keeps, which Word has no place for: each pending
 * suggestion whole, and each decided one another section names; the person of each room comment
 * by its `w:id`; the records each change record holds by takeover; the person of each namespace
 * the parts name as a typer; the records Word writes no change record for, decided or not; what
 * the room holds of a record that its change record does not say; the records made over a
 * change record's pending ones that it does not hold; and the list definitions, styles, list
 * levels, relationships, parts and media changes added. Every section names a suggestion by its
 * place in `suggestions`, and an element of the file by its label.
 */
export interface CheckpointFacts {
  readonly suggestions: readonly CheckpointSuggestion[];
  readonly commentPersons: ReadonlyMap<string, OrderedPerson>;
  readonly takeovers: readonly CheckpointTakeover[];
  readonly namespacePersons: ReadonlyMap<string, OrderedPerson>;
  readonly untracked: readonly CheckpointUntracked[];
  readonly changeRecords: readonly CheckpointChangeRecords[];
  readonly madeOver: readonly CheckpointMadeOver[];
  readonly lists: readonly CheckpointList[];
  readonly styles: readonly CheckpointStyle[];
  readonly resources: readonly CheckpointResource[];
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
  const facts = attempt(() => factsOf(manifest.fields));
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
    commentPersons: readComments(manifest.comments),
    takeovers: readTakeovers(manifest.takeovers, count),
    namespacePersons: readNamespaces(manifest.namespaces),
    untracked: readUntracked(manifest.untracked, count),
    // Sections a checkpoint of this version may not hold yet
    changeRecords:
      manifest.changeRecords === undefined ? [] : readChangeRecords(manifest.changeRecords, count),
    lists: manifest.lists === undefined ? [] : readLists(manifest.lists, count),
    styles: manifest.styles === undefined ? [] : readStyles(manifest.styles, count),
    resources: manifest.resources === undefined ? [] : readResources(manifest.resources, count),
    madeOver: manifest.madeOver === undefined ? [] : readMadeOver(manifest.madeOver, count),
  };
}

type Manifest =
  | {
      readonly ok: true;
      readonly media: readonly unknown[];
      readonly fields: Record<string, unknown>;
    }
  | { readonly ok: false; readonly reason: CheckpointRefusal };

/** The manifest's lists, after its format and version. */
function manifestOf(bytes: Uint8Array): Manifest {
  if (bytes.byteLength > XML_HARD_MAX_BYTES) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  // A key that reaches an object's prototype is refused anywhere, in a section a reader skips too
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes), (key, read) => {
      if (isForbiddenKey(key)) {
        throw new Invalid('forbidden-key', key);
      }

      return read;
    });
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

  return { ok: true, media, fields: value };
}

/**
 * The person of each listed comment by its `w:id`, each id named once.
 *
 * @throws Invalid when one does not read.
 */
function readComments(value: unknown): Map<string, OrderedPerson> {
  const comments = readNamedOnce(
    value,
    'comments',
    (entry, path) => {
      const f = plainFields(entry, path);
      return {
        id: xmlText(required(f, 'id'), at(path, 'id')),
        person: readPerson(required(f, 'person'), at(path, 'person')),
      };
    },
    (comment) => comment.id
  );
  return new Map(comments.map(({ id, person }) => [id, person]));
}

/**
 * The person of each listed namespace, each namespace named once.
 *
 * @throws Invalid when one does not read.
 */
function readNamespaces(value: unknown): Map<string, OrderedPerson> {
  const namespaces = readNamedOnce(
    value,
    'namespaces',
    (entry, path) => {
      const f = plainFields(entry, path);
      return {
        namespace: readNamespace(required(f, 'namespace'), at(path, 'namespace')),
        person: readPerson(required(f, 'person'), at(path, 'person')),
      };
    },
    (entry) => entry.namespace
  );
  return new Map(namespaces.map(({ namespace, person }) => [namespace, person]));
}

/** The listed media, each checked against `base`, or null when one does not fit. */
function mediaIn(base: OoxmlPackage, listed: readonly unknown[]): OrderedFileMedia[] | null {
  const media: OrderedFileMedia[] = [];
  for (const value of listed) {
    const item = mediaItem(value);
    if (item === null) {
      return null;
    }

    media.push(item);
  }

  return mediaFits(base, media) ? media : null;
}

/**
 * Whether each medium fits `base`, the package that holds them by reference: a part name of its
 * own in normal form, of a type the room stores and the content types give it, at its own place
 * among the entries.
 */
export function mediaFits(base: OoxmlPackage, media: readonly OrderedFileMedia[]): boolean {
  const taken = new Set<string>();
  for (const name of base.partBytes.keys()) {
    taken.add(partNameKey(name));
  }

  const entries = new Set<number>();
  const total = base.partBytes.size + media.length;
  for (const item of media) {
    const name = normalizePartName(item.part);
    if (!name.ok || name.partName !== item.part || referencedMediaType(item.contentType) === null) {
      return false;
    }

    if (item.entry >= total || entries.has(item.entry)) {
      return false;
    }

    // Another part's name, or a type the content types do not give it
    const key = partNameKey(item.part);
    if (taken.has(key) || resolveContentTypeOf(base, item.part) !== item.contentType) {
      return false;
    }

    taken.add(key);
    entries.add(item.entry);
  }

  return true;
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

  const type = referencedMediaType(contentType);
  const hash = attempt(() => sha256(value.sha256, 'sha256'));
  if (!hash.ok || type === null) {
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
