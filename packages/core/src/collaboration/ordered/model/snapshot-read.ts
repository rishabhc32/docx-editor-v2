// Reads a room's saved item list (checkpoint version 6, design D14) from an untrusted file: strict
// JSON with a size limit and no prototype keys, every field through the field table's codec, the
// Word XML by value through the engine's bounded reader, then the document's own rules.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { XML_HARD_MAX_BYTES } from '../../../store/package/xml-reader.ts';
import { attempt, fieldsOf, Invalid, required } from '../change/read.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  createOrderedDocument,
  indexOrderedStory,
  isSplitEdge,
  orderedStoryTree,
  suggestionsOf,
} from './document.ts';
import { personsFit } from './person.ts';
import type { OrderedFileMedia } from './resources.ts';
import { readXmlString, type SnapshotReadContext } from './snapshot-codec.ts';
import { SnapshotNodeReader } from './snapshot-nodes.ts';
import { restoreSide } from './snapshot-side.ts';
import {
  DOCUMENT_CODEC,
  headerText,
  MEDIA_CODEC,
  SNAPSHOT_FORMAT,
  SNAPSHOT_VERSION,
} from './snapshot-write.ts';
import { isUnit, type OrderedDocument, type OrderedItem, type OrderedStory } from './types.ts';

/** Why a saved item list does not read: as a checkpoint's refusals name it. */
export type SnapshotRefusal = 'checkpoint-format' | 'checkpoint-version' | 'checkpoint-manifest';

export type SnapshotRead =
  | {
      readonly ok: true;
      readonly document: OrderedDocument;
      /** The room's last order number the file holds. */
      readonly order: number;
      /** The file media `ordered/base.docx` names by reference, not yet checked against it. */
      readonly media: readonly OrderedFileMedia[];
    }
  | { readonly ok: false; readonly reason: SnapshotRefusal };

const BODY_FIELDS = new Set(['media', 'document', 'side']);

/**
 * The document `checkpoint.json` holds, under the Word file number its loaded spaces hold. `base`
 * is the package `ordered/base.docx` holds, which the saved nodes name.
 */
export function readSnapshotJson(bytes: Uint8Array, base: OoxmlPackage): SnapshotRead {
  if (bytes.byteLength > XML_HARD_MAX_BYTES) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  // No reviver: every object goes through a strict reader, which refuses a key that reaches a
  // prototype
  let value: unknown;
  try {
    value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  const header = attempt(() => headerOf(value));
  if (!header.ok) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  const { format, version, reads, file, order } = header.value;
  if (format !== SNAPSHOT_FORMAT) {
    return { ok: false, reason: 'checkpoint-format' };
  }

  if (reads > SNAPSHOT_VERSION || version < SNAPSHOT_VERSION) {
    return { ok: false, reason: 'checkpoint-version' };
  }

  if (!bodyFits(bytes, header.value)) {
    return { ok: false, reason: 'checkpoint-manifest' };
  }

  const read = attempt(() => {
    const body = fieldsOf((value as { body: unknown }).body, 'body', BODY_FIELDS);
    return documentOf(body, base, file);
  });
  return read.ok
    ? { ok: true, ...read.value, order }
    : { ok: false, reason: 'checkpoint-manifest' };
}

/** The header's fields. */
function headerOf(value: unknown) {
  const f = fieldsOf(
    value,
    '',
    new Set(['format', 'version', 'reads', 'file', 'order', 'sha256', 'body'])
  );
  const count = (key: string) => {
    const read = required(f, key);
    if (typeof read !== 'number' || !Number.isSafeInteger(read) || read < 0) {
      throw new Invalid('invalid-number', key);
    }

    return read;
  };
  return {
    format: readXmlString(required(f, 'format'), 'format', 1, 256),
    version: count('version'),
    reads: count('reads'),
    file: count('file'),
    order: count('order'),
    sha256: readXmlString(required(f, 'sha256'), 'sha256', 64, 64),
  };
}

/**
 * Whether the file is the header's text, the body, and `}`, and the body's bytes give the
 * header's checksum: the body read is the body the writer hashed.
 */
function bodyFits(bytes: Uint8Array, header: ReturnType<typeof headerOf>): boolean {
  const head = new TextEncoder().encode(headerText(header));
  if (bytes.byteLength <= head.byteLength || bytes[bytes.byteLength - 1] !== 0x7d) {
    return false;
  }

  for (let k = 0; k < head.byteLength; k += 1) {
    if (bytes[k] !== head[k]) {
      return false;
    }
  }

  return orderedBlobHash(bytes.subarray(head.byteLength, bytes.byteLength - 1)) === header.sha256;
}

/**
 * The document and media of a checked body.
 *
 * @throws Invalid when a field does not read or the document breaks one of its rules.
 */
function documentOf(
  body: ReturnType<typeof fieldsOf>,
  base: OoxmlPackage,
  file: number
): { readonly document: OrderedDocument; readonly media: readonly OrderedFileMedia[] } {
  const nodes = new SnapshotNodeReader(base);
  const doc = createOrderedDocument(file);
  const stories: (readonly OrderedItem[])[] = [];
  const r: SnapshotReadContext = {
    nodes,
    item: (value, path) => {
      const [s, k] = Array.isArray(value) && value.length === 2 ? value : [];
      const found = typeof s === 'number' && typeof k === 'number' ? stories[s]?.[k] : undefined;
      if (found === undefined) {
        throw new Invalid('invalid-value', path, 'no item');
      }

      return found;
    },
    suggestions: (suggestions) => {
      for (const [id, suggestion] of suggestions) {
        doc.suggestions.set(id, suggestion);
      }

      return doc.suggestions;
    },
    tree: (items) => {
      stories.push(items);
      return orderedStoryTree(doc, items);
    },
    story: (story) => {
      if (doc.stories.has(story.name)) {
        throw new Invalid('repeated-id', 'stories', story.name);
      }

      return indexOrderedStory(doc, story);
    },
  };
  const media = MEDIA_CODEC.read(required(body, 'media'), 'media', r);
  const read = DOCUMENT_CODEC.read(required(body, 'document'), 'document', r);
  Object.assign(doc, read, { stories: doc.stories, suggestions: doc.suggestions });
  restoreSide(doc, required(body, 'side'), r);
  nodes.check();
  const fault = documentFault(doc, read.stories);
  if (fault !== null) {
    throw new Invalid('invalid-value', 'document', fault);
  }

  return { document: doc, media };
}

/**
 * What breaks the document's rules (design D14), or null: a person under another name, a story
 * under another name, a space in two stories, two items over one unit, a container edge with no
 * partner, a unit after a story's last paragraph mark, a label of no suggestion, a tree whose sums
 * do not add up.
 */
function documentFault(
  doc: OrderedDocument,
  stories: ReadonlyMap<string, OrderedStory>
): string | null {
  if (!personsFit(doc)) {
    return "a person under another person's name";
  }

  const spaces = new Set<string>();
  for (const [name, story] of stories) {
    if (story.name !== name) {
      return `story ${name}`;
    }

    // A split edge sums by its partner's items, which are indexed only after its tree is built
    for (const item of story.tree.items()) {
      if (isSplitEdge(item)) {
        story.tree.update(item);
      }
    }

    const fault = storyFault(doc, story, spaces);
    if (fault !== null) {
      return `${name}: ${fault}`;
    }
  }

  return null;
}

function storyFault(doc: OrderedDocument, story: OrderedStory, spaces: Set<string>): string | null {
  for (const [space, items] of story.spaces) {
    if (spaces.has(space)) {
      return `space ${space} in two stories`;
    }

    spaces.add(space);
    for (let k = 1; k < items.length; k += 1) {
      if (items[k - 1]!.offset + items[k - 1]!.length > items[k]!.offset) {
        return `two items over one unit of ${space}`;
      }
    }
  }

  const open = new Map<string, number>();
  let afterMark = false;
  for (const item of story.tree.items()) {
    for (const id of suggestionsOf(item)) {
      if (!doc.suggestions.has(id)) {
        return `a label of no suggestion: ${id}`;
      }
    }

    if (item.kind === 'open' || item.kind === 'close') {
      const key = `${item.token ?? ''} ${item.pair ?? ''}`;
      const depth = (open.get(key) ?? 0) + (item.kind === 'open' ? 1 : -1);
      if (depth < 0) {
        return `a close item with no open item: ${key}`;
      }

      open.set(key, depth);
    }

    afterMark = item.kind === 'mark' ? false : afterMark || isUnit(item);
  }

  for (const [key, depth] of open) {
    if (depth !== 0) {
      return `an open item with no close item: ${key}`;
    }
  }

  if (afterMark) {
    return 'a unit after the last paragraph mark';
  }

  const problems = story.tree.checkInvariants(
    (a, b) => a.units === b.units && a.shown === b.shown && a.marks === b.marks
  );
  return problems[0] ?? null;
}
