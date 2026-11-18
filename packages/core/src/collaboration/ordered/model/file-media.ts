// The file's media that a room checkpoint names by reference (design D14): the bytes are in the
// room's storage, under their SHA-256, and the package holds none for them.

import { readsAsXmlPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { resolveRelationship } from '../../../store/package/relationships.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  orderedMediaType,
  type OrderedFileMedia,
  type OrderedMediaReference,
  type OrderedMediaType,
} from './resources.ts';
import type { OrderedDocument } from './types.ts';

/**
 * The type of media a checkpoint names by reference, or null: a type the room stores that the
 * package reader keeps as bytes. An ink part is read as XML, so it stays in the checkpoint.
 */
export function referencedMediaType(contentType: string): OrderedMediaType | null {
  const type = orderedMediaType(contentType);
  return type !== null && !readsAsXmlPart(type) ? type : null;
}

/** A part as a checkpoint names it by reference, or null when the checkpoint keeps it. */
export type OrderedPartReference = (
  part: string,
  bytes: Uint8Array
) => OrderedMediaReference | null;

/**
 * How a checkpoint names each part of `pkg`: media of a type `referencedMediaType` gives, with
 * at least one byte, by reference. A part the package root relates (the thumbnail) stays in the
 * checkpoint: it is the file's, not the document's.
 */
export function partReferences(pkg: OoxmlPackage): OrderedPartReference {
  const root = new Set<string>();
  for (const record of pkg.relationships.get('/') ?? []) {
    const resolved = resolveRelationship(record);
    if (resolved.mode === 'Internal' && resolved.target.ok) {
      root.add(partNameKey(resolved.target.partName));
    }
  }

  return (part, bytes) => {
    const contentType = referencedMediaType(resolveContentTypeOf(pkg, part) ?? '');
    if (contentType === null || bytes.byteLength === 0 || root.has(partNameKey(part))) {
      return null;
    }

    return { part, contentType, sha256: orderedBlobHash(bytes), size: bytes.byteLength };
  };
}

/** Record the file's media that a checkpoint names by reference: their part names are taken. */
export function recordFileMedia(doc: OrderedDocument, media: readonly OrderedFileMedia[]): void {
  for (const item of media) {
    doc.resources.fileParts.add(item.part.toLowerCase());
    doc.resources.fileMedia.set(partNameKey(item.part), item);
  }
}

/** The file's media at `part`, or undefined. */
export function fileMediaAt(doc: OrderedDocument, part: string): OrderedFileMedia | undefined {
  return doc.resources.fileMedia.get(partNameKey(part));
}

/** Whether the file's media holds `sha256`: the room's storage holds what a checkpoint names. */
export function isFileMediaHash(doc: OrderedDocument, sha256: string): boolean {
  for (const item of doc.resources.fileMedia.values()) {
    if (item.sha256 === sha256) {
      return true;
    }
  }

  return false;
}
