// Media bytes are named by SHA-256 and never travel in a change (design D14).

import { sha256FontBytes } from '../../../store/package/sha256.ts';
import { isCompoundFile, isXmlText } from '../../../store/store/clipboard-fragment-objects.ts';
import {
  ORDERED_EMBEDDED_OBJECT_TYPES,
  ORDERED_INK_TYPE,
  ORDERED_OLE_OBJECT_TYPE,
  type OrderedMediaType,
} from '../model/resources.ts';

const REFUSALS = ['blob-too-large', 'blob-hash', 'blob-type', 'blob-quota'] as const;

/**
 * Why the storage refused a blob. Stable.
 *
 * - `blob-too-large`: over the cap for its type (`orderedMediaCap`).
 * - `blob-hash`: the bytes' SHA-256 is not the one the client named.
 * - `blob-type`: the bytes are not of the content type the client named.
 * - `blob-quota`: the room or the person stored as much as they may.
 *
 * @public
 */
export type OrderedBlobRefusalCode = (typeof REFUSALS)[number];

/** Whether a value is one of the stable refusal codes. */
export function isOrderedBlobRefusalCode(value: unknown): value is OrderedBlobRefusalCode {
  return REFUSALS.some((code) => code === value);
}

/** The SHA-256 of bytes, as 64 lower-case hex digits. @public */
export function orderedBlobHash(bytes: Uint8Array): string {
  return sha256FontBytes(bytes).replace(/^sha256:/, '');
}

/** A zip's local file header (`PK\x03\x04`): an embedded Office file's signature. */
export function isZipSignature(bytes: Uint8Array): boolean {
  return (
    bytes.byteLength >= 4 &&
    bytes[0] === 0x50 &&
    bytes[1] === 0x4b &&
    bytes[2] === 0x03 &&
    bytes[3] === 0x04
  );
}

/**
 * Whether bytes kept as they are have their type's signature, as the editor alone checks them
 * on copy and paste: an OLE part is a compound file, an embedded Office file (and a chart's
 * workbook) a zip, an ink part XML text. Never opened, read or run. Undefined for an image
 * type, which is checked as an image.
 */
export function isOrderedObjectBytes(
  bytes: Uint8Array,
  type: OrderedMediaType
): boolean | undefined {
  if (type === ORDERED_OLE_OBJECT_TYPE) return isCompoundFile(bytes);
  if (type === ORDERED_INK_TYPE) return isXmlText(bytes);
  if (ORDERED_EMBEDDED_OBJECT_TYPES.includes(type)) return isZipSignature(bytes);
  return undefined;
}
