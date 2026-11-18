import { validateEmbeddedImageBytes } from '../../../store/package/drawing-package-edit.ts';
import {
  sniffImageMime,
  validateJpegHeader,
  validateTiffHeader,
  type SupportedImageMime,
} from '../../../store/package/image-resources.ts';
import { readOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { resolveImageResourceLimits } from '../../../store/runtime/limits.ts';
import { ORDERED_RESOURCE_LIMITS, orderedMediaCap } from '../change/resources.ts';
import type { OrderedMediaType } from '../model/resources.ts';
import { isOrderedObjectBytes } from '../protocol/blobs.ts';

/**
 * How many first bytes of media a host keeps while the rest streams to storage (1 MiB). They
 * decide {@link checkOrderedMediaPrefix} for most media.
 *
 * @public
 */
export const ORDERED_MEDIA_PREFIX = 1_048_576;

/**
 * Every signature and fixed header the checks read lies in these first bytes. An ink part's `<`
 * lies the farthest: after at most 4,096 bytes of white space.
 */
const SIGNATURE_BYTES = 8_192;

const SVG_NAMESPACE_URI = 'http://www.w3.org/2000/svg';

/** EMF's header record type, its least size, and its signature (" EMF", little-endian). */
const EMR_HEADER = 1;
const EMF_HEADER_SIZE = 88;
const ENHMETA_SIGNATURE = 0x464d4520;

/** WMF's placeable key and the size of that header. */
const WMF_PLACEABLE_KEY = 0x9ac6cdd7;
const WMF_PLACEABLE_SIZE = 22;

const viewOf = (bytes: Uint8Array): DataView =>
  new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

/**
 * UTF-8 SVG text whose root element is `svg` in the SVG namespace, read by the engine's bounded
 * part reader (no DTD, no entity, bounded size, elements and depth).
 */
function isSvgDocument(bytes: Uint8Array): boolean {
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return false;
  }

  const read = readOoxmlPart(
    text,
    { name: '/word/media/ordered.svg', contentType: 'image/svg+xml' },
    { maxBytes: ORDERED_RESOURCE_LIMITS.svgSize }
  );
  return (
    read.ok &&
    read.part.root.namespaceUri === SVG_NAMESPACE_URI &&
    read.part.root.localName === 'svg'
  );
}

/** A TIFF whose first image's header is within the engine's pixel caps. */
function isTiff(bytes: Uint8Array): boolean {
  const header = validateTiffHeader(bytes);
  if (header === null) return false;
  const limits = resolveImageResourceLimits();
  const { pixelWidth: width, pixelHeight: height } = header;
  return (
    width <= limits.maxDimension &&
    height <= limits.maxDimension &&
    width * height <= limits.maxPixels
  );
}

/**
 * An EMF header record of at least 88 bytes, with " EMF" at byte 40 ([MS-EMF] 2.3.4.2), inside
 * a file of `fileSize` bytes. `bytes` are the file's first bytes.
 */
function isEmf(bytes: Uint8Array, fileSize: number): boolean {
  if (bytes.byteLength < EMF_HEADER_SIZE) return false;
  const view = viewOf(bytes);
  const size = view.getUint32(4, true);
  return (
    view.getUint32(0, true) === EMR_HEADER &&
    view.getUint32(40, true) === ENHMETA_SIGNATURE &&
    size >= EMF_HEADER_SIZE &&
    size % 4 === 0 &&
    size <= fileSize
  );
}

/**
 * A WMF standard header ([MS-WMF] 2.3.2.2: type 1 or 2, nine words, version 1 or 3) at the
 * start, or after a placeable header.
 */
function isWmf(bytes: Uint8Array): boolean {
  const view = viewOf(bytes);
  const placeable = bytes.byteLength >= 4 && view.getUint32(0, true) === WMF_PLACEABLE_KEY;
  const at = placeable ? WMF_PLACEABLE_SIZE : 0;
  if (bytes.byteLength < at + 18) return false;
  const type = view.getUint16(at, true);
  const version = view.getUint16(at + 4, true);
  return (
    (type === 1 || type === 2) &&
    view.getUint16(at + 2, true) === 9 &&
    (version === 0x0100 || version === 0x0300)
  );
}

/**
 * Whether bytes are media of `type`, by content, within the media cap (an SVG within its own
 * cap). Every byte is untrusted. The engine never runs them: it paints an SVG through an image
 * element and keeps the other kinds, and a chart's workbook, an object or ink, as they are.
 *
 * @public
 */
export function isOrderedMediaBytes(bytes: Uint8Array, type: OrderedMediaType): boolean {
  return mediaVerdict(bytes, type, bytes.byteLength) === 'valid';
}

/**
 * Whether the first bytes of media decide its check ({@link isOrderedMediaBytes} on the whole
 * file of `size` bytes): "valid" or "invalid", or "whole" when only the whole file can. They
 * decide "invalid" only by a wrong signature or fixed header, a size over the cap, or an image
 * header read in them and over the pixel caps. An SVG needs the whole file, and so does a JPEG
 * whose frame header or a TIFF whose first image directory lies past the first bytes. When the
 * first bytes are the whole file, this is the whole check.
 *
 * @public
 */
export function checkOrderedMediaPrefix(
  prefix: Uint8Array,
  media: { readonly contentType: OrderedMediaType; readonly size: number }
): 'valid' | 'invalid' | 'whole' {
  return mediaVerdict(prefix, media.contentType, media.size);
}

const verdict = (ok: boolean): 'valid' | 'invalid' => (ok ? 'valid' : 'invalid');

/** The check of media of `type` and `fileSize` bytes on its first bytes, `bytes`. */
function mediaVerdict(
  bytes: Uint8Array,
  type: OrderedMediaType,
  fileSize: number
): 'valid' | 'invalid' | 'whole' {
  if (
    !Number.isSafeInteger(fileSize) ||
    fileSize < 1 ||
    fileSize > orderedMediaCap(type) ||
    bytes.byteLength > fileSize
  ) {
    return 'invalid';
  }

  const whole = bytes.byteLength === fileSize;
  if (!whole && bytes.byteLength < SIGNATURE_BYTES) {
    return 'whole';
  }

  const kept = isOrderedObjectBytes(bytes, type);
  if (kept !== undefined) {
    return verdict(kept);
  }

  if (sniffImageMime(bytes) !== type) {
    return 'invalid';
  }

  // An SVG, and a JPEG or TIFF header past the first bytes, need the whole file
  switch (type) {
    case 'image/svg+xml':
      return whole ? verdict(isSvgDocument(bytes)) : 'whole';
    case 'image/tiff':
      if (!whole && validateTiffHeader(bytes) === null) {
        return 'whole';
      }

      return verdict(isTiff(bytes));
    case 'image/jpeg':
      if (!whole && validateJpegHeader(bytes) === null) {
        return 'whole';
      }

      return verdict(validateEmbeddedImageBytes(bytes, type));
    case 'image/x-emf':
      return verdict(isEmf(bytes, fileSize));
    case 'image/x-wmf':
      return verdict(isWmf(bytes));
    default:
      return verdict(validateEmbeddedImageBytes(bytes, type satisfies SupportedImageMime));
  }
}
