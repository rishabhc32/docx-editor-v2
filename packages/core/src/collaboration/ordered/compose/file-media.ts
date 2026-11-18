// The file's media that a room checkpoint names by reference, in a composed package: an empty
// part holds each one's place while compose runs, so the engine's own rules keep or drop it as
// they keep or drop a part with bytes, and the bytes go in last.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OrderedDocument } from '../model/types.ts';
import type { OrderedMediaBytes } from './resources.ts';

/** The empty bytes that hold a media part's place until its bytes go in. */
export const MEDIA_PLACE = new Uint8Array(0);

/**
 * `base` with an empty part for each of the file's referenced media parts, at its place among
 * the file's entries: the place decides the part's place in the saved file.
 */
export function withFileMediaPlaces(base: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  const fileMedia = doc.resources.fileMedia;
  if (fileMedia.size === 0) {
    return base;
  }

  const places = new Map<number, string>();
  for (const item of fileMedia.values()) {
    places.set(item.entry, item.part);
  }

  const partBytes = new Map<string, Uint8Array>();
  let at = 0;
  const placeNext = (): void => {
    for (;;) {
      const part = places.get(at);
      if (part === undefined) {
        return;
      }

      partBytes.set(part, MEDIA_PLACE);
      at += 1;
    }
  };
  for (const [name, bytes] of base.partBytes) {
    placeNext();
    partBytes.set(name, bytes);
    at += 1;
  }

  placeNext();
  return Object.freeze({ ...base, partBytes });
}

/**
 * Each of the file's referenced media parts that `pkg` still holds, with the bytes `bytes`
 * gives, or left out when it gives none (a screen, until they arrive). A part that compose
 * dropped is not read.
 */
export function withFileMediaBytes(
  pkg: OoxmlPackage,
  doc: OrderedDocument,
  bytes: OrderedMediaBytes
): OoxmlPackage {
  if (doc.resources.fileMedia.size === 0) {
    return pkg;
  }

  const partBytes = new Map(pkg.partBytes);
  for (const item of doc.resources.fileMedia.values()) {
    if (!partBytes.has(item.part)) {
      continue;
    }

    const data = bytes(item);
    if (data === undefined) {
      partBytes.delete(item.part);
    } else {
      partBytes.set(item.part, data);
    }
  }

  return Object.freeze({ ...pkg, partBytes });
}
