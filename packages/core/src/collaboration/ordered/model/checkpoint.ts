// What every open of a room's file checks of its package: the manifest of a room file of version 5
// or older, which no engine reads now (design D14), and the media a checkpoint names by reference.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { normalizePartName, partNameKey } from '../../../store/package/opc-names.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { referencedMediaType } from './file-media.ts';
import type { OrderedFileMedia } from './resources.ts';

/**
 * The manifest's part name: a room file of version 5 or older is the Word file with this part
 * beside it. No engine reads it now: every open refuses a package that holds it.
 */
const CHECKPOINT_MANIFEST = '/ordered-checkpoint.json';

/** Whether a package holds a part under the manifest's name, in any case as OPC compares names. */
export function holdsCheckpointManifest(pkg: OoxmlPackage): boolean {
  const key = partNameKey(CHECKPOINT_MANIFEST);
  for (const name of pkg.partBytes.keys()) {
    if (partNameKey(name) === key) {
      return true;
    }
  }

  return false;
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
