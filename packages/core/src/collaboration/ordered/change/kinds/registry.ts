// The one table of change kinds the groups add (design, "How to add a group"). Each group keeps
// its kinds in its own file; this file only lists the files. The shared code asks
// `orderedKindSpecOf` for every kind that is not one of the format's first kinds.

import { isOrderedCoreKind, ORDERED_CORE_CHANGE_KINDS, type OrderedChangeKind } from '../types.ts';
import { CONTROL_KINDS, type OrderedControlChange } from './controls.ts';
import { LINK_FIELD_KINDS, type OrderedLinkFieldChange } from './links-fields.ts';
import { NOTE_KINDS, type OrderedNoteChange } from './notes.ts';
import { PASTE_BLOCK_KINDS, type OrderedPasteBlockChange } from './paste-blocks.ts';
import { PICTURE_KINDS, type OrderedPictureChange } from './pictures.ts';
import { SECTION_KINDS, type OrderedSectionChange } from './sections.ts';
import type { OrderedKindSpec } from './spec.ts';
import { TABLE_KINDS, type OrderedTableChange } from './tables.ts';

/**
 * Every change kind the groups add, as one union.
 *
 * @public
 */
export type OrderedGroupEditChange =
  | OrderedPictureChange
  | OrderedLinkFieldChange
  | OrderedControlChange
  | OrderedTableChange
  | OrderedSectionChange
  | OrderedNoteChange
  | OrderedPasteBlockChange;

/** The groups' spec tables, in the order of `tasks.md`. */
const GROUP_TABLES: readonly { readonly [kind: string]: OrderedKindSpec }[] = [
  PICTURE_KINDS,
  LINK_FIELD_KINDS,
  CONTROL_KINDS,
  TABLE_KINDS,
  SECTION_KINDS,
  NOTE_KINDS,
  PASTE_BLOCK_KINDS,
];

function specsOf(
  tables: readonly { readonly [kind: string]: OrderedKindSpec }[]
): ReadonlyMap<string, OrderedKindSpec> {
  const out = new Map<string, OrderedKindSpec>();
  for (const table of tables) {
    for (const [key, spec] of Object.entries(table)) {
      // A kind named twice, or under another key, would apply two ways on two copies
      if (key !== spec.kind || isOrderedCoreKind(key) || out.has(key)) {
        throw new Error(`The change kind ${key} is registered twice or under another name.`);
      }

      out.set(key, spec);
    }
  }

  return out;
}

const SPECS = specsOf(GROUP_TABLES);

/** The spec of a kind a group added; undefined for one of the format's first kinds. */
export function orderedKindSpecOf(kind: string): OrderedKindSpec | undefined {
  return SPECS.get(kind);
}

/**
 * Every kind of change, in the order the format lists them: the first kinds, then each
 * group's.
 *
 * @public
 */
export const ORDERED_CHANGE_KINDS: readonly OrderedChangeKind[] = Object.freeze([
  ...ORDERED_CORE_CHANGE_KINDS,
  // Every key is a group kind: `specsOf` checks it against its spec's own kind
  ...([...SPECS.keys()] as OrderedChangeKind[]),
]);
