// The one table of the editor's ops (design D13): every group's file joined, typed over every
// `TreeDocOp` kind. A new op in the engine does not compile until a group lists it; the guard
// test (`__tests__/client/op-coverage.test.ts`) fails until each one is translated.

import type { TreeDocOpKind } from '../../../../store/store/tree-op-types.ts';
import { CONTROL_OPS, CONTROL_OPTIONS } from './controls.ts';
import { LINK_FIELD_OPS, LINK_FIELD_OPTIONS } from './links-fields.ts';
import { NOTE_OPS, NOTE_OPTIONS } from './notes.ts';
import { PARAGRAPH_OPS, PARAGRAPH_OPTIONS } from './paragraphs.ts';
import { PASTE_BLOCK_OPS, PASTE_BLOCK_OPTIONS } from './paste-blocks.ts';
import { PICTURE_OPS, PICTURE_OPTIONS } from './pictures.ts';
import { REVIEW_OPS } from './review.ts';
import { SECTION_OPS, SECTION_OPTIONS } from './sections.ts';
import { TABLE_OPS, TABLE_OPTIONS } from './tables.ts';
import { TEXT_OPS, TEXT_OPTIONS } from './text.ts';
import type { OrderedOpEntry, OrderedOpOption, OrderedOpTable } from './types.ts';

export { ORDERED_CHANGES_WITHOUT_OP } from './without-op.ts';
export { ORDERED_WORD_REFUSALS } from './word-refusals.ts';

/** The groups' op tables, in the order of `tasks.md`. Each op is in exactly one. */
export const ORDERED_OP_GROUPS = [
  TEXT_OPS,
  PARAGRAPH_OPS,
  REVIEW_OPS,
  PICTURE_OPS,
  LINK_FIELD_OPS,
  CONTROL_OPS,
  TABLE_OPS,
  SECTION_OPS,
  NOTE_OPS,
  PASTE_BLOCK_OPS,
] as const;

/** Every op of the editor and what the client makes of it. */
export const ORDERED_OP_COVERAGE: OrderedOpTable<TreeDocOpKind> = {
  ...TEXT_OPS,
  ...PARAGRAPH_OPS,
  ...REVIEW_OPS,
  ...PICTURE_OPS,
  ...LINK_FIELD_OPS,
  ...CONTROL_OPS,
  ...TABLE_OPS,
  ...SECTION_OPS,
  ...NOTE_OPS,
  ...PASTE_BLOCK_OPS,
};

/** The options that change what an op means, and their state. */
export const ORDERED_OP_OPTIONS: readonly OrderedOpOption[] = [
  ...TEXT_OPTIONS,
  ...PARAGRAPH_OPTIONS,
  ...PICTURE_OPTIONS,
  ...LINK_FIELD_OPTIONS,
  ...CONTROL_OPTIONS,
  ...TABLE_OPTIONS,
  ...SECTION_OPTIONS,
  ...NOTE_OPTIONS,
  ...PASTE_BLOCK_OPTIONS,
];

/** An op's entry. */
export function orderedOpEntry(kind: TreeDocOpKind): OrderedOpEntry {
  return ORDERED_OP_COVERAGE[kind];
}
