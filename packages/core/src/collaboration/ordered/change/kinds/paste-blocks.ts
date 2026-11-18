// Change kinds of group 9.10 (paste with formatting, story block replacement, block deletion,
// new list definitions and levels, document protection and revision attribution). The group's
// own file: its change types and their specs, and nothing else edits it. Each spec lives in
// `paste/`.

import { ATTRIBUTION_SPEC, type OrderedAttributionChange } from './paste/attribution.ts';
import {
  DELETE_PARAGRAPH_SPEC,
  type OrderedDeleteParagraphChange,
} from './paste/delete-paragraph.ts';
import { INSERT_FRAGMENT_SPEC, type OrderedInsertFragmentChange } from './paste/fragment.ts';
import { LIST_LEVEL_SPEC, type OrderedListLevelChange } from './paste/list-level.ts';
import { NEW_LIST_SPEC, type OrderedNewListChange } from './paste/new-list.ts';
import { REPLACE_STORY_SPEC, type OrderedReplaceStoryChange } from './paste/replace-story.ts';
import { SETTING_SPEC, type OrderedSettingChange } from './paste/setting.ts';
import type { OrderedKindSpec } from './spec.ts';

export type { OrderedAttributionChange } from './paste/attribution.ts';
export type { OrderedDeleteParagraphChange } from './paste/delete-paragraph.ts';
export type { OrderedInsertFragmentChange } from './paste/fragment.ts';
export type { OrderedListLevelChange } from './paste/list-level.ts';
export type { OrderedNewListChange, OrderedNewListParagraph } from './paste/new-list.ts';
export type { OrderedReplaceStoryChange } from './paste/replace-story.ts';
export type { OrderedSettingChange } from './paste/setting.ts';

/**
 * The change kinds of group 9.10: a union of the group's change interfaces, each extending
 * `OrderedEditBase` with its own `kind`.
 *
 * @public
 */
export type OrderedPasteBlockChange =
  | OrderedSettingChange
  | OrderedDeleteParagraphChange
  | OrderedInsertFragmentChange
  | OrderedNewListChange
  | OrderedListLevelChange
  | OrderedReplaceStoryChange
  | OrderedAttributionChange;

/** The group's specs: one for each kind of {@link OrderedPasteBlockChange}. */
export const PASTE_BLOCK_KINDS: {
  readonly [K in OrderedPasteBlockChange['kind']]: OrderedKindSpec;
} = {
  setting: SETTING_SPEC,
  'delete-paragraph': DELETE_PARAGRAPH_SPEC,
  'insert-fragment': INSERT_FRAGMENT_SPEC,
  'new-list': NEW_LIST_SPEC,
  'list-level': LIST_LEVEL_SPEC,
  'replace-story': REPLACE_STORY_SPEC,
  attribution: ATTRIBUTION_SPEC,
};
