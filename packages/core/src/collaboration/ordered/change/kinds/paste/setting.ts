// The `setting` change kind (group 9.10): a write to the settings part, replayed at compose.

import { settingKeeps } from '../../../compose/settings.ts';
import { contextOf, suggestionFor } from '../../../edits/context.ts';
import { settle } from '../../../edits/direct.ts';
import { addSetting } from '../../../edits/resources.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import { changeSetting, type OrderedChangeSetting } from '../../resources.ts';
import { fieldPath, required } from '../../read.ts';
import type { OrderedEditBase } from '../../types.ts';
import { orderedKindSpec } from '../spec.ts';

/**
 * Write a document setting: Protect Document on or off (`setDocumentProtection`). A direct
 * edit only: Word records no settings change as a suggestion (ECMA-376 Part 1, §17.15.1). A
 * party's is the room's `package` rule, as the host hashes `word/settings.xml`.
 *
 * @public
 */
export interface OrderedSettingChange extends OrderedEditBase {
  readonly kind: 'setting';
  readonly setting: OrderedChangeSetting;
}

export const SETTING_SPEC = orderedKindSpec<OrderedSettingChange>({
  kind: 'setting',
  fields: ['setting'],
  read: (fields) => ({
    setting: changeSetting(required(fields, 'setting'), fieldPath(fields, 'setting')),
  }),
  // A setting names no place in a story: the change's part only says where it was made
  touches: () => {},
  apply(change, { edit }) {
    if (edit.direct !== true)
      throw new OrderedRefusal('invalid-edit', 'Word records no settings change as a suggestion');
    const ctx = contextOf(edit, []);

    // A write of what the settings say already changes nothing: no record
    if (settingKeeps(ctx.doc, change.setting)) {
      return { suggestions: [], replaced: [] };
    }

    const suggestion = suggestionFor(ctx).id;
    addSetting(ctx, suggestion, change.setting);
    settle(ctx.doc, suggestion);
    return { suggestions: [suggestion], replaced: [] };
  },
  policy: { text: false, properties: false },
  // No paragraph changes: the package goes out with the settings part (`client/dirty.ts`)
  publish: 'paragraphs',
});
