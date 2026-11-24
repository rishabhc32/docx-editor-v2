// Notes' properties (group 9.9): their number format, start, restart and place.

import { sectionPropertiesOf } from '../compose/sections.ts';
import { settingKeeps } from '../compose/settings.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { addOrderedResource } from '../model/resources.ts';
import type { OrderedItem, OrderedNotePropertiesOp } from '../model/types.ts';
import { suggestionFor, type Ctx } from './context.ts';
import { settle } from './direct.ts';
import { addSectionWrites, sectionAt } from './sections.ts';

/**
 * Write the notes' properties of the section `holder` holds, or of the document when it is null,
 * as one direct edit: the engine writes them directly, and no file Word wrote shows a tracked
 * form. A section's write is a section record on `holder` (the mark that ends the section, or
 * the body's last section), which compose runs with the section's other writes; the document's
 * is a settings write (design D14). Both are the edit's suggestion, so `revert` takes them back.
 * Values the section or the settings show already make no suggestion. Returns the edit's
 * suggestion, if any.
 */
export function setNoteProperties(
  ctx: Ctx,
  holder: OrderedItem | null,
  op: OrderedNotePropertiesOp
): string | null {
  if (!ctx.direct) throw new OrderedRefusal('invalid-edit', 'note properties are written directly');
  if (holder !== null) {
    if (sectionPropertiesOf(holder) === null)
      throw new OrderedRefusal('invalid-edit', 'not the end of a section');
    if (isHidden(ctx.doc, holder)) throw new OrderedRefusal('hidden-target', 'the section');
    const made = addSectionWrites(ctx, [sectionAt(ctx, holder).section], op, null);
    settle(ctx.doc, made);
    return made;
  }

  if (settingKeeps(ctx.doc, op)) {
    return null;
  }

  const suggestion = suggestionFor(ctx);
  addOrderedResource(ctx.doc, {
    kind: 'setting',
    key: ctx.ids.resource(),
    suggestion: suggestion.id,
    op,
  });
  settle(ctx.doc, suggestion.id);
  return suggestion.id;
}
