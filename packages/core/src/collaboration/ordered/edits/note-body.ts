// A note's text goes with its references: struck with the last one, taken back with a new one.

import {
  noteItemsOf,
  notesStoryOfReference,
  notesWithHiddenReferences,
  referenceKeyOf,
  struckNotes,
} from '../compose/note-references.ts';
import { isHidden, pendingBy, pendingByOther, pendingDeleters } from '../model/document.ts';
import type { OrderedMaker } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { strikeItems, takeBack, type Ctx } from './context.ts';

/** Whether an item carries a pending change of someone other than `maker`. */
function othersOn(doc: OrderedDocument, item: OrderedItem, maker: OrderedMaker): boolean {
  const other = (id: string) => pendingByOther(doc, id, maker);
  return (
    (item.by !== null && other(item.by)) ||
    item.dels.some(other) ||
    item.fmts.some((f) => other(f.suggestion)) ||
    item.props.some((r) => other(r.suggestion)) ||
    item.within.by.some(other) ||
    item.within.dels.some(other)
  );
}

/**
 * Take back the new notes that showed only through references in `taken`, items an edit took
 * back for their author: a note that the author's own pending suggestion inserted goes whole,
 * with the author's later text in it, as the engine removes a note with its reference
 * (`cascadeDeletedNoteReferences`); else the room would show words the saved file no longer
 * holds. Another author's pending change in it is refused, as in the author's own new row
 * (`others-in-row`): taking the note back would take theirs back too.
 */
export function takeBackNotesOf(ctx: Ctx, taken: readonly OrderedItem[]): void {
  const references = taken.filter((item) => referenceKeyOf(item) !== undefined);
  if (references.length === 0) {
    return;
  }

  const orphans = notesWithHiddenReferences(ctx.doc);
  const maker = ctx.attribution;
  for (const reference of references) {
    const key = referenceKeyOf(reference)!;
    const story = orphans.has(key) ? notesStoryOfReference(ctx.doc, reference) : null;
    const items = story ? noteItemsOf(story, key) : [];

    // Only a new note: one a pending suggestion of the author inserted
    const open = items[0]?.by;
    const by = open ? pendingBy(ctx.doc, open, maker) : null;
    if (!story || !by) {
      continue;
    }

    if (items.some((item) => !isHidden(ctx.doc, item) && othersOn(ctx.doc, item, maker))) {
      throw new OrderedRefusal('others-in-row', "a new note that holds another author's change");
    }

    for (const item of items) {
      takeBack(ctx, story, item);
    }
  }
}

/**
 * Strike, for `suggestion`, the text of each note whose reference `struck` holds and whose shown
 * references are all struck by a pending suggestion: every shown unit and paragraph mark of the
 * note that no pending suggestion strikes yet, typed text of the suggestion's own author too.
 * Word writes a deleted note so (the engine's `note-body-strike.ts`; Word 2016 files
 * `RP050-Deleted-Footnote.docx` and `tdf141660.docx`), and the note's decision and undo go with
 * its reference's. Text typed in the note later is not struck.
 */
export function strikeStruckNoteBodies(
  ctx: Ctx,
  struck: readonly OrderedItem[],
  suggestion: string
): void {
  const references = struck.filter((item) => referenceKeyOf(item) !== undefined);
  if (references.length === 0) return;
  const deleters = struckNotes(ctx.doc);
  for (const reference of references) {
    const key = referenceKeyOf(reference)!;
    if (deleters.get(key) === undefined) continue;
    const story = notesStoryOfReference(ctx.doc, reference);
    if (!story) continue;
    for (const item of noteItemsOf(story, key)) {
      if (item.kind !== 'text' && item.kind !== 'atom' && item.kind !== 'mark') continue;
      if (item.prompt && item.kind !== 'mark') continue;
      if (isHidden(ctx.doc, item) || pendingDeleters(ctx.doc, item).length > 0) continue;
      strikeItems(ctx, story, [item], suggestion);
    }
  }
}
