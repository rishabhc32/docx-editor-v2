// The notes a decision hid, settled after it (group 9.9).

import {
  noteItemsOf,
  notesStoryOfReference,
  notesWithHiddenReferences,
  referenceKeyOf,
} from '../compose/note-references.ts';
import { isHidden } from '../model/document.ts';
import { relabel } from '../model/mutations.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';

/**
 * Mark gone the shown items of each note whose last shown reference `decided` hid. A note shows
 * only while one of its references shows (`compose/note-references.ts`), and compose leaves it
 * out of the saved file, so its items go in the model too, in the decision's undo record.
 */
export function settleHiddenNotes(doc: OrderedDocument, decided: readonly string[]): void {
  const references = new Map<string, OrderedItem>();
  for (const id of decided)
    for (const item of doc.labelled.get(id) ?? []) {
      const key = referenceKeyOf(item);
      if (key !== undefined && isHidden(doc, item)) references.set(key, item);
    }

  if (references.size === 0) return;
  const orphans = notesWithHiddenReferences(doc);
  for (const [key, reference] of references) {
    const story = orphans.has(key) ? notesStoryOfReference(doc, reference) : null;
    if (!story) continue;
    for (const item of noteItemsOf(story, key))
      if (!isHidden(doc, item)) relabel(doc, story, item, { gone: true });
  }
}
