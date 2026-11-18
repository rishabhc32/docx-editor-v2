// A note shows only while its citing reference shows, as Word removes a note with its reference.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { isHidden, pendingDeleters } from '../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { attributeOf } from '../ooxml-nodes.ts';

const NOTES_ROOTS = new Set(['footnotes', 'endnotes']);
const REFERENCE = /^(footnote|endnote)Reference$/;

/** Whether a story is a notes part. */
export const isNotesStory = (story: OrderedStory): boolean =>
  NOTES_ROOTS.has(story.part.root.localName);

/** A note reference's kind (`footnote`, `endnote`) and node, or null for any other item. */
function referenceOf(
  item: OrderedItem
): { readonly kind: string; readonly node: OoxmlElement } | null {
  const node = item.node;
  if (item.kind !== 'atom' || !node || node.kind === 'textValue') return null;
  const kind = REFERENCE.exec(node.localName)?.[1];
  return kind === undefined ? null : { kind, node };
}

/** The `kind|id` a note reference names, or undefined for any other item. */
export function referenceKeyOf(item: OrderedItem): string | undefined {
  const reference = referenceOf(item);
  const id = reference ? attributeOf(reference.node, 'id') : undefined;
  return reference === null || id === undefined ? undefined : `${reference.kind}|${id}`;
}

/** Each note reference in the stories that are not notes parts, with its `kind|id`. */
function* referencesIn(doc: OrderedDocument): Generator<[OrderedItem, string]> {
  for (const story of doc.stories.values()) {
    if (isNotesStory(story)) continue;
    for (const item of story.tree.items()) {
      const key = referenceKeyOf(item);
      if (key !== undefined) yield [item, key];
    }
  }
}

/**
 * `kind|id` of the notes whose references the stories hold and show nowhere. A note's items carry
 * the reference's suggestion when it is new (`linkNewNotes`), but a decision on a container can
 * hide the reference alone (a row rejected or struck for good), so compose leaves out every such
 * note. A note the file holds with no reference at all stays as the file had it.
 */
export function notesWithHiddenReferences(doc: OrderedDocument): Set<string> {
  const shown = new Set<string>();
  const hidden = new Set<string>();
  for (const [item, key] of referencesIn(doc)) (isHidden(doc, item) ? hidden : shown).add(key);
  for (const key of shown) hidden.delete(key);
  return hidden;
}

/**
 * `kind|id` of the notes whose shown references are all struck, to the suggestion that shows as
 * striking the first of them. Word marks such a note's body deleted with its reference, under the
 * reference's author and date (the engine's `note-body-strike.ts`), so compose draws the body as
 * struck by that suggestion, and its decision decides both.
 */
export function struckNotes(doc: OrderedDocument): Map<string, string> {
  const deleters = new Map<string, string | null>();
  for (const [item, key] of referencesIn(doc)) {
    if (isHidden(doc, item)) continue;
    const deleter = pendingDeleters(doc, item)[0] ?? null;
    if (!deleters.has(key)) {
      deleters.set(key, deleter);
    } else if (deleter === null) {
      deleters.set(key, null);
    }
  }

  const out = new Map<string, string>();
  for (const [key, deleter] of deleters) if (deleter !== null) out.set(key, deleter);
  return out;
}

/** The `kind|id` of a note's open item, if it is one. */
export function noteKeyOf(item: OrderedItem): string | undefined {
  const node = item.node;
  if (item.token !== 'note' || !node || node.kind === 'textValue') return undefined;
  const id = attributeOf(node, 'id');
  return id === undefined ? undefined : `${node.localName}|${id}`;
}

/** The items of the note `key` names in a notes story, from its open item to its close item. */
export function noteItemsOf(story: OrderedStory, key: string): OrderedItem[] {
  const out: OrderedItem[] = [];
  let inside: string | null = null;
  for (const item of story.tree.toArray()) {
    if (inside === null) {
      if (item.kind !== 'open' || item.token !== 'note' || noteKeyOf(item) !== key) continue;
      inside = item.pair ?? null;
    }

    out.push(item);
    if (item.kind === 'close' && item.pair === inside) break;
  }

  return out;
}

/**
 * The items of every note compose leaves out (`notesWithHiddenReferences`): its references show
 * nowhere, as after the organiser's direct deletion of the row that held them. Word and the
 * engine remove such a note with its reference, so nothing in it shows, a suggestion on no card.
 * The items stay as they are: the organiser's undo brings the reference and the note back.
 */
export function leftOutNoteItems(doc: OrderedDocument): Set<OrderedItem> {
  const out = new Set<OrderedItem>();
  const orphans = notesWithHiddenReferences(doc);
  if (orphans.size === 0) return out;
  for (const story of doc.stories.values()) {
    if (!isNotesStory(story)) continue;
    let inside: string | null = null;
    for (const item of story.tree.items()) {
      if (inside === null && orphans.has(noteKeyOf(item) ?? '')) inside = item.pair ?? null;
      if (inside === null) continue;
      out.add(item);
      if (item.kind === 'close' && item.pair === inside) inside = null;
    }
  }

  return out;
}

/** The notes story a note reference names, or null for any other item. */
export function notesStoryOfReference(
  doc: OrderedDocument,
  item: OrderedItem
): OrderedStory | null {
  const reference = referenceOf(item);
  if (reference === null) return null;
  const root = `${reference.kind}s`;
  for (const story of doc.stories.values()) if (story.part.root.localName === root) return story;
  return null;
}
