// What the group 9.9 ops name, found in the list: a note's references (a note deletion and a
// conversion read them; the editor names a note by its kind and its `w:id`, as the list
// composes it), and an equation object (the editor names an equation by its `m:oMath` node).

import { equationsOfAtom, isOmmlEquationAtom } from '../../../../store/package/omml-display.ts';
import { isNotesStory, notesWithHiddenReferences } from '../../compose/note-references.ts';
import { sectionPropertiesOf } from '../../compose/sections.ts';
import { notesStoryOf } from '../../edits/notes.ts';
import { isHidden, mainStory } from '../../model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../../model/types.ts';
import { attributeOf } from '../../ooxml-nodes.ts';
import { nextShownMark, paragraphOfMark, type ListParagraph } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import type { OrderedTranslateBatch } from './types.ts';

/** A shown note reference: its paragraph and its offset there. */
export interface NoteReferenceAt {
  readonly paragraph: ListParagraph;
  readonly offset: number;
  readonly item: OrderedItem;
}

/** Whether an item is the reference of a note of a kind, and of `id` when given. */
export function isNoteReference(
  item: OrderedItem,
  kind: 'footnote' | 'endnote',
  id?: number
): boolean {
  const node = item.node;
  if (item.kind !== 'atom' || !node || node.kind === 'textValue') return false;
  if (node.localName !== `${kind}Reference`) return false;
  return id === undefined || attributeOf(node, 'id') === String(id);
}

/**
 * The shown paragraph an item is in, read through the batch (its text is checked against the
 * editor's), and the item's offset there. Null when the item shows in no paragraph.
 */
function shownPlaceOf(
  batch: OrderedTranslateBatch,
  story: OrderedStory,
  item: OrderedItem
): { readonly paragraph: ListParagraph; readonly offset: number } | null {
  const mark = nextShownMark(batch.doc, story, item);
  const found = mark ? paragraphOfMark(batch.doc, story, mark) : null;
  const nodeId = found?.mark.node?.id;
  if (!found || nodeId === undefined) {
    return null;
  }

  const paragraph = batch.paragraph(nodeId);
  const offset = paragraph.view.chars.findIndex((char) => char.item === item);
  return offset < 0 ? null : { paragraph, offset };
}

/**
 * Every shown reference of the notes of a kind (of one note, with `id`), in document order: the
 * stories other than the notes parts, each in order. The editor's text of each paragraph is
 * checked through the batch, as for any op that names a paragraph.
 */
export function noteReferences(
  batch: OrderedTranslateBatch,
  kind: 'footnote' | 'endnote',
  id?: number
): NoteReferenceAt[] {
  const doc = batch.doc;
  const out: NoteReferenceAt[] = [];
  for (const story of doc.stories.values()) {
    if (isNotesStory(story)) continue;
    for (const item of story.tree.items()) {
      if (!isNoteReference(item, kind, id) || isHidden(doc, item)) {
        continue;
      }

      const at = shownPlaceOf(batch, story, item);
      if (!at) {
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a note reference');
      }

      out.push({ ...at, item });
    }
  }

  return out;
}

/** The notes of a kind that show, as their open items, in the notes story that holds them. */
export interface ShownNotes {
  readonly story: OrderedStory;
  readonly opens: readonly OrderedItem[];
}

/**
 * The notes of a kind that show (the one of `id`, when given), in their story's order: Word's
 * own notes (separators) are not among them, and a note shows while one of its references does
 * (or it has none, as the file had it). Null when the document has no notes of the kind.
 */
export function shownNotes(
  doc: OrderedDocument,
  kind: 'footnote' | 'endnote',
  id?: number
): ShownNotes | null {
  const story = notesStoryOf(doc, kind);
  if (!story) return null;
  const hidden = notesWithHiddenReferences(doc);
  const opens: OrderedItem[] = [];
  for (const item of story.tree.items()) {
    const node = item.node;
    if (item.kind !== 'open' || item.token !== 'note' || !node || node.kind === 'textValue')
      continue;
    const type = attributeOf(node, 'type');
    const noteId = attributeOf(node, 'id');
    if ((type !== undefined && type !== 'normal') || noteId === undefined) continue;
    if (id !== undefined && noteId !== String(id)) continue;
    if (!isHidden(doc, item) && !hidden.has(`${kind}|${noteId}`)) opens.push(item);
  }

  return { story, opens };
}

/** The main document's sections, by the items that hold their `w:sectPr`. */
export interface SectionHolders {
  readonly story: OrderedStory;
  readonly holders: readonly OrderedItem[];
}

/**
 * The items that hold the main document's sections' `w:sectPr`, in the engine's section order
 * (`collectSectionPropertyNodes`): the shown paragraph marks that end a section, outside tables
 * (a table is one block there), then the body's last section. Null without a main document.
 */
export function sectionHolders(doc: OrderedDocument): SectionHolders | null {
  const story = mainStory(doc);
  if (!story) {
    return null;
  }

  const holders: OrderedItem[] = [];
  let last: OrderedItem | null = null;
  let tables = 0;
  for (const item of story.tree.items()) {
    if (item.token === 'tbl' && (item.kind === 'open' || item.kind === 'close')) {
      tables += item.kind === 'open' ? 1 : -1;
      continue;
    }

    if (sectionPropertiesOf(item) === null || isHidden(doc, item)) continue;
    if (item.kind === 'zero') last = item;
    else if (tables === 0) holders.push(item);
  }

  return { story, holders: last ? [...holders, last] : holders };
}

/** A shown equation object, the equation of it the editor names, and its paragraph. */
export interface EquationAt {
  readonly paragraph: ListParagraph;
  readonly offset: number;
  readonly item: OrderedItem;
  /** The equation's place among the object's equations (0 for an `m:oMath`). */
  readonly equation: number;
  /** How many equations the object holds. */
  readonly count: number;
}

/**
 * The shown equation object that holds the `m:oMath` with this node id, in any story. The
 * editor's equation ids are the node ids of the package the client last published, which are
 * the list's own.
 */
export function equationOf(batch: OrderedTranslateBatch, equationId: string): EquationAt {
  const doc = batch.doc;
  for (const story of doc.stories.values())
    for (const item of story.tree.items()) {
      const node = item.node;
      if (item.kind !== 'atom' || !item.bare || !node || !isOmmlEquationAtom(node)) continue;
      const equations = equationsOfAtom(node);
      const equation = equations.findIndex((candidate) => candidate.id === equationId);
      if (equation < 0 || isHidden(doc, item)) {
        continue;
      }

      const at = shownPlaceOf(batch, story, item);
      if (!at) {
        break;
      }

      return { ...at, item, equation, count: equations.length };
    }

  throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no such equation');
}
