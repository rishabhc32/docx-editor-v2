// Notes (group 9.9): a new footnote or endnote, its notes part when the file has none, and the
// notes a reference names.

import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
} from '../../../store/package/ooxml-tree.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { addOrderedStory, labelUnindex, mainStory } from '../model/document.ts';
import { insertItems } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem, OrderedStory } from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  attributeOf,
  childNamed,
  childWhere,
  firstElement,
  isRunProperties,
} from '../ooxml-nodes.ts';
import { withNewIds, type Ctx } from './context.ts';
import { addRelationshipResource, addStoryResource, emptyStoryRoot } from './resources.ts';
import { insertText, noteOwnAddition, unitsOf } from './text.ts';
import { templateItems } from './templates.ts';

const W = WML_NAMESPACE_URI;

/** A footnote or an endnote. @public */
export type OrderedNoteKind = 'footnote' | 'endnote';

/** The notes part's root element of a kind. */
export const notesRootOf = (kind: OrderedNoteKind): 'footnotes' | 'endnotes' =>
  kind === 'footnote' ? 'footnotes' : 'endnotes';

const CONTENT_TYPES = {
  footnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  endnotes: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
} as const;

/** The notes story of a kind, or null when the document has none. */
export function notesStoryOf(doc: Ctx['doc'], kind: OrderedNoteKind): OrderedStory | null {
  const root = notesRootOf(kind);
  for (const story of doc.stories.values())
    if (story.part.root.localName === root && story.part.root.namespaceUri === W) return story;
  return null;
}

/** A node parsed from engine literals (every value in it is a constant or a computed integer). */
function parsed(xml: string, name = '/word/ordered-note.xml'): OoxmlElement {
  const read = readOoxmlPart(xml, { name, contentType: 'application/xml' });
  if (!read.ok) throw new OrderedRefusal('engine-refused', 'a note template did not parse');
  const first = firstElement(read.part.root);
  if (!first) throw new OrderedRefusal('engine-refused', 'an empty note template');
  return first;
}

/**
 * Word's separator notes, as the engine writes them in a new notes part (`ensureNotesPart`): the
 * separator (`w:id="-1"`) and the continuation separator (`w:id="0"`).
 */
function separators(kind: OrderedNoteKind): OoxmlElement[] {
  const root = notesRootOf(kind);
  const note = (type: string, id: number, mark: string) =>
    parsed(
      `<w:${root} xmlns:w="${W}"><w:${kind} w:type="${type}" w:id="${id}">` +
        `<w:p><w:r><w:${mark}/></w:r></w:p></w:${kind}></w:${root}>`
    );
  return [
    note('separator', -1, 'separator'),
    note('continuationSeparator', 0, 'continuationSeparator'),
  ];
}

/**
 * A new notes story of a kind, with its separators, and the resources that write it: the story
 * part and the main document's relationship to it (design D14). Inside the caller's undo record.
 * The part belongs to the notes in it, not to the first note's suggestion: a later note of
 * another person goes into it too. Once its last note goes (deleted, undone, taken back or
 * rejected), it stays with its separators, as Word keeps a notes part and as the editor alone
 * does (`compose/resources.ts`).
 */
export function addNotesStory(ctx: Ctx, kind: OrderedNoteKind, suggestion: string): OrderedStory {
  const doc = ctx.doc;
  const role = notesRootOf(kind);
  const main = mainStory(doc);
  if (!main) throw new OrderedRefusal('unknown-story', 'the document has no main part');
  const resource = addStoryResource(ctx, suggestion, role);
  const name = resource.part;
  const root = emptyStoryRoot(role, name, CONTENT_TYPES[role]);
  const story = addOrderedStory(doc, name, root, []);
  recordUndo(doc, () => {
    for (const item of story.tree.items()) labelUnindex(doc, item);
    doc.stories.delete(name);
  });
  const inStory: Ctx = { ...ctx, story };
  const items = separators(kind).flatMap((note) => templateItems(inStory, note, suggestion));
  insertItems(doc, story, null, items);
  addRelationshipResource(ctx, ctx.ids.resource(), suggestion, main.name, role, {
    resource: resource.key,
  });
  return story;
}

/** The largest note id Word reads ([MS-OI29500] 17.11.2(a)). */
export const WORD_NOTE_ID_MAX = 32767;

/**
 * New note ids for a story, one per call: past every note in the story, its own new ones and
 * hidden ones included, from 1. Past Word's largest id (a file can hold it), the lowest free one:
 * Word numbers notes by the order of their references, not by their ids. Every copy holds the
 * same story, so every copy picks the same ids.
 */
export function noteIdsFor(story: OrderedStory): () => number {
  const used = new Set<number>();
  let max = 0;
  for (const item of story.tree.items()) {
    if (item.kind !== 'open' || item.token !== 'note') continue;
    const node = item.node;
    const id = node && node.kind !== 'textValue' ? Number(attributeOf(node, 'id')) : Number.NaN;
    if (!Number.isSafeInteger(id)) continue;
    used.add(id);
    if (id > max) max = id;
  }

  let free = 1;
  return () => {
    if (max < WORD_NOTE_ID_MAX) {
      max += 1;
      return max;
    }

    while (free <= WORD_NOTE_ID_MAX && used.has(free)) free += 1;
    if (free > WORD_NOTE_ID_MAX)
      throw new OrderedRefusal(
        'too-many-items',
        `every note id up to ${WORD_NOTE_ID_MAX} is taken`
      );
    used.add(free);
    return free;
  };
}

/**
 * The items of a new note of a kind and id, all of them the suggestion's: as Word writes a
 * tracked new note (Word 16 for Mac, 2026-10-29), compose draws its runs (its `w:footnoteRef` run
 * and its text) in a `w:ins` of the suggestion with a `w:id` of their own, and leaves its
 * paragraph marks unmarked (`quiet`).
 */
export function newNoteItems(
  ctx: Ctx,
  notes: OrderedStory,
  kind: OrderedNoteKind,
  id: number,
  suggestion: string
): OrderedItem[] {
  const root = notesRootOf(kind);
  const style = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
  const note = parsed(
    `<w:${root} xmlns:w="${W}"><w:${kind} w:id="${id}"><w:p><w:r><w:rPr>` +
      `<w:rStyle w:val="${style}"/></w:rPr><w:${kind}Ref/></w:r>` +
      `<w:r><w:t xml:space="preserve"></w:t></w:r></w:p></w:${kind}></w:${root}>`
  );

  // Fresh items, in no story yet: their labels are set before they go in
  return templateItems({ ...ctx, story: notes }, note, suggestion).map(
    (item): OrderedItem =>
      item.kind === 'mark' ? { ...item, by: suggestion, quiet: true } : { ...item, by: suggestion }
  );
}

/** A note reference run's face and atom, of a kind and id, from engine literals. */
export function noteReference(
  ctx: Ctx,
  kind: OrderedNoteKind,
  id: number
): { readonly face: OoxmlElement; readonly atom: OoxmlElement } {
  const style = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
  const body = parsed(
    `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:rStyle w:val="${style}"/></w:rPr>` +
      `<w:${kind}Reference w:id="${id}"/></w:r></w:p></w:body></w:document>`
  );
  const paragraph = firstElement(body);
  const run = paragraph ? firstElement(paragraph) : undefined;
  const face = run ? childWhere(run, isRunProperties) : undefined;
  const atom = run ? childNamed(run, `${kind}Reference`) : undefined;
  if (!face || !atom) throw new OrderedRefusal('engine-refused', 'a note reference template');
  return { face: withNewIds(ctx, face), atom: withNewIds(ctx, atom) };
}

/**
 * A new footnote or endnote at a view offset of the main document: its reference, typed as an
 * object, and its note at the end of the notes story, in one suggestion. A file with no notes
 * part of the kind gets one. Inside the author's own earlier suggestion, the reference and the
 * note are noted for the change's undo.
 */
export function insertNote(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  kind: OrderedNoteKind
): string {
  if (ctx.story.part.root.localName !== 'document')
    throw new OrderedRefusal('note-outside-body', 'a note reference outside the main document');
  const existing = notesStoryOf(ctx.doc, kind);
  const id = existing ? noteIdsFor(existing)() : 1;
  const { face, atom } = noteReference(ctx, kind, id);
  const sid = insertText(ctx, p, offset, '', {
    pieces: [{ text: '￼', face, content: atom }],
  });
  const notes = existing ?? addNotesStory(ctx, kind, sid);
  const items = newNoteItems(ctx, notes, kind, id, sid);
  insertItems(ctx.doc, notes, notes.tree.last(), items);
  noteOwnAddition(ctx, sid, unitsOf(notes, items));
  return sid;
}

/**
 * Text typed in the note `insertNote` just made, at the end of its paragraph: part of the
 * note's suggestion, inside its insertion as the rest of the note's runs are.
 */
export function typeInNewNote(
  ctx: Ctx,
  suggestion: string,
  kind: OrderedNoteKind,
  text: string
): void {
  const notes = notesStoryOf(ctx.doc, kind);
  if (!notes) throw new OrderedRefusal('unknown-story', `the document has no ${notesRootOf(kind)}`);
  const p = orderedViewParagraph(ctx.doc, notes.name, notes.tree.summary.marks - 1);
  if (!p || p.mark.by !== suggestion)
    throw new OrderedRefusal('malformed-story', 'the new note has no paragraph');
  const into: Ctx = { ...ctx, story: notes, join: suggestion, direct: false };
  insertText(into, p, p.chars.length, text);
}
