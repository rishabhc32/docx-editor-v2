// The owners of new and deleted notes' content, read once the stories are loaded.

import {
  isNotesStory,
  noteKeyOf,
  referenceKeyOf,
  struckNotes,
} from '../compose/note-references.ts';
import { relabel } from '../model/mutations.ts';
import { adoptRevisionRoles } from '../model/revision-ids.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';

/** Whether two suggestions have one author and date. */
function sameStamp(doc: OrderedDocument, a: string, b: string): boolean {
  const x = doc.suggestions.get(a)?.attribution;
  const y = doc.suggestions.get(b)?.attribution;
  return x !== undefined && y !== undefined && x.author === y.author && x.date === y.date;
}

/**
 * Give each joined suggestion's `w:id`s to the one it joined, for the roles it had, and drop it
 * when nothing names it any more.
 */
function finishJoins(doc: OrderedDocument, joined: ReadonlyMap<string, string>): void {
  for (const [from, to] of joined) {
    adoptRevisionRoles(doc, from, to);
    if (!doc.labelled.has(from)) doc.suggestions.delete(from);
  }
}

/**
 * Every item of the notes whose owner `ownerOf` names from the note's open item, with that owner,
 * and whether the item is an edge of a note.
 */
function* notesWithOwner(
  doc: OrderedDocument,
  ownerOf: (open: OrderedItem) => string | null
): Generator<[OrderedStory, OrderedItem, string, boolean]> {
  for (const story of doc.stories.values()) {
    if (!isNotesStory(story)) {
      continue;
    }

    let owner: string | null = null;
    let depth = 0;
    for (const item of story.tree.toArray()) {
      const isNote = (item.kind === 'open' || item.kind === 'close') && item.token === 'note';
      if (isNote && item.kind === 'open') {
        if (depth === 0) {
          owner = ownerOf(item);
        }

        depth += 1;
      }

      if (owner !== null) {
        yield [story, item, owner, isNote];
      }

      if (isNote && item.kind === 'close') {
        depth -= 1;
        if (depth === 0) {
          owner = null;
        }
      }
    }
  }
}

/**
 * A note whose every citing reference a suggestion inserts is a new note: give its own content
 * the suggestion that inserted its reference. Word draws the new note's runs in a `w:ins` of their
 * own under the reference's author and date, and its paragraph marks plain (an older file drew the
 * whole note plain), so the notes part alone cannot say so. The plain items take the suggestion
 * (unmarked, as the note's paragraph marks are), and so do the items of an insertion with the
 * reference's author and date (the note's runs, joined with their `w:id`s). Text typed in it at
 * another time keeps its own label. A note that a reference of the file's own text cites is no
 * new note, though a suggestion inserts another reference to it: a move's copy of a reference
 * names the note its source names (null in `cited`).
 */
export function linkNewNotes(doc: OrderedDocument): void {
  const cited = new Map<string, string | null>();
  for (const story of doc.stories.values()) {
    if (isNotesStory(story)) {
      continue;
    }

    for (const item of story.tree.items()) {
      const key = referenceKeyOf(item);
      if (key !== undefined && cited.get(key) !== null) {
        cited.set(key, item.by);
      }
    }
  }

  if (cited.size === 0) {
    return;
  }

  const joined = new Map<string, string>();
  const ownerOf = (open: OrderedItem) => cited.get(noteKeyOf(open) ?? '') ?? null;
  for (const [story, item, owner, isNote] of notesWithOwner(doc, ownerOf)) {
    if (isNote) {
      relabel(doc, story, item, { by: owner });
    } else if (item.by === null) {
      relabel(doc, story, item, { by: owner, quiet: true });
    } else if (item.by !== owner && sameStamp(doc, item.by, owner)) {
      joined.set(item.by, owner);
      relabel(doc, story, item, { by: owner });
    }
  }

  finishJoins(doc, joined);
}

/**
 * A note whose references a suggestion struck is a deleted note: Word marks its body deleted
 * under the reference's author and date. Give each deleted note's body marks to the suggestion
 * that struck its references, when they have its author and date: they are one decision, as the
 * engine decides them. The joined suggestion keeps its marks' `w:id`s for the roles it had, and
 * goes when nothing names it.
 */
export function linkStruckNotes(doc: OrderedDocument): void {
  const struck = struckNotes(doc);
  if (struck.size === 0) return;
  const joined = new Map<string, string>();
  const ownerOf = (open: OrderedItem) => struck.get(noteKeyOf(open) ?? '') ?? null;
  for (const [story, item, struckBy, isNote] of notesWithOwner(doc, ownerOf)) {
    if (isNote || item.dels.length === 0) {
      continue;
    }

    const dels = item.dels.map((d) =>
      d !== struckBy && sameStamp(doc, d, struckBy) ? struckBy : d
    );
    if (dels.every((d, k) => d === item.dels[k])) {
      continue;
    }

    for (const d of item.dels) {
      if (d !== struckBy && sameStamp(doc, d, struckBy)) {
        joined.set(d, struckBy);
      }
    }

    relabel(doc, story, item, { dels: [...new Set(dels)] });
  }

  finishJoins(doc, joined);
}
