// A deleted note, and a new one, is ONE decision in the review queue: the reference's card.

import { findNoteById } from '../package/note-nodes.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import {
  everyMarked,
  isNotesPart,
  noteBodyParagraphIds,
  noteReferenceMarks,
  parseNoteKey,
  sameRevision,
  type NoteMarkSide,
  type NoteReferenceRevision,
} from './note-reference-marks.ts';
import type { ReviewRevisionItem } from './review-items.ts';

/** The side of a note body mark an item is, when it lies in the notes parts only. */
function bodySide(item: ReviewRevisionItem, notes: ReadonlySet<string>): NoteMarkSide | null {
  if (item.ranges.length === 0 || !item.ranges.every((range) => notes.has(range.partName))) {
    return null;
  }

  const mark = item.revisionKind === 'paragraphMark' ? item.markDirection : undefined;
  if (item.revisionKind === 'delete' || mark === 'delete') return 'del';
  if (item.revisionKind === 'insert' || mark === 'insert') return 'ins';
  return null;
}

type BodyMarks = { readonly [side in NoteMarkSide]: readonly NoteReferenceRevision[] | null };

/**
 * The revisions without the marks of struck and new notes' bodies: those in a notes part, over
 * paragraphs of one note whose references are all struck (or all inserted) by a mark of the
 * same author and date. Unchanged when no notes part holds a deletion or an insertion.
 *
 * The body's marks follow the reference's (`note-body-strike.ts`, `note-body-insert.ts`), so a
 * card of their own would decide half a note: accepting a struck body empties the note and
 * leaves its reference standing.
 */
export function withoutNoteBodyMarks(
  revisions: readonly ReviewRevisionItem[],
  parts: readonly OoxmlPart[]
): readonly ReviewRevisionItem[] {
  const notesParts = parts.filter(isNotesPart);
  if (notesParts.length === 0) return revisions;
  const notesNames = new Set(notesParts.map((part) => part.name));
  if (!revisions.some((item) => bodySide(item, notesNames) !== null)) return revisions;
  const marks = noteReferenceMarks(parts);
  if (marks === null) return revisions;

  // Each struck or new note's body paragraphs, to the marks of its references
  const bodies = new Map<string, BodyMarks>();
  for (const [key, list] of marks) {
    const body: BodyMarks = { del: everyMarked(list, 'del'), ins: everyMarked(list, 'ins') };
    if (body.del === null && body.ins === null) continue;
    const { kind, noteId } = parseNoteKey(key);
    const part = notesParts.find((p) => p.root.localName === `${kind}s`);
    const note = part && findNoteById(part.root, noteId);
    if (!note) continue;
    for (const paragraphId of noteBodyParagraphIds(note)) {
      bodies.set(`${part.name}\u0000${paragraphId}`, body);
    }
  }

  if (bodies.size === 0) return revisions;
  const covered = (item: ReviewRevisionItem): boolean => {
    const side = bodySide(item, notesNames);
    return (
      side !== null &&
      item.ranges.every((range) =>
        [range.start.paragraphId, range.end.paragraphId].every((paragraphId) =>
          bodies
            .get(`${range.partName}\u0000${paragraphId}`)
            ?.[side]?.some((d) => sameRevision(d, item))
        )
      )
    );
  };
  return revisions.filter((item) => !covered(item));
}
