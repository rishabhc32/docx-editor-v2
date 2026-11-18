// A deleted note's body as Word marks it in suggesting mode, and the store's note cascade.

import { resolveNotesPart } from '../package/note-references.ts';
import { cascadeDeletedNoteReferences } from '../package/note-lifecycle.ts';
import { findNoteById } from '../package/note-nodes.ts';
import { findNode } from '../package/ooxml-edit.ts';
import { withPart, type OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { decideNoteBody, syncInsertedNoteBodies } from './note-body-insert.ts';
import {
  attributionOf,
  bodyNoteOf,
  everyMarked,
  noteBodyParagraphIds,
  noteReferenceMarks,
  sameRevision,
  type NoteReferenceMarks,
} from './note-reference-marks.ts';
import { paragraphOffsetIndex } from './tree-op-segments.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { applyDeleteTracked } from './tree-op-tracked-delete.ts';
import { applyParagraphMarkRevision } from './tree-op-tracked-marks.ts';

/**
 * The note body marked deleted by `deletion`, as Word marks it; null when the engine refuses.
 * When a tracked deletion strikes the last reference to a note, Word marks its runs (the
 * `w:footnoteRef` run as well) in `w:del`, and each paragraph mark with `w:pPr/w:rPr/w:del`,
 * under the reference's author and date (Word 2016 files: Open-Xml-PowerTools
 * `RP050-Deleted-Footnote.docx` and LibreOffice `tdf141660.docx`).
 */
function strikeNoteBody(
  part: OoxmlPart,
  noteId: number,
  deletion: RevisionAttributionInput
): OoxmlPart | null {
  const note = findNoteById(part.root, noteId);
  if (!note) return part;
  let current = part;
  for (const paragraphId of noteBodyParagraphIds(note)) {
    const paragraph = findNode(current, paragraphId);
    if (!paragraph || paragraph.kind !== 'paragraph') return null;
    const length = paragraphOffsetIndex(paragraph).length;
    if (length > 0) {
      const struck = applyDeleteTracked(current, paragraph, 0, length, deletion, undefined, {
        strikeOwn: true,
      });
      if (!struck.ok) return null;
      current = struck.part;
    }

    const marked = findNode(current, paragraphId);
    if (!marked || marked.kind !== 'paragraph') return null;
    const result = applyParagraphMarkRevision(
      current,
      marked as OoxmlParagraphNode,
      'del',
      deletion
    );
    if (!result.ok) return null;
    current = result.part;
  }

  return current;
}

/**
 * Keep each note's body marked as its references are: a note whose references a transaction
 * struck (all of them, none before) has its body marked deleted with the new deletion's author
 * and date; a note one of whose references a transaction gave back has those marks rejected.
 * `was` and `now` are the references' marks before and after the transaction. Null when the
 * engine refused a mark (the caller refuses the edit).
 */
export function syncStruckNoteBodies(
  after: OoxmlPackage,
  was: ReadonlyMap<string, readonly NoteReferenceMarks[]>,
  now: ReadonlyMap<string, readonly NoteReferenceMarks[]>
): OoxmlPackage | null {
  let next = after;
  for (const [key, list] of now) {
    const prior = was.get(key);
    if (prior === undefined || list.length === 0) continue;
    const struck = everyMarked(list, 'del');
    const struckBefore = everyMarked(prior, 'del');
    if ((struck === null) === (struckBefore === null)) continue;
    const note = bodyNoteOf(key);
    if (note === null) continue;
    const part = resolveNotesPart(next, note.kind);
    if (!part) continue;
    const priorDeletions = prior.flatMap((marks) => (marks.del === null ? [] : [marks.del]));
    let written: OoxmlPart | null;
    if (struck !== null) {
      const deletion =
        struck.find((d) => !priorDeletions.some((k) => sameRevision(k, d))) ?? struck[0]!;
      written = strikeNoteBody(part, note.noteId, attributionOf(deletion));
    } else {
      written = decideNoteBody(part, note.noteId, 'del', priorDeletions, 'rejectRevision');
    }

    if (written === null) return null;
    if (written !== part) next = withPart(next, written);
  }

  return next;
}

/**
 * The package store's note cascade: the notes whose references went are deleted, then the
 * bodies of the notes whose references were struck, given back or accepted follow them. Null
 * when a scan ran out of budget or the engine refused a mark.
 */
export function cascadeNoteReferenceChanges(
  before: OoxmlPackage,
  after: OoxmlPackage
): OoxmlPackage | null {
  const cascaded = cascadeDeletedNoteReferences(before, after);
  if (cascaded === null) return null;
  const was = noteReferenceMarks(before.parts.values());
  const now = noteReferenceMarks(cascaded.parts.values());
  if (was === null || now === null) return null;
  const struck = syncStruckNoteBodies(cascaded, was, now);
  return struck === null ? null : syncInsertedNoteBodies(struck, was, now);
}
