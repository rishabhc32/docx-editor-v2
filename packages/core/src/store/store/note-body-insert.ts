// A new note's body as Word marks it in suggesting mode, and its decision with its reference.

import { resolveNotesPart } from '../package/note-references.ts';
import { findNoteById, type NoteKind } from '../package/note-nodes.ts';
import { createNodeIdAllocator, findNode, replaceChildren } from '../package/ooxml-edit.ts';
import { withPart, type OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import {
  attributionOf,
  bodyNoteOf,
  everyMarked,
  isRevisionMark,
  marksAround,
  noteBodyParagraphIds,
  revisionOf,
  sameRevision,
  type NoteMarkSide,
  type NoteReferenceMarks,
  type NoteReferenceRevision,
} from './note-reference-marks.ts';
import { applyTreeOp } from './tree-op-apply.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { build, revisionAttributes } from './tree-op-tracked-builders.ts';

/**
 * The package with a new note's runs in one `w:ins` per paragraph, with the author and date of
 * the insertion around its reference (the run `referenceRunId` of a paragraph of `storyPart`).
 * Null when the reference has no insertion around it or an edit fails.
 *
 * Word marks the note's runs (its `w:footnoteRef` run and its text) with a `w:id` of their own,
 * and leaves the note's paragraph mark unmarked (Word 16 for Mac, 2026-10-29: a footnote made
 * through Word's object model with tracking on). Accepting the reference accepts the body
 * (`syncInsertedNoteBodies`), and rejecting it removes the note with its reference
 * (`cascadeDeletedNoteReferences`).
 */
export function markNewNoteBody(
  pkg: OoxmlPackage,
  kind: NoteKind,
  noteId: number,
  storyPart: OoxmlPart,
  paragraphId: string,
  referenceRunId: string
): OoxmlPackage | null {
  const reference = marksAround(storyPart, paragraphId, referenceRunId).ins;
  const part = resolveNotesPart(pkg, kind);
  const note = part && findNoteById(part.root, noteId);
  if (reference === null || !part || !note) return null;
  const mint = createNodeIdAllocator(part);
  const mintRevision = nextRevisionId(part);
  const attribution = attributionOf(reference);
  let current = part;
  for (const id of noteBodyParagraphIds(note)) {
    const paragraph = findNode(current, id);
    if (!paragraph || paragraph.kind !== 'paragraph') return null;
    const runs = paragraph.children.filter((child) => child.kind === 'run');
    if (runs.length === 0) continue;

    // Its own `w:id`, as Word gives it: apart from the reference's, which sits in another part
    let revisionId = mintRevision();
    if (revisionId === reference.id) revisionId = mintRevision();
    const wrapper = build(
      mint(),
      'revisionInsert',
      'ins',
      revisionAttributes(revisionId, attribution),
      runs
    );
    const at = paragraph.children.indexOf(runs[0]!);
    const rest = paragraph.children.filter((child) => child.kind !== 'run');
    const edited = replaceChildren(current, id, [...rest.slice(0, at), wrapper, ...rest.slice(at)]);
    if (!edited.ok) return null;
    current = edited.part;
  }

  return withPart(pkg, current);
}

/** The site nodes of the `side` marks in a note that `revisions` made, by revision address. */
function revisionSitesIn(
  note: OoxmlNode,
  side: NoteMarkSide,
  revisions: readonly NoteReferenceRevision[]
): Map<string, { id: string; revision: NoteReferenceRevision; ids: string[] }> {
  const sites = new Map<string, { id: string; revision: NoteReferenceRevision; ids: string[] }>();
  const stack: OoxmlNode[] = [note];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    if (isRevisionMark(node, side)) {
      const { id, ...revision } = revisionOf(node);
      if (id !== undefined && revisions.some((r) => sameRevision(r, revision))) {
        const key = `${id}\u0000${revision.author}\u0000${revision.date ?? ''}`;
        const entry = sites.get(key) ?? { id, revision, ids: [] as string[] };
        entry.ids.push(node.id);
        sites.set(key, entry);
      }
    }

    for (const child of node.children) {
      stack.push(child);
    }
  }

  return sites;
}

/**
 * A note body with the revision marks of one side (`w:del` or `w:ins`) that `revisions` made
 * accepted or rejected, run wrappers and paragraph marks alike. Null when the engine refuses.
 */
export function decideNoteBody(
  part: OoxmlPart,
  noteId: number,
  side: NoteMarkSide,
  revisions: readonly NoteReferenceRevision[],
  op: 'acceptRevision' | 'rejectRevision'
): OoxmlPart | null {
  const note = findNoteById(part.root, noteId);
  if (!note) return part;
  let current = part;
  for (const { id, revision, ids } of revisionSitesIn(note, side, revisions).values()) {
    const result = applyTreeOp(current, {
      op,
      revision: { id, ...revision },
      localName: side,
      siteNodeIds: ids,
    });
    if (!result.ok) return null;
    current = result.part;
  }

  return current;
}

/**
 * Accept the body insertion of each note whose references were all inserted and one of which a
 * transaction accepted (it still shows, with no insertion around it): the marks with the
 * attribution of the references' insertions. A note whose references went is the deletion
 * cascade's. Null when the engine refused a decision (the caller refuses the edit).
 */
export function syncInsertedNoteBodies(
  pkg: OoxmlPackage,
  was: ReadonlyMap<string, readonly NoteReferenceMarks[]>,
  now: ReadonlyMap<string, readonly NoteReferenceMarks[]>
): OoxmlPackage | null {
  let next = pkg;
  for (const [key, prior] of was) {
    const inserted = everyMarked(prior, 'ins');
    const list = now.get(key);
    if (inserted === null || list === undefined || list.length === 0) continue;
    if (everyMarked(list, 'ins') !== null) continue;
    const note = bodyNoteOf(key);
    if (note === null) continue;
    const part = resolveNotesPart(next, note.kind);
    if (!part) continue;
    const written = decideNoteBody(part, note.noteId, 'ins', inserted, 'acceptRevision');
    if (written === null) return null;
    if (written !== part) next = withPart(next, written);
  }

  return next;
}
