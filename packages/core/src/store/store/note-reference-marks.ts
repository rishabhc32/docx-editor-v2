// Who struck or inserted each note reference: the `w:del` and the `w:ins` around every footnote
// and endnote reference outside the notes parts.

import {
  createNoteReferenceScanBudget,
  collectNoteReferences,
  MAX_NOTE_REFERENCE_MUTATION_SCAN,
  type NoteReferenceScanBudget,
} from '../package/note-references.ts';
import type { NoteKind } from '../package/note-nodes.ts';
import { findNode } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI } from '../package/ooxml-shared.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { attributeValueOf } from './tree-op-nodes.ts';

/** Who made a revision mark around a note reference: the mark's attribution and `w:id`. */
export interface NoteReferenceRevision {
  readonly author: string;
  readonly date?: string;
  readonly id?: string;
}

/** The nearest deletion and insertion around one note reference (null for none). */
export interface NoteReferenceMarks {
  readonly del: NoteReferenceRevision | null;
  readonly ins: NoteReferenceRevision | null;
}

/** A side of a reference's marks: struck (`del`) or inserted (`ins`). */
export type NoteMarkSide = keyof NoteReferenceMarks;

const NOTES_ROOTS: ReadonlySet<string> = new Set(['footnotes', 'endnotes']);

/** Whether a part is a notes part (`w:footnotes` or `w:endnotes`). */
export const isNotesPart = (part: OoxmlPart): boolean => NOTES_ROOTS.has(part.root.localName);

/** Whether a node is a `w:del` or a `w:ins` (`side`). */
export const isRevisionMark = (node: OoxmlNode, side: NoteMarkSide): boolean =>
  node.kind !== 'textValue' && node.localName === side && node.namespaceUri === WML_NAMESPACE_URI;

/** A revision mark's attribution and `w:id`. */
export function revisionOf(mark: OoxmlNode): NoteReferenceRevision {
  const date = attributeValueOf(mark, 'date', WML_NAMESPACE_URI);
  const id = attributeValueOf(mark, 'id', WML_NAMESPACE_URI);
  return {
    author: attributeValueOf(mark, 'author', WML_NAMESPACE_URI) ?? '',
    ...(date === undefined ? {} : { date }),
    ...(id === undefined ? {} : { id }),
  };
}

/** A mark's author and date, for a body mark that follows it. */
export function attributionOf(revision: NoteReferenceRevision): RevisionAttributionInput {
  return {
    author: revision.author,
    ...(revision.date === undefined ? {} : { date: revision.date }),
  };
}

/** The key of a note's references in {@link noteReferenceMarks}: `kind:id`. */
export function noteKey(kind: NoteKind, noteId: number): string {
  return `${kind}:${noteId}`;
}

/** The note kind and id a {@link noteKey} names. */
export function parseNoteKey(key: string): { readonly kind: NoteKind; readonly noteId: number } {
  const [kind, id] = key.split(':') as [NoteKind, string];
  return { kind, noteId: Number(id) };
}

/** The note a {@link noteKey} names, or null for a separator, which is not a note's body. */
export function bodyNoteOf(
  key: string
): { readonly kind: NoteKind; readonly noteId: number } | null {
  const note = parseNoteKey(key);
  return note.noteId > 0 ? note : null;
}

const NO_MARKS: NoteReferenceMarks = { del: null, ins: null };

/** The nearest `w:del` and `w:ins` around a node of a paragraph. */
export function marksAround(
  part: OoxmlPart,
  paragraphId: string,
  nodeId: string
): NoteReferenceMarks {
  const paragraph = findNode(part, paragraphId);
  if (!paragraph || paragraph.kind === 'textValue') return NO_MARKS;

  // A paragraph is small and bounded by the part's own limits: one walk with the path kept
  const stack: { node: OoxmlNode; path: readonly OoxmlNode[] }[] = [{ node: paragraph, path: [] }];
  for (let entry = stack.pop(); entry !== undefined; entry = stack.pop()) {
    const { node, path } = entry;
    if (node.id === nodeId) {
      const nearest = (side: NoteMarkSide): NoteReferenceRevision | null => {
        for (let k = path.length - 1; k >= 0; k -= 1) {
          if (isRevisionMark(path[k]!, side)) return revisionOf(path[k]!);
        }

        return null;
      };
      return { del: nearest('del'), ins: nearest('ins') };
    }

    if (node.kind === 'textValue') continue;
    for (const child of node.children) {
      stack.push({ node: child, path: [...path, node] });
    }
  }

  return NO_MARKS;
}

/**
 * Each note's references outside the notes parts, as `kind:id` to the marks around each
 * reference. Null when the scan ran out of budget.
 */
export function noteReferenceMarks(
  parts: Iterable<OoxmlPart>,
  budget: NoteReferenceScanBudget = createNoteReferenceScanBudget(MAX_NOTE_REFERENCE_MUTATION_SCAN)
): Map<string, NoteReferenceMarks[]> | null {
  const out = new Map<string, NoteReferenceMarks[]>();
  for (const part of parts) {
    if (!part.name.endsWith('.xml') || isNotesPart(part)) continue;
    const hits = collectNoteReferences(part, { maxHits: Number.POSITIVE_INFINITY, budget });
    if (budget.truncated) return null;
    for (const hit of hits) {
      const key = noteKey(hit.noteKind, hit.noteId);
      const list = out.get(key) ?? [];
      list.push(marksAround(part, hit.paragraphId, hit.nodeId));
      out.set(key, list);
    }
  }

  return out;
}

/** Whether two marks have one author and date. */
export const sameRevision = (a: NoteReferenceRevision, b: NoteReferenceRevision): boolean =>
  a.author === b.author && a.date === b.date;

/** The `side` marks of a note's references when every one has one (and there is one); else null. */
export function everyMarked(
  list: readonly NoteReferenceMarks[] | undefined,
  side: NoteMarkSide
): NoteReferenceRevision[] | null {
  if (list === undefined || list.length === 0) return null;
  const out: NoteReferenceRevision[] = [];
  for (const marks of list) {
    const mark = marks[side];
    if (mark === null) return null;
    out.push(mark);
  }

  return out;
}

/** The paragraphs of a note's body, in a table or a control too; not those in a text box. */
export function noteBodyParagraphIds(note: OoxmlNode): string[] {
  const out: string[] = [];
  const stack: OoxmlNode[] = note.kind === 'textValue' ? [] : [...note.children].reverse();
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    if (node.kind === 'paragraph') {
      out.push(node.id);
      continue;
    }

    for (let k = node.children.length - 1; k >= 0; k -= 1) {
      stack.push(node.children[k]!);
    }
  }

  return out;
}
