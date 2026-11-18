// Cheap package-level gates for `TreePackageStore`.

import { findNode } from '../package/ooxml-edit.ts';
import type { OoxmlPackage } from '../package/ooxml-package.ts';
import { resolveRelationship, type RelationshipRecord } from '../package/relationships.ts';
import { resolveNotesPart } from '../package/note-references.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { segmentsOf } from './tree-op-segments.ts';
import type { TreeDocOp } from './tree-ops.ts';
import { hasAnyComment } from './comment-reads.ts';

/** Cap on nodes visited while probing one `deleteBlock` subtree for note refs. */
export const MAX_DELETE_BLOCK_NOTE_GATE_NODES = 4_000;

/** Nesting depth cap for the same probe (nested tables / SDT husks). */
export const MAX_DELETE_BLOCK_NOTE_GATE_DEPTH = 64;

/** Whether the package declares a footnotes or endnotes part at all. */
export function hasNotesPart(pkg: OoxmlPackage): boolean {
  return resolveNotesPart(pkg, 'footnote') !== null || resolveNotesPart(pkg, 'endnote') !== null;
}

/**
 * Whether one text delete can strand a note body and therefore needs a package cascade.
 * Ambiguous repeated deletes in one paragraph fail closed into the bounded cascade.
 */
export function deleteMayStrandNote(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  op: Extract<TreeDocOp, { op: 'deleteText' }>,
  seenParagraphs: Set<string>
): boolean {
  if (!hasNotesPart(pkg)) return false;
  if (seenParagraphs.has(op.paragraphId)) return true;
  seenParagraphs.add(op.paragraphId);
  const paragraph = findNode(part, op.paragraphId);
  if (!paragraph || paragraph.kind !== 'paragraph') return true;
  return segmentsOf(paragraph).some(
    (segment) =>
      segment.node.kind === 'noteReference' && segment.start < op.end && segment.end > op.start
  );
}

/**
 * Whether removing a block subtree can strand a note body.
 *
 * Range deletion of a fully selected table uses `deleteBlock` (no `deleteText` on cell
 * paragraphs), so this gate walks only that subtree — never the whole package — and uses
 * the same segment model as text deletes. Budget overrun or a missing block fails closed.
 */
export function deleteBlockMayStrandNote(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  op: Extract<TreeDocOp, { op: 'deleteBlock' }>,
  seenBlocks: Set<string>
): boolean {
  if (!hasNotesPart(pkg)) return false;
  if (seenBlocks.has(op.blockId)) return true;
  seenBlocks.add(op.blockId);
  const block = findNode(part, op.blockId);
  if (!block || block.kind === 'textValue') return true;

  let visited = 0;
  const walk = (node: OoxmlNode, depth: number): boolean => {
    if (visited >= MAX_DELETE_BLOCK_NOTE_GATE_NODES || depth > MAX_DELETE_BLOCK_NOTE_GATE_DEPTH) {
      return true;
    }
    visited += 1;
    if (node.kind === 'textValue') return false;
    if (node.kind === 'paragraph') {
      return segmentsOf(node).some((segment) => segment.node.kind === 'noteReference');
    }
    for (const child of node.children) {
      if (walk(child, depth + 1)) return true;
    }
    return false;
  };
  return walk(block, 0);
}

/** Whether a subtree holds a comment range marker, within the same budget the note gates use. */
function subtreeHasCommentMarker(
  node: OoxmlNode,
  depth: number,
  visited: { count: number }
): boolean {
  if (
    visited.count >= MAX_DELETE_BLOCK_NOTE_GATE_NODES ||
    depth > MAX_DELETE_BLOCK_NOTE_GATE_DEPTH
  ) {
    return true;
  }
  visited.count += 1;
  if (node.kind === 'textValue') return false;
  if (
    node.kind === 'commentRangeStart' ||
    node.kind === 'commentRangeEnd' ||
    node.kind === 'commentReference'
  ) {
    return true;
  }
  for (const child of node.children) {
    if (subtreeHasCommentMarker(child, depth + 1, visited)) return true;
  }
  return false;
}

/**
 * Whether an op can empty a comment's range and therefore needs the package reap.
 *
 * MARKERS, not offsets. A `deleteText` that removes every character between a comment's start
 * and end never touches the markers themselves — they occupy no offsets — so the only cheap
 * question worth asking is whether the paragraph (or the block subtree) holds any marker at
 * all. Documents with no comments answer no on the first line and pay nothing.
 *
 * Fails CLOSED: a paragraph the store cannot find, or a subtree past the budget, runs the reap
 * rather than skipping it. The reap is itself a diff and does nothing when nothing was emptied.
 */
export function deleteMayEmptyCommentRange(
  anyComment: () => boolean,
  part: OoxmlPart,
  op: Extract<TreeDocOp, { op: 'deleteText' | 'deleteBlock' }>,
  seenTargets: Set<string>
): boolean {
  if (!anyComment()) return false;
  const targetId = op.op === 'deleteText' ? op.paragraphId : op.blockId;
  if (seenTargets.has(targetId)) return true;
  seenTargets.add(targetId);
  const target = findNode(part, targetId);
  if (!target || target.kind === 'textValue') return true;
  return subtreeHasCommentMarker(target, 0, { count: 0 });
}

export const HEADER_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header';
export const FOOTER_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

type LocatedHeaderFooter =
  | { readonly ok: true; readonly partName: string; readonly part: OoxmlPart }
  | {
      readonly ok: false;
      readonly reason:
        | 'dangling-relationship'
        | 'wrong-relationship-type'
        | 'external-relationship'
        | 'bad-relationship-target'
        | 'missing-part'
        | 'not-a-story-part';
      readonly detail?: string;
    };

/** Resolve a header/footer rId to its canonical part, or a typed refusal. */
export function locateHeaderFooterPart(pkg: OoxmlPackage, rId: string): LocatedHeaderFooter {
  const relationships = pkg.relationships.get(pkg.mainDocumentPart) ?? [];
  const record = relationships.find((rel) => rel.id === rId);
  if (!record) {
    return { ok: false, reason: 'dangling-relationship', detail: rId };
  }
  if (record.type !== HEADER_REL_TYPE && record.type !== FOOTER_REL_TYPE) {
    return { ok: false, reason: 'wrong-relationship-type', detail: record.type };
  }
  return resolveInternalStoryPart(pkg, record);
}

function resolveInternalStoryPart(
  pkg: OoxmlPackage,
  record: RelationshipRecord
): LocatedHeaderFooter {
  const resolved = resolveRelationship(record);
  if (resolved.mode === 'External') {
    return { ok: false, reason: 'external-relationship', detail: record.id };
  }
  if (!resolved.target.ok) {
    return {
      ok: false,
      reason: 'bad-relationship-target',
      detail: resolved.target.reason,
    };
  }
  const part = pkg.parts.get(resolved.target.partName);
  if (!part) {
    return { ok: false, reason: 'missing-part', detail: resolved.target.partName };
  }
  const rootName = part.root.localName;
  if (rootName !== 'hdr' && rootName !== 'ftr') {
    return { ok: false, reason: 'not-a-story-part', detail: rootName || part.name };
  }
  return { ok: true, partName: part.name, part };
}

/** Revision ops whose result can remove a note reference along with the content it sits in. */
export const RESOLUTION_OPS: ReadonlySet<string> = new Set([
  'acceptRevision',
  'rejectRevision',
  'acceptAllRevisions',
  'rejectAllRevisions',
]);

/**
 * Ops that remove whole blocks without naming one, so no cheap subtree probe exists.
 *
 * Row and column deletion take a table id and carry away every cell paragraph under it —
 * comment range markers included. Gated by kind rather than by content: the reap they open is
 * a before/after diff and finds nothing when the table held no comment.
 */
export const CONTENT_REMOVING_OPS: ReadonlySet<string> = new Set([
  'authorTable',
  'deleteTableRow',
  'deleteTableColumn',
  'removeContentControl',
  'removeRepeatingSectionItem',
]);

/**
 * What one story transaction's ops say about the cascades its commit may need.
 *
 * `deleteText` / `deleteBlock` can remove noteReference atoms; every other op skips the
 * package-wide note cascade. The gates stay local to the op target (paragraph range or block
 * subtree), so ordinary structural deletion never scans the whole package. The comment gate asks
 * the same kind of question: whether the transaction can leave a comment covering no characters.
 * Word deletes a comment whose words are deleted, and the reap that does it is a before/after
 * diff, so it needs the same "was it even possible" gate.
 */
export class StoryDeletionGates {
  mayDeleteNoteAtoms = false;
  mayEmptyComments = false;
  private anyComment: boolean | undefined;
  private readonly deleteTargets = new Set<string>();
  private readonly commentTargets = new Set<string>();
  /** The package as the transaction found it: one walk answers every op. */
  private readonly pkg: OoxmlPackage;

  constructor(pkg: OoxmlPackage) {
    this.pkg = pkg;
  }

  observe(op: TreeDocOp, part: OoxmlPart): void {
    if (!this.mayDeleteNoteAtoms) {
      if (op.op === 'deleteText' && deleteMayStrandNote(this.pkg, part, op, this.deleteTargets)) {
        this.mayDeleteNoteAtoms = true;
      } else if (
        op.op === 'deleteBlock' &&
        deleteBlockMayStrandNote(this.pkg, part, op, this.deleteTargets)
      ) {
        this.mayDeleteNoteAtoms = true;
      } else if (RESOLUTION_OPS.has(op.op) || CONTENT_REMOVING_OPS.has(op.op)) {
        // Resolving a revision removes the content it covers, and a note reference is one
        // model unit, so a struck-through selection carries it away. Not narrowable to a
        // paragraph range: a revision's sites are wherever the file put them. The cascade
        // is a diff, so it is free.
        this.mayDeleteNoteAtoms = true;
      }
    }
    if (this.mayEmptyComments) return;
    if (op.op === 'deleteText' || op.op === 'deleteBlock') {
      this.mayEmptyComments = deleteMayEmptyCommentRange(
        () => (this.anyComment ??= hasAnyComment(this.pkg)),
        part,
        op,
        this.commentTargets
      );
    } else if (RESOLUTION_OPS.has(op.op) || CONTENT_REMOVING_OPS.has(op.op)) {
      // Rejecting an insertion removes the words it inserted, which a comment can be
      // anchored over. A row or column deletion removes whole cell PARAGRAPHS, markers
      // and all, and names a table rather than a paragraph, so there is no cheap
      // subtree to probe. Opens the gate outright; the reap is a diff, so it is free.
      this.mayEmptyComments = true;
    }
  }
}

/** Ops that can replace a content control's prompt with what they insert, in either mode. */
const PROMPT_REPLACING_OPS: ReadonlySet<string> = new Set([
  'insertText',
  'insertFragment',
  'insertTab',
  'insertHardBreak',
  'insertPageBreak',
  'insertPageField',
]);

/**
 * Whether an op can keep a content control's prompt in the glossary part: an insertion over the
 * prompt, direct or suggested, or a suggested value (`content-control-prompt-glossary.ts`).
 */
export function mayRetainPrompt(op: TreeDocOp): boolean {
  return (
    PROMPT_REPLACING_OPS.has(op.op) ||
    (op.op === 'setContentControlValue' && op.revision !== undefined)
  );
}
