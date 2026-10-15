// The opt-in strict half of suggesting mode: refuse a local edit it cannot record.
//
// Suggesting attributes every op it has a tracked form for (`attributeTrackedOps`) and lets the
// rest through as they are — a page-setup change, a list toggle, a table property — so the
// document changes with no card for a reviewer to accept or reject. A review module created with
// `untrackedEdits: 'refuse'` turns that lane off: the surface refuses the WHOLE batch instead.
//
// Kept off the composition root so `paginated-surface.ts` stays under its line cap.

import { applyTreeOp, type TreeDocOp } from '../store/store/tree-ops.ts';
import { paragraphMarkRevisionOf } from '../store/store/tree-op-tracked-marks.ts';
import { PROPERTY_CHANGE_WRAPPER_OF_OP } from '../store/store/tree-op-tracked-properties.ts';
import { insertionAuthor } from '../store/store/tree-op-retraction.ts';
import { nextRevisionId } from '../store/store/tree-op-revision-ids.ts';
import { paragraphOffsetIndex, type OffsetSpan } from '../store/store/tree-op-segments.ts';
import { hyperlinkAnchorOf, hyperlinkRelationshipIdOf } from '../store/package/hyperlink.ts';
import {
  findNode,
  parentNodeOf,
  type TransactionRevisionIds,
} from '../store/package/ooxml-edit.ts';
import { WML_NAMESPACE_URI } from '../store/package/ooxml-shared.ts';
import type { OoxmlNode, OoxmlPart } from '../store/package/ooxml-tree.ts';
import { isDocumentProtectionBatch } from './surface-hf-editing.ts';
import { isTrackedEdit } from './surface-tracked-ops.ts';

/** Why suggesting refused an edit that would have changed the document with no suggestion. */
export const UNTRACKED_EDIT_REFUSAL = 'suggesting cannot record this change as a suggestion';

/** The live part to look a node up in, or null. Read for the ops judged by the link they name. */
export type PartOfNode = (nodeId: string) => OoxmlPart | null;

/** Ops that change the DOCUMENT, as opposed to reading or resolving it. */
export function isDocumentEdit(op: TreeDocOp): boolean {
  return (
    op.op !== 'acceptRevision' &&
    op.op !== 'rejectRevision' &&
    op.op !== 'acceptAllRevisions' &&
    op.op !== 'rejectAllRevisions' &&
    op.op !== 'setRevisionAttribution'
  );
}

/** Whether the batch proposes this paragraph's mark as inserted. */
function proposesMark(batch: readonly TreeDocOp[], paragraphId: string): boolean {
  return batch.some(
    (other) =>
      other.op === 'setParagraphMarkRevision' &&
      other.kind === 'ins' &&
      other.paragraphId === paragraphId
  );
}

/** What the rule reads beside each op: the whole batch, the document, and its author. */
interface Judging {
  readonly batch: readonly TreeDocOp[];
  readonly formattingTracked: boolean;
  readonly partOf: PartOfNode;
  /** The author suggesting attributes the batch to; without one, no insertion is its own. */
  readonly author: string | undefined;
}

/**
 * The ops a link lane writes around a copy of its words, the only ones replayed to find it.
 * True for the two the store hands its transaction's revision bookkeeping, as it does.
 */
const LINK_COPY_OPS: ReadonlyMap<TreeDocOp['op'], boolean> = new Map([
  ['deleteText', true],
  ['insertText', true],
  ['setRunProperties', false],
  ['removeHyperlink', false],
  ['insertHyperlink', false],
  ['setHyperlinkTarget', false],
]);

/**
 * The part holding `nodeId` as the batch's ops before `index` leave it, with the revision ids
 * they minted. Null when that cannot be told: an op that is not a link lane's, or one the store
 * refuses.
 *
 * Asked of the store's own appliers, with the bookkeeping a transaction keeps, so an insertion
 * aimed into words the batch struck lands past them exactly as it will on commit.
 */
function replayBefore(
  context: Judging,
  index: number,
  nodeId: string
): { readonly part: OoxmlPart; readonly minted: ReadonlySet<string> } | null {
  const start = context.partOf(nodeId);
  if (!start) return null;
  const minted = new Set<string>();
  const written = new Set<string>();
  let next: (() => string) | null = null;
  const trackedRevisionIds: TransactionRevisionIds = {
    mint: () => {
      const id = (next ??= nextRevisionId(start))();
      minted.add(id);
      return id;
    },
    wrote: (key) => void written.add(key),
    wroteUnder: (key) => written.has(key),
  };
  let part = start;
  for (const op of context.batch.slice(0, index)) {
    const tracked = LINK_COPY_OPS.get(op.op);
    if (tracked === undefined) return null;
    const result = applyTreeOp(part, op, {
      deferValidation: true,
      ...(tracked ? { trackedRevisionIds } : {}),
    });
    if (!result.ok) return null;
    part = result.part;
  }
  return { part, minted };
}

/** A revision wrapper's `@w:id`, or undefined. */
function revisionIdOf(node: OoxmlNode): string | undefined {
  if (node.kind === 'textValue') return undefined;
  return node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
}

/**
 * Whether `[start, end)` of a paragraph is text that a tracked insertion earlier in the batch
 * wrote, and nothing struck, once the ops before `index` apply. `paragraphOf` finds the
 * paragraph in that part; `nodeId` names a node of the part.
 */
function holdsBatchInsertion(
  context: Judging,
  index: number,
  nodeId: string,
  paragraphOf: (part: OoxmlPart) => OoxmlNode | null,
  start: number,
  end: number
): boolean {
  const replay = context.author === undefined ? null : replayBefore(context, index, nodeId);
  const paragraph = replay ? paragraphOf(replay.part) : null;
  if (!replay || paragraph?.kind !== 'paragraph' || start >= end) return false;
  const offsets = paragraphOffsetIndex(paragraph);
  const inserted: OffsetSpan[] = [];
  let struck = false;
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'revisionDelete' || node.kind === 'revisionMoveFrom') {
      const span = offsets.spanOf(node);
      // A strike the offset walk cannot place might be anywhere.
      if (!span || (span.start < end && start < span.end)) struck = true;
      return;
    }
    const id = revisionIdOf(node);
    if (
      node.kind === 'revisionInsert' &&
      id !== undefined &&
      replay.minted.has(id) &&
      insertionAuthor([node]) === context.author
    ) {
      const span = offsets.spanOf(node);
      if (span) inserted.push(span);
    }
    for (const child of node.children) visit(child);
  };
  for (const child of paragraph.children) visit(child);
  inserted.sort((a, b) => a.start - b.start);
  let covered = start;
  for (const span of inserted) {
    if (span.start > covered) break;
    covered = Math.max(covered, span.end);
  }
  return !struck && covered >= end;
}

/**
 * Whether re-aiming a whole link is no untracked edit. Either the link holds nothing but the
 * author's own pending insertion, whose rejection takes the link and its target with it, or the
 * re-aim leaves the link exactly as it is: the same `r:id` and `w:anchor`, and the same tooltip
 * when the op names one.
 */
function reaimRecorded(
  op: Extract<TreeDocOp, { op: 'setHyperlinkTarget' }>,
  partOf: PartOfNode,
  author: string | undefined
): boolean {
  const part = op.range ? null : partOf(op.linkId);
  const link = part ? findNode(part, op.linkId) : null;
  if (link?.kind !== 'hyperlink') return false;
  if (heldByOwnInsertion(link, author)) return true;
  const tooltip = link.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'tooltip'
  )?.value;
  return (
    hyperlinkRelationshipIdOf(link) === op.relationshipId &&
    hyperlinkAnchorOf(link) === op.anchor &&
    (op.tooltip === undefined || op.tooltip === tooltip)
  );
}

/**
 * Whether a link holds nothing but `author`'s own pending insertion, whose rejection takes the
 * link, its target and its words with it.
 */
function heldByOwnInsertion(link: OoxmlNode, author: string | undefined): boolean {
  return (
    author !== undefined &&
    link.kind === 'hyperlink' &&
    link.children.length > 0 &&
    link.children.every(
      (child) => child.kind === 'revisionInsert' && insertionAuthor([child]) === author
    )
  );
}

/**
 * Whether `op`, as it will be WRITTEN, shows up as a suggestion — or changes nothing a reviewer
 * decides on.
 *
 * Read after attribution, so it judges exactly what the store receives. Beyond the ops that carry
 * their own record (`isTrackedEdit`), five ride on a record elsewhere in the SAME batch, because
 * the lanes that write them put them there: a split is recorded by the inserted mark attribution
 * adds for it, a section break's `w:sectPr` lives on that mark and leaves with it, a list's
 * blank-separator Enter proposes every mark its fragment adds, and a new link wraps, or a removed
 * link lets go of, only the tracked insertion the link lane wrote for its text, where the store
 * lands it. A re-aimed or unwrapped link rides on a record already in the document: the author's
 * own pending insertion.
 */
function recorded(op: TreeDocOp, index: number, context: Judging): boolean {
  const { batch } = context;
  if (isTrackedEdit(op) || !isDocumentEdit(op)) return true;
  // `trackFormatting: false` or `w:doNotTrackFormatting`: the reader or the document chose to
  // apply formatting untracked, and that choice outranks this one.
  if (!context.formattingTracked && PROPERTY_CHANGE_WRAPPER_OF_OP.has(op.op)) return true;
  switch (op.op) {
    case 'insertCommentMarker':
    // Bookmarks are not revisions: their markers carry no author, and the editing API writes
    // them while it tracks changes too.
    case 'insertBookmark':
    case 'deleteBookmark':
      return true;
    case 'splitParagraph':
    case 'setSectionMark':
      return proposesMark(batch, op.paragraphId);
    case 'insertFragment':
      // Empty paragraphs only. The first merges into the paragraph the fragment lands in, whose
      // mark the batch proposes; each later one carries its own inserted mark.
      return op.blocks.every(
        (block, index) =>
          block.kind === 'paragraph' &&
          block.children.every((child) => child.kind === 'paragraphProperties') &&
          (index === 0
            ? proposesMark(batch, op.paragraphId)
            : paragraphMarkRevisionOf(block, 'ins') !== undefined)
      );
    case 'insertHyperlink':
      return holdsBatchInsertion(
        context,
        index,
        op.paragraphId,
        (part) => findNode(part, op.paragraphId),
        op.start,
        op.end
      );
    case 'removeHyperlink': {
      if (op.range !== undefined) {
        return holdsBatchInsertion(
          context,
          index,
          op.linkId,
          (part) => parentNodeOf(part, op.linkId),
          op.range.start,
          op.range.end
        );
      }
      // A link lane unwraps the author's own pending link in place: rejecting the insertion
      // still takes its words.
      const replay = replayBefore(context, index, op.linkId);
      const link = replay ? findNode(replay.part, op.linkId) : null;
      return link !== null && heldByOwnInsertion(link, context.author);
    }
    case 'setHyperlinkTarget':
      // A link lane re-aims the author's own pending link in place, and names the target it
      // keeps beside a suggested rename of a link's text.
      return reaimRecorded(op, context.partOf, context.author);
    default:
      return false;
  }
}

/**
 * The first op of an attributed batch that would change the document with no suggestion, or null.
 * `author` is the one suggesting attributes the batch to; without it no pending insertion is
 * this batch's own.
 *
 * Protection is the one document write left alone, as every mode rule leaves it: it sets who may
 * edit, not what the document says. A batch with no ops has nothing to judge: the bookmark lane
 * commits one to drop a name's copies from other stories, which is no revision either.
 */
export function untrackedEditOf(
  ops: readonly TreeDocOp[],
  formattingTracked: boolean,
  partOf: PartOfNode,
  author?: string
): TreeDocOp | null {
  if (isDocumentProtectionBatch(ops)) return null;
  const context: Judging = { batch: ops, formattingTracked, partOf, author };
  return ops.find((op, index) => !recorded(op, index, context)) ?? null;
}

/**
 * The surface's gate for the policy: asked with every batch a local lane commits.
 *
 * The automation lane runs `exempt`: it keeps its own tracked-change policy, which already
 * refuses every operation with no tracked form while the surface is suggesting, and which
 * knows about an author's own pending table and control insertions that this op-level rule
 * does not.
 */
export function createUntrackedEditGate(deps: {
  readonly policy: 'apply' | 'refuse' | undefined;
  readonly suggesting: () => boolean;
  readonly formattingTracked: () => boolean;
  readonly partOf: PartOfNode;
  /** The author suggesting attributes edits to, or undefined when it attributes none. */
  readonly author: () => string | undefined;
}): {
  readonly refusal: (ops: readonly TreeDocOp[]) => string | null;
  /**
   * Whether the policy is in force now. A lane that only writes untracked asks this and skips
   * its write: the save-time REF field refresh, because a field result has no tracked form.
   */
  readonly inForce: () => boolean;
  readonly exempt: <T>(run: () => T) => T;
} {
  let exempt = 0;
  const inForce = () => deps.policy === 'refuse' && exempt === 0 && deps.suggesting();
  return {
    refusal: (ops) =>
      inForce() && untrackedEditOf(ops, deps.formattingTracked(), deps.partOf, deps.author())
        ? UNTRACKED_EDIT_REFUSAL
        : null,
    inForce,
    exempt(run) {
      exempt += 1;
      try {
        return run();
      } finally {
        exempt -= 1;
      }
    },
  };
}
