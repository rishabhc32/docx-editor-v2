// The ops the surface stamps in suggesting mode, pinned to the ops it reports as tracked.
//
// `isRevisionCapable` picks the ops `attributeTrackedOps` stamps with the suggesting author.
// `isTrackedEdit` picks the ops that count as a proposal. An op the second knows and the first
// does not lands untracked in suggesting mode unless its producer attributes it: the #458
// failure for tabs and breaks, and again for the list buttons.
//
// So every op that can carry an attribution is stamped, or listed below with the reason it is
// not. Two gates hold the list. At run time every kind in `TREE_DOC_OP_KINDS` is probed through
// both predicates. At compile time the list must name exactly the attributable kinds the
// stamping set leaves out, so a new op with a `revision` fails `bun check` until it is listed.

import { describe, expect, test } from 'bun:test';
import { isRevisionCapable, isTrackedEdit } from '../editor/surface-tracked-ops.ts';
import type { RevisionAttributionInput } from '../store/store/tree-op-revision-attribution.ts';
import {
  TREE_DOC_OP_KINDS,
  type TreeDocOp,
  type TreeDocOpKind,
} from '../store/store/tree-op-types.ts';

/** Kinds whose `revision` takes an attribution, not the address accept and reject take. */
type AttributableKind<O = TreeDocOp> = O extends {
  readonly op: infer K;
  readonly revision?: infer R;
}
  ? 'revision' extends keyof O
    ? RevisionAttributionInput extends R
      ? K
      : never
    : never
  : never;

/** The kinds `attributeTrackedOps` stamps, read off the predicate that picks them. */
type StampedKind = (typeof isRevisionCapable extends ((
  op: TreeDocOp
) => op is infer T extends TreeDocOp)
  ? T
  : never)['op'];

/**
 * Every attributable op the surface does not stamp. `tracked` is what `isTrackedEdit` answers
 * for one that carries a revision.
 */
const NOT_STAMPED = {
  setParagraphMarkRevision: {
    tracked: true,
    reason:
      'Only a proposal, so its revision is required. attributeTrackedOps emits it already ' +
      'attributed for a split, and so does Enter when it inserts a list separator.',
  },
  proposeParagraphMerge: {
    tracked: true,
    reason:
      'Only a proposal, so its revision is required. attributeTrackedOps emits it already ' +
      'attributed in place of joinParagraphs.',
  },
  splitParagraphMany: {
    tracked: true,
    reason:
      'attributeTrackedOps stamps it in a branch of its own beside splitParagraph: paste ' +
      'splits with one op, and its revision proposes every mark the store mints for it.',
  },
  insertContentControl: {
    tracked: true,
    reason:
      'A revision turns the wrap into a replacement, which the store refuses unless the range ' +
      'is nonempty ordinary text (canTrackContentControl). Automation checks the range and ' +
      'attributes the op itself; stamping every editor insert would refuse one at a caret.',
  },
  authorTable: {
    tracked: true,
    reason:
      'Only automation emits it (plan-tables.ts). The planner attributes it when tracking and ' +
      'refuses table mutations that have no tracked form.',
  },
  insertDrawing: {
    tracked: false,
    reason:
      'Suggesting mode refuses drawing ops from the editor (applyDrawingOps) and automation. ' +
      'A suggested picture goes through insertImage, which attributes this op itself.',
  },
  deleteDrawing: {
    tracked: false,
    reason:
      'The store refuses its revision (trackedDrawingDeletionUnsupported). A suggested ' +
      'picture deletion is a tracked deleteText over the drawing (deleteImageTracked).',
  },
  setRevisionAttribution: {
    tracked: false,
    reason:
      'Its revision names the new author of changes already pending. It proposes nothing, ' +
      'and the surface treats it as review authoring, not a document edit (isDocumentEdit).',
  },
} as const satisfies Record<
  Exclude<AttributableKind, StampedKind>,
  { readonly tracked: boolean; readonly reason: string }
>;

/** An op of `kind` with an author. Both predicates read only `op` and `revision`. */
function attributed(kind: TreeDocOpKind): TreeDocOp {
  return { op: kind, revision: { author: 'Ada' } } as TreeDocOp;
}

describe('revision-capable ops', () => {
  test('every op reported as tracked is stamped, or listed with its reason', () => {
    const unstamped = Object.fromEntries(
      TREE_DOC_OP_KINDS.flatMap((kind) => {
        const op = attributed(kind);
        if (isRevisionCapable(op)) return [];
        const tracked = isTrackedEdit(op);
        return tracked || Object.hasOwn(NOT_STAMPED, kind) ? [[kind, tracked]] : [];
      })
    );
    expect(unstamped).toEqual(
      Object.fromEntries(Object.entries(NOT_STAMPED).map(([kind, { tracked }]) => [kind, tracked]))
    );
  });
});
