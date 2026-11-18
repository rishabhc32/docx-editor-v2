// The paginated surface's internal members, which other editor modules reach by cast.

import type { TreeApplyResult } from '@docx-editor.dev/core/binding';
import type { SelectionMark } from '@docx-editor.dev/core/store';
import type { CustomNodePackageEdit } from '../store/store/custom-node-write-plan.ts';
import type { StoryScope } from '../store/store/tree-package-store.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { PaginatedSurface } from './paginated-surface-contract.ts';

/** The surface with the internal members that other editor modules reach by cast. */
export type ScaleMutableSurface = PaginatedSurface & {
  setScale(nextScale: number): boolean;
  /**
   * Why the mode rules would refuse these automation ops now, or null: the check
   * `applyAutomationOps` makes, without writing. A preview fork asks it so it refuses what a
   * commit would.
   */
  automationWriteRefusal(ops: readonly TreeDocOp[], packageEdits: boolean): string | null;
  /**
   * The gated, attributed session write for command lanes that live outside
   * `paginated-surface.ts`, such as the content-control commands. Same collaboration gate,
   * actor and revision attribution as the typing lane (`trackedOps`, `applyJournaledOps`); the
   * command's own mode and lock gate has already run. Internal, reached by cast like
   * `setScale`; see `applyGatedSurfaceTreeOps` in `content-controls.ts` for the accessor.
   */
  applyGatedTreeOps(
    ops: readonly TreeDocOp[],
    selectionBefore?: SelectionMark | null,
    selectionAfter?: SelectionMark | null,
    scope?: StoryScope
  ): TreeApplyResult;
  /**
   * A write planned with its package edits (a custom node's store and its bound control),
   * committed as one command: the mode rules, then the collaboration gate with the package the
   * ops meet, then one transaction. The ops are written as given, never tracked. Internal, reached
   * by cast; see `commitCustomNodeInsert` in `custom-node-commit.ts`.
   */
  applyGatedPackageWrite(
    ops: readonly TreeDocOp[],
    scope: StoryScope,
    packageEdits: readonly CustomNodePackageEdit[]
  ): TreeApplyResult;
};
