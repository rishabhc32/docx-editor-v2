// A custom node written through the editor's gated lane: planned against the package as it is
// now (`custom-node-write-plan.ts`), then committed as one command with its package edits.

import type { StoryScope } from '../store/store/tree-package-store.ts';
import {
  authoredControlId,
  contentControlIdsIn,
  type InsertCustomNodeWrite,
} from '../store/store/custom-node-writes.ts';
import {
  planCustomNodeInsert,
  planCustomNodeRemoval,
  type CustomNodeWritePlan,
} from '../store/store/custom-node-write-plan.ts';
import type { PaginatedSurface, ScaleMutableSurface } from './paginated-surface-contract.ts';

/**
 * A custom-node write through the editor: the control it authored, or why it was refused. A
 * refusal's reason is the plan's (`CustomNodeWriteRejection`), the store's, or the editor's own
 * (its mode, a collaboration session).
 *
 * @internal
 */
export type CustomNodeCommitResult =
  | { readonly ok: true; readonly changed: boolean; readonly nodeId?: string }
  | { readonly ok: false; readonly reason: string; readonly detail?: string };

/**
 * Commit a plan in one story as one command, and name the control it authored. The
 * collaboration gate sees the control's ops with the package that holds its store and node, so
 * a room carries the node with the control.
 */
function commitPlan(
  surface: PaginatedSurface,
  plan: CustomNodeWritePlan,
  scope: StoryScope
): CustomNodeCommitResult {
  const before = surface.session.partFor(scope);
  const controlsBefore = before ? contentControlIdsIn(before) : new Set<string>();
  const result = (surface as ScaleMutableSurface).applyGatedPackageWrite(
    plan.ops,
    scope,
    plan.packageEdits
  );
  if (!result.committed) return { ok: false, reason: String(result.reason ?? 'rejected') };
  const after = surface.session.partFor(scope);
  const nodeId = after ? authoredControlId(after, controlsBefore) : undefined;
  return { ok: true, changed: true, ...(nodeId === undefined ? {} : { nodeId }) };
}

/**
 * Insert (or rewrite, with `replaceControlId`) one custom node in `scope`. Its store hangs off
 * the main document part, which is where Word reads data stores from.
 *
 * @internal
 */
export function commitCustomNodeInsert(
  surface: PaginatedSurface,
  write: InsertCustomNodeWrite,
  scope: StoryScope
): CustomNodeCommitResult {
  const story = surface.session.partFor(scope);
  if (!story) return { ok: false, reason: 'unknown-story' };
  const pkg = surface.session.currentPackage();
  const planned = planCustomNodeInsert(pkg, story, write, pkg.mainDocumentPart);
  if (!planned.ok) return planned;
  return commitPlan(surface, planned.plan, scope);
}

/**
 * Remove one custom node from `scope`, with the payload it bound.
 *
 * @internal
 */
export function commitCustomNodeRemoval(
  surface: PaginatedSurface,
  controlNodeId: string,
  scope: StoryScope
): CustomNodeCommitResult {
  const story = surface.session.partFor(scope);
  if (!story) return { ok: false, reason: 'unknown-story' };
  const plan = planCustomNodeRemoval(surface.session.currentPackage(), story, controlNodeId);
  return commitPlan(surface, plan, scope);
}
