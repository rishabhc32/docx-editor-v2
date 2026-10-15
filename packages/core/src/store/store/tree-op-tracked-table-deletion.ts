// A TRACKED table deletion — `deleteBlock` on a `w:tbl`, with revision attribution.
//
// Word records a whole table deleted under track changes as every row deleted: `w:trPr/w:del`
// with a `w:cellDel` in each cell, the form a tracked `deleteTableRow` writes for one row. The
// table stays until it is reviewed. Accepting removes every row and the table with them
// (`tableRevisionRemovals`); rejecting drops the markers and the table is back.
//
// ONE op rather than a tracked `deleteTableRow` per row, because that op refuses the last row
// without a pending deletion: pending deletions cannot keep a table alive for another row
// suggestion, so a row-by-row proposal of the whole table was refused at its last row.
//
// A row this author proposed inserting is WITHDRAWN instead — the same edit as rejecting it —
// rather than proposed for deletion, as a tracked column deletion withdraws a proposed cell.

import {
  applyEdits,
  createNodeIdAllocator,
  findNode,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { wmlFreshNamespaceContextAt } from '../package/wml-namespace.ts';
import { tableChildren } from './revision-table-children.ts';
import { revisionAttribute } from './revision-table-plan.ts';
import { paragraphIdsWithin, survivingCaretAfterBlockRemoval } from './tree-op-blocks.ts';
import { fromEdit, TEXT_DEPS } from './tree-op-nodes.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { resolveRevisions } from './tree-op-revisions.ts';
import { wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';
import { readEditableTableTopology, type EditableTableTopology } from './tree-op-table-topology.ts';
import { withTrackedRowMarker } from './tree-op-tables.ts';
import type { TreeDocOp, TreeOpEffect, TreeOpRejection, TreeOpResult } from './tree-op-types.ts';

type DeleteBlockOp = Extract<TreeDocOp, { op: 'deleteBlock' }>;

/** The row's own `w:ins` or `w:del` marker, when it carries one. */
function rowMarker(row: OoxmlElement, localName: 'ins' | 'del'): OoxmlElement | undefined {
  const properties = wmlChildNamed(row, 'trPr');
  return properties && wmlChildNamed(properties, localName);
}

/** Whether any revision under `node` was written by someone other than `author`. */
export function holdsOtherAuthorsRevision(node: OoxmlNode, author: string): boolean {
  if (node.kind === 'textValue') return false;
  const written =
    node.namespaceUri === WML_NAMESPACE_URI ? revisionAttribute(node, 'author') : undefined;
  if (written !== undefined && written !== author) return true;
  return node.children.some((child) => holdsOtherAuthorsRevision(child, author));
}

/** Whether `author` proposed inserting the row. */
function isOwnInsertedRow(row: OoxmlElement, author: string): boolean {
  const inserted = rowMarker(row, 'ins');
  return inserted !== undefined && wmlAttributeValue(inserted, 'author') === author;
}

/**
 * Whether the deletion WITHDRAWS this author's own proposal instead of making one.
 *
 * A table whose every row is the author's pending insertion — a table they proposed — never
 * existed for anyone else. Deleting it removes it outright, as Word does, rather than
 * proposing to delete a proposal.
 */
function withdrawsOwnTable(topology: EditableTableTopology, author: string): boolean {
  return topology.rows.every(({ row }) => isOwnInsertedRow(row, author));
}

/** Whether `op` proposes a table deletion, as opposed to removing its block outright. */
export function proposesTableDeletion(part: OoxmlPart, op: DeleteBlockOp): boolean {
  if (op.revision === undefined) return false;
  const result = readEditableTableTopology(part.root, op.blockId);
  return !result.ok || !withdrawsOwnTable(result.topology, op.revision.author);
}

/**
 * Why a tracked `deleteBlock` is refused, or null when it may proceed. Runs after the op's
 * untracked form passed, which is the whole rule of an untracked `deleteBlock`.
 */
export function validateTrackedTableDeletion(
  part: OoxmlPart,
  op: DeleteBlockOp,
  revision: RevisionAttributionInput
): TreeOpRejection | null {
  // Only a table has a tracked deletion of its own; a paragraph's is struck text and mark.
  if (findNode(part, op.blockId)?.kind !== 'table') return 'unsupported-revision';
  const result = readEditableTableTopology(part.root, op.blockId);
  if (!result.ok) return result.reason === 'duplicate-node-id' ? 'unknown-table' : result.reason;
  const { topology } = result;
  // A row inside a content control is not one the topology lists, so it would be left
  // standing by an Accept that took every other row.
  if (tableChildren(topology.table, 'tableRow').length !== topology.rows.length) {
    return 'unsupported-revision';
  }
  if (withdrawsOwnTable(topology, revision.author)) return null;
  for (const { row, cells } of topology.rows) {
    // Already proposed for deletion: that decision stands and needs no second one.
    if (rowMarker(row, 'del')) continue;
    // Withdrawing removes the row, so it must not take someone else's pending change along.
    if (isOwnInsertedRow(row, revision.author)) {
      if (holdsOtherAuthorsRevision(row, revision.author)) return 'unsupported-revision';
      continue;
    }
    // Otherwise the subset a tracked `deleteTableRow` supports: no pending row or cell revision.
    if (rowMarker(row, 'ins')) return 'unsupported-revision';
    for (const cell of cells) {
      const tcPr = wmlChildNamed(cell, 'tcPr');
      if (
        tcPr &&
        (wmlChildNamed(tcPr, 'cellIns') ||
          wmlChildNamed(tcPr, 'cellDel') ||
          wmlChildNamed(tcPr, 'cellMerge'))
      )
        return 'unsupported-revision';
    }
  }
  return null;
}

/** Mark every row of the table deleted. Runs after `validateTrackedTableDeletion` admitted it. */
export function applyTrackedTableDeletion(
  part: OoxmlPart,
  op: DeleteBlockOp,
  options?: EditOptions
): TreeOpResult {
  const result = readEditableTableTopology(part.root, op.blockId);
  if (!result.ok || op.revision === undefined) return { ok: false, reason: 'unknown-table' };
  const { topology } = result;
  // The caret leaves the struck table for the paragraph an outright removal would give it.
  const caretParagraphId = survivingCaretAfterBlockRemoval(part, op.blockId);
  if (!caretParagraphId) return { ok: false, reason: 'block-required' };
  // ONE id for every row, as a tracked `insertTable` writes: one gesture is one decision.
  const revisionId = options?.revisionIds?.() ?? nextRevisionId(part)();
  const nextId = createNodeIdAllocator(part);
  const wml = wmlFreshNamespaceContextAt(part, topology.table);
  const revision = op.revision;
  const withdrawn = topology.rows
    .map(({ row }) => row)
    .filter((row) => isOwnInsertedRow(row, revision.author));
  const marked = topology.rows
    .filter(({ row }) => !rowMarker(row, 'del') && !withdrawn.includes(row))
    .map(({ row }) => withTrackedRowMarker(row, 'del', revisionId, revision, nextId, wml));
  const effect: TreeOpEffect = {
    dirty: [topology.table.id, ...marked.map((row) => row.id)],
    created: [],
    deleted: withdrawn.flatMap((row) => [row.id, ...paragraphIdsWithin(row)]),
    dependencyKeys: TEXT_DEPS,
    impact: 'flow-structural',
    caret: { paragraphId: caretParagraphId },
  };
  const edits = marked.map(
    (row) => (current: OoxmlPart) => replaceNode(current, row.id, row, options)
  );
  const proposed = fromEdit(applyEdits(part, edits, options), effect);
  if (!proposed.ok || withdrawn.length === 0) return proposed;
  const withdrawal = resolveRevisions(proposed.part, 'reject', undefined, {
    ...options,
    siteNodeIds: withdrawn.map((row) => rowMarker(row, 'ins')!.id),
  });
  if (!withdrawal.ok || !withdrawal.part) {
    return { ok: false, reason: withdrawal.reason ?? 'unsupported-revision' };
  }
  return { ok: true, part: withdrawal.part, effect };
}
