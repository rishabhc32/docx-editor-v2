import { retainedNestedRowSites } from './revision-table-preserve-nested.ts';
import { unboundTableHistories } from './revision-table-unbound-history.ts';
import { deferredTableGridSites } from './revision-table-grid-history.ts';
import { ordinaryMoveRanges, planOrdinaryMoves } from './revision-move-ranges.ts';
import {
  tableRevisionContentTarget,
  tableRevisionRemovals,
  tableRevisionNeighbours,
  tableMergeDependencies,
} from './revision-table-plan.ts';
import { paragraphMergeSources } from './revision-paragraph-merge.ts';
import { containedOwnChangeSites } from './revision-contained-changes.ts';
import { preservedRevisionProperty } from './revision-property-records.ts';
import { recordedProperties } from './tree-op-tracked-properties.ts';
import { findNode, parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { revisionItemsOf } from './review-reads.ts';
import { reviewItemKey, revisionSiteNodeIdsOf, type ReviewRevisionItem } from './review-items.ts';
import {
  collectRevisionSites,
  resolveRevisions,
  namedMoveRanges,
  orphanMoveDestinationSites,
} from './tree-op-revisions.ts';
import type { RevisionSiteRefusal } from './revision-site-refusal.ts';
import type { TreeDocOp } from './tree-op-types.ts';

/** A canonical review decision. Keys are valid only in the open document. @public */
export interface RevisionBatchEntry {
  readonly key: string;
  readonly author: string;
  readonly partName: string;
  readonly revisionKind: ReviewRevisionItem['revisionKind'];
  /** For a formatting change, its OOXML property-change element, as the review item has it. */
  readonly formattingKind?: string;
}

/** Why a requested decision remains pending. @public */
export type RevisionBatchSkipReason =
  | 'unsupported-revision'
  | 'incomplete-group'
  | 'retained-structure'
  | 'unknown-revision';

/**
 * Why an `unsupported-revision` skip happened: a refused site, a formatting change whose earlier
 * properties were not recorded so it cannot be rejected, a read-only review item, or a change
 * with no sites in the part.
 *
 * @public
 */
export type RevisionBatchSkipDetail =
  | RevisionSiteRefusal
  | 'unrecorded-properties'
  | 'read-only'
  | 'no-sites';

/** Outcome of one selected-set decision. Counts refer to review decisions, not XML markers. @public */
export interface RevisionBatchResult {
  readonly resolved: readonly RevisionBatchEntry[];
  readonly skipped: readonly {
    readonly key: string;
    readonly reason: RevisionBatchSkipReason;
    readonly revision?: RevisionBatchEntry;
    /** For `unsupported-revision`, which engine condition refused it. */
    readonly detail?: RevisionBatchSkipDetail;
  }[];
  /** All pending decisions in the addressed scope, including excluded and unsupported changes. */
  readonly remaining: number;
}

/**
 * Whether one review item's decision is planned with its dependencies (`planRevisionBatch`): a
 * grouped row or column decision can remove later sites, so does a move, and a format record over
 * several sites resolves together. Any other item resolves its own sites.
 */
export function decidedAsBatch(item: ReviewRevisionItem, part: OoxmlPart): boolean {
  return (
    item.revisionKind === 'structural' ||
    ordinaryMoveRanges(part.root).length > 0 ||
    (item.revisionKind === 'format' && revisionSiteNodeIdsOf(item).length > 1)
  );
}

/** @internal Pure preflight; callers commit its operations through their normal write guards. */
export function planRevisionBatch(
  part: OoxmlPart,
  action: 'accept' | 'reject',
  keys?: readonly string[],
  scopeRoot?: OoxmlNode
): { ops: readonly TreeDocOp[]; result: RevisionBatchResult } {
  const root = scopeRoot ?? part.root;
  const scopedPart = root.kind === 'textValue' ? part : { ...part, root };
  const items = revisionItemsOf(scopedPart);
  const auxiliaryIds =
    keys === undefined
      ? [...unboundTableHistories(scopedPart, collectRevisionSites(scopedPart))]
      : [];
  const cleanup: TreeDocOp[] = auxiliaryIds.length
    ? [
        {
          op: 'acceptAllRevisions',
          siteNodeIds: auxiliaryIds,
          ...(scopeRoot ? { scopeRootId: scopeRoot.id } : {}),
        },
      ]
    : [];
  if (keys?.length === 0 || items.length === 0) {
    return {
      ops: cleanup,
      result: {
        resolved: [],
        skipped: [...new Set(keys ?? [])].map((key) => ({ key, reason: 'unknown-revision' })),
        remaining: items.length,
      },
    };
  }
  const selected = new Set(keys ?? items.map(reviewItemKey));
  const byKey = new Map(items.map((item) => [reviewItemKey(item), item]));
  const selectedSites = new Set<string>();
  for (const key of selected) {
    const item = byKey.get(key);
    if (item) for (const id of revisionSiteNodeIdsOf(item)) selectedSites.add(id);
  }
  const sites = collectRevisionSites(scopedPart);
  const movePlan = planOrdinaryMoves(root, sites, selectedSites, action);
  for (const id of movePlan.selected) selectedSites.add(id);
  // The author's own changes inside what the decision removes go with it; their other sites
  // stay their own decision (`contained` below)
  const contained = new Set(containedOwnChangeSites(part, sites, selectedSites, action));
  for (const id of contained) selectedSites.add(id);
  for (const item of items) {
    const ids = revisionSiteNodeIdsOf(item);
    if (ids.length && ids.every((id) => selectedSites.has(id))) selected.add(reviewItemKey(item));
  }
  const orphanDestinations = orphanMoveDestinationSites(root, sites);
  const indices = new Map(sites.map((site, index) => [site.node.id, index]));
  const parents = sites.map((_, index) => index);
  const find = (index: number): number => {
    while (parents[index] !== index) {
      parents[index] = parents[parents[index]!]!;
      index = parents[index]!;
    }
    return index;
  };
  const join = (a: number, b: number): void => {
    parents[find(b)] = find(a);
  };
  for (const dependency of movePlan.dependencies) {
    const own = dependency.flatMap((id) => (indices.has(id) ? [indices.get(id)!] : []));
    for (const index of own) join(own[0]!, index);
  }
  const owners = new Map<string, number>();
  // A card's sites are decided together. A site that goes with a removed cell or row is not
  // tied to its card's other sites: one record of a fill over a row's cells, one of them in a
  // deleted column, stays pending on the other cells, as Word keeps it when it accepts the
  // cell deletion alone
  for (const item of items) {
    const own = revisionSiteNodeIdsOf(item).flatMap((id) => {
      const index = indices.get(id);
      return index === undefined || contained.has(id) ? [] : [index];
    });
    for (const index of own) join(own[0]!, index);
  }
  const selectedTableSites = sites.filter((site) => selectedSites.has(site.node.id));
  const removedTables = new Set(tableRevisionRemovals(part, selectedTableSites, action).keys());
  const mergeSources = new Map<string, ReadonlySet<string>>();
  // Row decisions can remove every descendant. Property restoration replaces the live
  // property container. Containment also protects wrappers the resolver would sweep empty.
  for (const [index, site] of sites.entries()) {
    if (
      action === 'reject' &&
      site.propertyChange &&
      site.parent &&
      selectedSites.has(site.node.id)
    ) {
      // Restoration preserves live paragraph marks and, for pPrChange, rPr/sectPr.
      // Only children actually replaced by the resolver depend on this decision.
      for (const child of site.parent.children) {
        const preserved = preservedRevisionProperty(site.node.localName, child);
        if (!preserved) owners.set(child.id, index);
      }
    }
    // Removing a paragraph mark merges its text forward and drops its paragraph properties.
    // Keep numbering, section and formatting decisions in those properties in the same group.
    const removesMark =
      action === 'accept'
        ? site.node.localName === 'del' || site.node.localName === 'moveFrom'
        : site.node.localName === 'ins' || site.node.localName === 'moveTo';
    if (
      site.paragraphMark &&
      site.parent &&
      removesMark &&
      !orphanDestinations.has(site.node.id) &&
      selectedSites.has(site.node.id)
    ) {
      const properties = parentNodeOf(part, site.parent.id);
      const paragraph = properties && parentNodeOf(part, properties.id);
      const container = paragraph && parentNodeOf(part, paragraph.id);
      if (properties && paragraph && container) {
        let sources = mergeSources.get(container.id);
        if (!sources) {
          sources = paragraphMergeSources(container.children, removedTables);
          mergeSources.set(container.id, sources);
        }
        if (sources.has(paragraph.id)) owners.set(properties.id, index);
      }
    }
    if (selectedSites.has(site.node.id) && tableRevisionContentTarget(part, site, action)) {
      const target = tableRevisionContentTarget(part, site, action);
      if (target && !site.refused) owners.set(target.id, index);
    }
  }
  // A neighbour that takes a removed cell's width: the decision rewrites its properties. Its own
  // pending property change stays its own decision, as Word keeps it (the change record keeps
  // the width from before)
  const neighbourProperties = new Set<string>();
  for (const [cellId, markers] of tableRevisionNeighbours(part, selectedTableSites, action)) {
    const cell = findNode(part, cellId);
    const markerIndex = indices.get(markers[0]!);
    if (cell?.kind !== 'tableCell' || markerIndex === undefined) continue;
    for (const property of cell.children) {
      if (property.localName !== 'tcPr') continue;
      owners.set(property.id, markerIndex);
      neighbourProperties.add(property.id);
    }
  }
  const isTextChange = (index: number): boolean =>
    sites[index]!.node.kind === 'revisionInsert' || sites[index]!.node.kind === 'revisionDelete';
  const visit = (node: OoxmlNode, ancestor?: number, neighbour = false): void => {
    const site = indices.get(node.id);
    const owner = owners.get(node.id);
    if (site !== undefined && owner !== undefined) join(owner, site);
    const own = site ?? owner;

    // A format record on runs inside a text insertion or deletion is decided apart from it, as
    // Word decides them: deciding the record keeps or restores the runs' properties, and deciding
    // the text change keeps or removes the runs with the record on them
    const apart =
      own !== undefined &&
      sites[own]!.propertyChange &&
      ((neighbour && site !== undefined) || (ancestor !== undefined && isTextChange(ancestor)));
    if (own !== undefined && ancestor !== undefined && !apart) join(ancestor, own);
    const next = own ?? ancestor;
    const below = neighbour || neighbourProperties.has(node.id);
    if (node.kind !== 'textValue') for (const child of node.children) visit(child, next, below);
  };
  visit(root);
  for (const range of namedMoveRanges(root).values()) {
    const own = range.wrappers.flatMap((node) => {
      const index = indices.get(node.id);
      return index === undefined ? [] : [index];
    });
    for (const index of own) join(own[0]!, index);
  }
  for (const ids of sites.some((site) => site.node.localName === 'cellMerge')
    ? tableMergeDependencies(scopedPart)
    : []) {
    const group = ids.flatMap((id) => {
      const index = indices.get(id);
      return index === undefined ? [] : [index];
    });
    for (const index of group) join(group[0]!, index);
  }
  const groups = new Map<number, typeof sites>();
  for (const [index, site] of sites.entries()) {
    const id = find(index);
    const group = groups.get(id) ?? [];
    group.push(site);
    groups.set(id, group);
  }
  const reasons = new Map<number, RevisionBatchSkipReason>();
  const details = new Map<number, RevisionBatchSkipDetail>();
  for (const [id, group] of groups) {
    const refused = group.find((site) => site.refused);
    const unrecorded =
      action === 'reject' &&
      group.some((site) => site.propertyChange && recordedProperties(site.node) === null);
    if (refused || unrecorded) {
      reasons.set(id, 'unsupported-revision');
      details.set(id, refused?.refusal ?? 'unrecorded-properties');
    } else if (group.some((site) => !selectedSites.has(site.node.id))) {
      reasons.set(id, 'incomplete-group');
    }
  }
  for (const item of items) {
    if (!item.readOnly) continue;
    for (const id of revisionSiteNodeIdsOf(item)) {
      const index = indices.get(id);
      if (index === undefined) continue;
      reasons.set(find(index), 'unsupported-revision');
      // An item is read-only BECAUSE a site refused it: keep that site's specific reason.
      if (!details.has(find(index))) details.set(find(index), 'read-only');
    }
  }
  // Cell removal may remove its row; row removal may remove the containing table.
  // Only eligible groups contribute to this closure, so a retained row does not block peers.
  const eligible = sites.filter(
    (site, index) => selectedSites.has(site.node.id) && !reasons.has(find(index))
  );
  let removedContainer = false;
  const eligibleRemovals = tableRevisionRemovals(part, eligible, action);
  for (const [id, markerIds] of eligibleRemovals) {
    if (owners.has(id)) continue;
    const owner = indices.get(markerIds[0]!);
    if (owner !== undefined) {
      owners.set(id, owner);
      removedContainer = true;
    }
  }
  if (removedContainer) {
    visit(root);
    for (const [group, reason] of [...reasons]) {
      const root = find(group);
      if (reasons.get(root) !== 'unsupported-revision') {
        reasons.set(root, reason);
        const detail = details.get(group);
        if (detail) details.set(root, detail);
      }
    }
  }
  const resolved: RevisionBatchEntry[] = [];
  const skipped: RevisionBatchResult['skipped'][number][] = [];
  const acceptedSites = new Set<string>();
  const retainedRows = new Set(retainedNestedRowSites(part, sites, action));
  // A nested row can remain pending only if no eligible ancestor action removes it.
  for (const id of retainedRows) {
    let ancestor = parentNodeOf(part, id);
    while (ancestor) {
      if (eligibleRemovals.has(ancestor.id)) {
        retainedRows.delete(id);
        break;
      }
      ancestor = parentNodeOf(part, ancestor.id);
    }
  }
  for (const key of selected) {
    const item = byKey.get(key);
    if (!item) {
      skipped.push({ key, reason: 'unknown-revision' });
      continue;
    }
    const entry: RevisionBatchEntry = {
      key,
      author: item.author,
      partName: part.name,
      revisionKind: item.revisionKind,
      ...(item.formattingKind ? { formattingKind: item.formattingKind } : {}),
    };
    const ids = revisionSiteNodeIdsOf(item);
    let reason: RevisionBatchSkipReason | undefined;
    let detail: RevisionBatchSkipDetail | undefined;
    for (const id of ids) {
      const index = indices.get(id);
      reason = index === undefined ? 'incomplete-group' : reasons.get(find(index));
      detail = index === undefined ? undefined : details.get(find(index));
      if (reason) break;
    }
    if (!ids.length) {
      reason = 'unsupported-revision';
      detail = 'no-sites';
    }
    if (reason)
      skipped.push({
        key,
        reason,
        revision: entry,
        ...(reason === 'unsupported-revision' && detail ? { detail } : {}),
      });
    else {
      if (ids.some((id) => retainedRows.has(id)))
        skipped.push({ key, reason: 'retained-structure', revision: entry });
      else resolved.push(entry);
      for (const id of ids) if (!retainedRows.has(id)) acceptedSites.add(id);
    }
  }
  const gridCandidates = deferredTableGridSites(part, sites, acceptedSites);
  const deferredGrids = new Set<string>();
  for (const item of items) {
    const ids = revisionSiteNodeIdsOf(item);
    // Only a grouped row decision carries deferred grid metadata. A standalone
    // grid request is explicit and must retain its normal resolution semantics.
    if (item.formattingKind !== 'tblGridChange' && ids.length > 1)
      for (const id of ids) if (gridCandidates.has(id)) deferredGrids.add(id);
  }
  const siteNodeIds = [...acceptedSites].filter(
    (id) => !movePlan.implicit.has(id) && !deferredGrids.has(id)
  );
  let remaining = items.length - resolved.length;
  // Removing a row can make two formerly separated formatting groups adjacent.
  // Partial outcomes must count the resulting queue, not subtract from the old one.
  if (remaining > 0 && siteNodeIds.length) {
    // Resolution transfers its mutable node index to the rebuilt root. Keep the
    // original root's index and allocator state untouched during pure preflight.
    const previewPart = { ...part, root: { ...part.root } };
    const preview = resolveRevisions(previewPart, action, undefined, {
      siteNodeIds,
      ...(scopeRoot ? { scopeRootId: scopeRoot.id } : {}),
    });
    if (preview.ok && preview.part) {
      const afterRoot = scopeRoot ? findNode(preview.part, scopeRoot.id) : preview.part.root;
      remaining =
        afterRoot && afterRoot.kind !== 'textValue'
          ? revisionItemsOf({ ...preview.part, root: afterRoot }).length
          : 0;
    }
  }
  return {
    ops: acceptedSites.size
      ? [
          ...cleanup,
          {
            op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions',
            siteNodeIds,
            ...(scopeRoot ? { scopeRootId: scopeRoot.id } : {}),
          },
        ]
      : cleanup,
    result: { resolved, skipped, remaining },
  };
}
