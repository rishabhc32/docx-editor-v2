// Pending changes a structural decision takes with it.
//
// Rejecting a proposed row, cell or table — or accepting a proposed deletion of one — removes it
// and everything inside it. A change pending INSIDE then has nothing left to decide, and
// `planRevisionBatch` refused the whole decision while it stayed open: rejecting a table its
// author had proposed and typed into was refused until the typing's own card was answered first.
// Word resolves the two together.
//
// Only the SAME author's changes go along. Another reviewer's pending change inside is theirs to
// answer, so the decision stays refused until it is — with a reason that says so
// (`holdsPendingChange`). And only changes whose effect ends with the structure go along: text
// insertions and deletions, paragraph marks and formatting records. A move has a half outside,
// and a row or cell marker is a structural decision of its own, so neither ever does. Nor does
// anything in a nested table: Word reviews a nested table's changes apart from the rows around
// it (`review-table-groups.ts`), so those stay a dependency the reviewer answers first.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { revisionSiteNodeIdsOf, type ReviewRevisionItem } from './review-items.ts';
import { revisionAttribute, tableRevisionContentTarget } from './revision-table-plan.ts';
import { tableAncestorOf } from './tree-op-table-shared.ts';
import { collectRevisionSites, type RevisionSite } from './tree-op-revisions.ts';

/** Whether a row or cell sits in a table nested inside another table's cell. */
function inNestedTable(part: OoxmlPart, target: OoxmlElement): boolean {
  const table = tableAncestorOf(part, target.id);
  const host = table && parentNodeOf(part, table.id);
  return !!host && tableAncestorOf(part, host.id) !== null;
}

/**
 * Node id inside a structure the decision removes → the author whose marker removes it.
 *
 * `nested` includes nested tables. Without it the scope stops at them: Word reviews a nested
 * table's changes apart from the rows around it, so an outer decision never takes them along,
 * and a nested row's decision takes nothing along.
 */
function removedScope(
  part: OoxmlPart,
  sites: readonly RevisionSite[],
  selected: ReadonlySet<string>,
  action: 'accept' | 'reject',
  nested: boolean
): ReadonlyMap<string, string> {
  const scope = new Map<string, string>();
  const visit = (node: OoxmlNode, author: string): void => {
    if (!nested && node.kind === 'table') return;
    scope.set(node.id, author);
    if (node.kind !== 'textValue') for (const child of node.children) visit(child, author);
  };
  for (const site of sites) {
    if (!selected.has(site.node.id)) continue;
    const target = tableRevisionContentTarget(part, site, action);
    const author = revisionAttribute(site.node, 'author');
    if (!target || author === undefined || (!nested && inNestedTable(part, target))) continue;
    visit(target, author);
  }
  return scope;
}

/** A change whose whole effect lies where it sits, so it can vanish with its structure. */
function endsWithStructure(site: RevisionSite): boolean {
  return (
    !site.refused &&
    (site.paragraphMark ||
      site.propertyChange ||
      site.node.kind === 'revisionInsert' ||
      site.node.kind === 'revisionDelete')
  );
}

/**
 * The sites of the selected decisions' authors' own changes inside what those decisions remove.
 * `planRevisionBatch` selects them too, so the decision and the changes it removes resolve in
 * one transaction.
 */
export function containedOwnChangeSites(
  part: OoxmlPart,
  sites: readonly RevisionSite[],
  selected: ReadonlySet<string>,
  action: 'accept' | 'reject'
): readonly string[] {
  const scope = removedScope(part, sites, selected, action, false);
  if (scope.size === 0) return [];
  return sites
    .filter(
      (site) =>
        !selected.has(site.node.id) &&
        endsWithStructure(site) &&
        scope.get(site.node.id) === revisionAttribute(site.node, 'author')
    )
    .map((site) => site.node.id);
}

/**
 * Whether resolving `item` this way is held up by a change pending inside what it removes — one
 * that does not go along with it: another reviewer's, a move, or one in a nested table.
 */
export function holdsPendingChange(
  part: OoxmlPart,
  item: ReviewRevisionItem,
  action: 'accept' | 'reject'
): boolean {
  const sites = collectRevisionSites(part);
  const own = new Set(revisionSiteNodeIdsOf(item));
  const scope = removedScope(part, sites, own, action, true);
  if (scope.size === 0) return false;
  const going = new Set([...own, ...containedOwnChangeSites(part, sites, own, action)]);
  return sites.some((site) => !going.has(site.node.id) && scope.has(site.node.id));
}
