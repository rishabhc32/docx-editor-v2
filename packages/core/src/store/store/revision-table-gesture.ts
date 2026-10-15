// A tracked table resize resolves as one decision.
//
// A tracked resize records the table's grid (`w:tblGridChange`) beside the row and cell records
// of its table, all under one revision (`withTablePropertyChanges`). `deferredTableGridSites`
// keeps a selected grid record standing while another row decision of its table survives,
// because a grid Word wrote is shared by them all. A grid selected together with a row record
// of its own revision is that gesture's, so the two resolve together.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import type { RevisionSite } from './tree-op-revisions.ts';

const ROW_RECORDS: ReadonlySet<string> = new Set(['trPrChange', 'tcPrChange', 'tblPrExChange']);

function revisionIdOf(site: RevisionSite): string | undefined {
  return site.node.attributes.find(
    (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
  )?.value;
}

/** The nearest table at or above `node`. */
function tableAt(part: OoxmlPart, node: OoxmlNode | null): OoxmlNode | null {
  let ancestor = node;
  while (ancestor && ancestor.kind !== 'table') ancestor = parentNodeOf(part, ancestor.id);
  return ancestor;
}

/** Whether a selected row record of the grid's table shares the grid record's `@w:id`. */
export function selectedWithItsGesture(
  part: OoxmlPart,
  grid: RevisionSite,
  sites: readonly RevisionSite[],
  selected: ReadonlySet<string>
): boolean {
  const table = grid.parent && parentNodeOf(part, grid.parent.id);
  if (!table) return false;
  const id = revisionIdOf(grid);
  return sites.some(
    (site) =>
      ROW_RECORDS.has(site.node.localName) &&
      selected.has(site.node.id) &&
      revisionIdOf(site) === id &&
      tableAt(part, site.parent)?.id === table.id
  );
}
