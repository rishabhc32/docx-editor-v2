// One tracked table gesture is one review card.
//
// A resize records the table's own properties (`w:tblPrChange`) and grid (`w:tblGridChange`)
// beside every cell of every row, all under one revision: one `@w:id`, author and date. Word keeps
// a table record that changes something as its own decision, apart from the row groups, and
// Word writes every decision under its own `@w:id`. A table record that shares its revision with
// the one other formatting decision of its table was written by the same gesture, so it joins
// that decision, and Accept or Reject answers the gesture whole. The row groups cannot carry it
// when another author's cell record splits the gesture's rows apart: the grid then stood alone,
// and it hid behind the other author's card once the cells were answered.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import {
  registerRevisionSiteNodeIds,
  revisionSiteNodeIdsOf,
  type ReviewRevisionItem,
} from './review-items.ts';
import type { RevisionSite } from './tree-op-revisions.ts';

const TABLE_RECORD_NAMES = ['tblGridChange', 'tblPrChange'];

/** The items, with each table record that is a card of its own joined to its gesture's card. */
export function withTableRecordsJoined(
  part: OoxmlPart,
  items: readonly ReviewRevisionItem[],
  sites: ReadonlyMap<string, RevisionSite>
): ReviewRevisionItem[] {
  const identities = new Map<string, string>();
  /** The site's table and revision; a `w:tblGridChange` (`CT_Markup`) names an id alone. */
  const identity = (id: string, names: readonly string[]): string => {
    const key = `${names.length}:${id}`;
    const known = identities.get(key);
    if (known !== undefined) return known;
    const site = sites.get(id);
    let table: OoxmlNode | null = site?.parent ?? null;
    while (table && table.kind !== 'table') table = parentNodeOf(part, table.id);
    const attributes = site?.node.attributes ?? [];
    const value = (name: string) =>
      attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)?.value;
    const found = JSON.stringify([table?.id, ...names.map(value)]);
    identities.set(key, found);
    return found;
  };
  const result = [...items];
  for (const record of items) {
    const ids = revisionSiteNodeIdsOf(record);
    if (record.readOnly || !ids.length || !TABLE_RECORD_NAMES.includes(record.formattingKind ?? ''))
      continue;
    const names = record.formattingKind === 'tblGridChange' ? ['id'] : ['id', 'author', 'date'];
    const key = identity(ids[0]!, names);
    if (ids.some((id) => identity(id, names) !== key)) continue;
    const owners = result.filter(
      (item) =>
        item.revisionKind === 'format' &&
        !item.readOnly &&
        !TABLE_RECORD_NAMES.includes(item.formattingKind ?? '') &&
        revisionSiteNodeIdsOf(item).some((id) => identity(id, names) === key)
    );
    if (owners.length !== 1) continue;
    const owner = owners[0]!;
    const addresses = new Map(
      [...owner.addresses, ...record.addresses].map((address) => [
        JSON.stringify([address.id, address.author, address.date ?? null]),
        address,
      ])
    );
    result[result.indexOf(owner)] = registerRevisionSiteNodeIds(
      { ...owner, addresses: [...addresses.values()] },
      [...revisionSiteNodeIdsOf(owner), ...ids]
    );
    result.splice(result.indexOf(record), 1);
  }
  return result;
}
