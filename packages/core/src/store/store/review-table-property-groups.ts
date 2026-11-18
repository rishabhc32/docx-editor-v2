// One tracked table gesture is one review card.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import { WML_NAMESPACE_URI, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import {
  registerRevisionSiteNodeIds,
  revisionSiteNodeIdsOf,
  type ReviewRevisionItem,
} from './review-items.ts';
import type { RevisionSite } from './tree-op-revisions.ts';

const TABLE_RECORD_NAMES = ['tblGridChange', 'tblPrChange'];

/**
 * The identity of a revision site: its table, and the revision attributes `names`. Cached by site
 * and names. A `w:tblGridChange` (`CT_Markup`) names an id alone.
 */
function tableRevisionIdentity(
  part: OoxmlPart,
  sites: ReadonlyMap<string, RevisionSite>
): (id: string, names: readonly string[]) => string {
  const identities = new Map<string, string>();
  return (id, names) => {
    const key = `${names.join(',')}:${id}`;
    const known = identities.get(key);
    if (known !== undefined) return known;
    const site = sites.get(id);
    let table: OoxmlNode | null = site?.parent ?? null;
    while (table && table.kind !== 'table') {
      table = parentNodeOf(part, table.id);
    }

    const attributes = site?.node.attributes ?? [];
    const value = (name: string) =>
      attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)?.value;
    const found = JSON.stringify([table?.id, ...names.map(value)]);
    identities.set(key, found);
    return found;
  };
}

/** `owner` with `record`'s addresses and its revision sites `ids` joined to its own. */
function joinedOwner(
  owner: ReviewRevisionItem,
  record: ReviewRevisionItem,
  ids: readonly string[]
): ReviewRevisionItem {
  const addresses = new Map(
    [...owner.addresses, ...record.addresses].map((address) => [
      JSON.stringify([address.id, address.author, address.date ?? null]),
      address,
    ])
  );
  return registerRevisionSiteNodeIds({ ...owner, addresses: [...addresses.values()] }, [
    ...revisionSiteNodeIdsOf(owner),
    ...ids,
  ]);
}

/**
 * The items, with each table record that is a card of its own joined to its gesture's card.
 *
 * A resize records the table's own properties (`w:tblPrChange`) and grid (`w:tblGridChange`)
 * beside every cell of every row, all under one revision: one `@w:id`, author and date. Word
 * writes every decision under its own `@w:id`, so a table record that shares its revision with
 * the one other formatting decision of its table was written by the same gesture: it joins that
 * decision, and Accept or Reject answers the gesture whole. The row groups cannot carry it when
 * another author's cell record splits the gesture's rows apart.
 */
export function withTableRecordsJoined(
  part: OoxmlPart,
  items: readonly ReviewRevisionItem[],
  sites: ReadonlyMap<string, RevisionSite>
): ReviewRevisionItem[] {
  const identity = tableRevisionIdentity(part, sites);
  const result = [...items];
  for (const record of items) {
    const ids = revisionSiteNodeIdsOf(record);
    if (
      record.readOnly ||
      ids.length === 0 ||
      !TABLE_RECORD_NAMES.includes(record.formattingKind ?? '')
    ) {
      continue;
    }

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
    result[result.indexOf(owner)] = joinedOwner(owner, record, ids);
    result.splice(result.indexOf(record), 1);
  }

  return result;
}
