// Tracked TABLE property changes — `w:tcPrChange` (§17.13.5.37), `w:trPrChange`,
// `w:tblPrChange` and `w:tblGridChange`, the table twins of `w:rPrChange`.
//
// The table ops work out each new `w:tcPr`, `w:trPr`, `w:tblPr` and `w:tblGrid` exactly as an
// untracked write does. This folds a change record into every container that write replaced,
// through `withPropertyChangeRecord`, so a table follows the rules a run does: the record holds
// the ORIGINAL properties, a second change keeps it, and a write that lands back on it leaves
// none. `resolveRevisions` already resolves each record — Accept drops it, Reject restores what
// it holds — so this is the WRITE half.
//
// ONE RULE IS THE TABLE'S OWN. Word reads the records in a row as one formatting decision, and
// rejecting it resets every cell of the row that has no record to default formatting
// (`restoreUnrecordedCells`, confirmed against Word). A record on the changed cell alone made
// Reject wipe its neighbours' widths and shading. So while a row records a change — to one of
// its cells or to the row itself — every cell of that row records its properties, changed or
// not (`recordUnchanged`). The table's own records, `w:tblPrChange` and `w:tblGridChange`,
// reset no cell.
//
// SO A ROW HOLDS ONE AUTHOR'S DECISION AT A TIME. A second author's write took the first
// author's records over on the cells it changed and left them on the rest. Rejecting the
// second author's change then left those cells with no record, and rejecting the first
// author's reset them: their widths were lost. A write that changes a row holding another
// author's pending property change is refused until that change is accepted or rejected
// (`holdsOtherAuthorsRowChange`). The author's own records take a second change as before.
//
// A `w:tblGridChange` is `CT_Markup`: an `@w:id` and nothing else (`asGridMarkup`).

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
import { tableChildren } from './revision-table-children.ts';
import { fromEdit } from './tree-op-nodes.ts';
import type { RevisionAttributionInput } from './tree-op-revision-attribution.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { wmlAttributeValue, wmlChildNamed } from './tree-op-table-shared.ts';
import { build, isWmlNamed } from './tree-op-tracked.ts';
import { insideOwnInsertion, withPropertyChangeRecord } from './tree-op-tracked-properties.ts';
import type { TableContainer } from './tree-op-tracked-table-records.ts';
import type { TreeOpResult } from './tree-op-types.ts';

/** A node as an element, or undefined for a text value or a missing node. */
function elementOf(node: OoxmlNode | null): OoxmlElement | undefined {
  return node && node.kind !== 'textValue' ? node : undefined;
}

/** A container's children, or none when its owner has no such container. */
function propertiesOf(owner: OoxmlElement | undefined, name: string): readonly OoxmlNode[] {
  return (owner && wmlChildNamed(owner, name)?.children) ?? [];
}

const hasRecord = (children: readonly OoxmlNode[], name: string): boolean =>
  children.some((child) => isWmlNamed(child, name));

const sameChildren = (left: readonly OoxmlNode[], right: readonly OoxmlNode[]): boolean =>
  left.length === right.length && left.every((child, index) => child === right[index]);

/**
 * The `detail` of a write refused by `holdsOtherAuthorsRowChange`. The surface names the
 * refusal for the user with it (`surface-tracked-table-refusal.ts`).
 */
export const OTHER_AUTHORS_ROW_CHANGE = 'other-authors-row-change';

/** The records that make a row's cells one decision (`restoresRowCells`). */
const ROW_RECORDS: readonly string[] = ['trPrChange', 'tblPrExChange', 'tcPrChange'];

/**
 * Whether the row holds a pending property change of someone other than `author`. A record
 * with no `@w:author` is nobody's, so it is someone else's too.
 */
function holdsOtherAuthorsRowChange(row: OoxmlElement, author: string): boolean {
  const properties = [
    ...propertiesOf(row, 'trPr'),
    ...propertiesOf(row, 'tblPrEx'),
    ...tableChildren(row, 'tableCell').flatMap((cell) => propertiesOf(cell, 'tcPr')),
  ];
  return properties.some(
    (child) =>
      child.kind !== 'textValue' &&
      ROW_RECORDS.some((name) => isWmlNamed(child, name)) &&
      wmlAttributeValue(child, 'author') !== author
  );
}

/** The cell with these `w:tcPr` children; `w:tcPr` leads the cell (`CT_Tc`). */
function withCellProperties(
  cell: OoxmlElement,
  children: readonly OoxmlNode[],
  mint: () => string
): OoxmlElement {
  const properties = wmlChildNamed(cell, 'tcPr');
  const replaced = properties
    ? ({ ...properties, children } as OoxmlElement)
    : build(mint(), 'generic', 'tcPr', [], children);
  return {
    ...cell,
    children: [replaced, ...cell.children.filter((child) => child !== properties)],
  } as OoxmlElement;
}

/**
 * A grid's children with its record as `CT_Markup` writes it: an `@w:id` and nothing else.
 *
 * So a grid record names no author, and has none to protect. The record engine keeps a record
 * with no `@w:author` standing, as somebody else's, when a write lands back on the original; a
 * grid back at its record decides nothing either way, so its record goes.
 */
function asGridMarkup(
  prior: readonly OoxmlNode[],
  children: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  const standing = prior.find((child) => isWmlNamed(child, 'tblGridChange'));
  return children.flatMap((child): OoxmlNode[] => {
    if (child === standing) return [];
    if (child.kind === 'textValue' || !isWmlNamed(child, 'tblGridChange')) return [child];
    const id = child.attributes.filter(
      (attribute) => attribute.namespaceUri === WML_NAMESPACE_URI && attribute.localName === 'id'
    );
    return [{ ...child, attributes: id } as OoxmlElement];
  });
}

/**
 * Whether this author proposed the row, or something around it, as an insertion. Rejecting that
 * insertion removes the row, formatting and all, so nothing in it needs a record (#495).
 */
function insideOwnRow(part: OoxmlPart, row: OoxmlElement, author: string): boolean {
  const inserted = propertiesOf(row, 'trPr').find((child) => isWmlNamed(child, 'ins'));
  if (inserted && inserted.kind !== 'textValue') {
    return wmlAttributeValue(inserted, 'author') === author;
  }
  return insideOwnInsertion(part, row.id, author);
}

/**
 * An untracked table write, recorded as tracked format changes.
 *
 * `part` is the part before the write and `result` is the write. A refused write, or one that
 * changed nothing, passes through as it is: a press that moves no property records nothing. An
 * op without a `revision` is untracked and passes through too. A write that changes a row
 * holding another author's pending property change is refused (`unsupported-revision`).
 */
export function withTablePropertyChanges(
  part: OoxmlPart,
  result: TreeOpResult,
  op: {
    readonly tableId: string;
    readonly justification?: string | null;
    readonly revision?: RevisionAttributionInput;
  },
  options?: EditOptions
): TreeOpResult {
  const { revision } = op;
  if (!result.ok || revision === undefined) return result;
  const before = findNode(part, op.tableId);
  const after = findNode(result.part, op.tableId);
  if (before?.kind !== 'table' || after?.kind !== 'table') return result;
  const nextId = createNodeIdAllocator(result.part);
  // ONE id for the whole op, taken LAZILY, and the transaction's when it lends one — the rule
  // `setRunProperties` follows, so one gesture over many cells, rows and the grid is one card.
  let ownId: string | null = null;
  const mintRevisionId = (): string =>
    options?.revisionIds?.() ?? (ownId ??= nextRevisionId(result.part)());
  const record = (
    container: TableContainer,
    prior: readonly OoxmlNode[],
    next: readonly OoxmlNode[],
    recordUnchanged = false
  ): readonly OoxmlNode[] => {
    const children = withPropertyChangeRecord({
      container,
      prior,
      next,
      revision,
      mint: nextId,
      nextRevisionId: mintRevisionId,
      recordUnchanged,
    });
    return container === 'tableGrid' ? asGridMarkup(prior, children) : children;
  };

  const edits: ((current: OoxmlPart) => ReturnType<typeof replaceNode>)[] = [];
  const dirty = new Set(result.effect.dirty);
  const replace = (owner: OoxmlElement, node: OoxmlElement): void => {
    dirty.add(owner.id);
    edits.push((latest) => replaceNode(latest, node.id, node, options));
  };
  /** The children of `owner`'s container `name`, with a record when the write replaced it. */
  const recordContainer = (
    container: TableContainer,
    name: string,
    prior: OoxmlElement,
    owner: OoxmlElement
  ): readonly OoxmlNode[] => {
    const now = wmlChildNamed(owner, name);
    const was = wmlChildNamed(prior, name);
    if (!now && was) {
      // The write emptied the container away; the record still needs one to stand in, placed
      // after a row's `w:tblPrEx` as `CT_Row` orders them.
      const children = record(container, was.children, []);
      if (children.length === 0) return children;
      const at = owner.children.findIndex((child) => isWmlNamed(child, 'tblPrEx')) + 1;
      const restored = build(nextId(), 'generic', name, [], children);
      replace(owner, {
        ...owner,
        children: [...owner.children.slice(0, at), restored, ...owner.children.slice(at)],
      } as OoxmlElement);
      return children;
    }
    if (!now || now === was) return now?.children ?? [];
    const children = record(container, was?.children ?? [], now.children);
    if (!sameChildren(children, now.children)) replace(owner, { ...now, children } as OoxmlElement);
    return children;
  };

  const rows = tableChildren(after, 'tableRow');
  const ownRows = new Set(
    rows
      .filter((row) =>
        insideOwnRow(part, elementOf(findNode(part, row.id)) ?? row, revision.author)
      )
      .map((row) => row.id)
  );
  // A table this author proposed whole goes, properties and grid too, when that is rejected.
  let aligned = false;
  if (rows.some((row) => !ownRows.has(row.id))) {
    const tableProperties = recordContainer('tableProperties', 'tblPr', before, after);
    recordContainer('tableGrid', 'tblGrid', before, after);
    // A table ALIGNMENT records every cell too, the form Word writes for it. Word reads a lone
    // `w:tblPrChange` that only adds a `w:jc` as unused history, not a decision
    // (`unboundTableHistories`), so on its own the proposal would offer nothing to review.
    aligned = op.justification !== undefined && hasRecord(tableProperties, 'tblPrChange');
  }

  for (const row of rows) {
    const prior = elementOf(findNode(part, row.id));
    if (prior === undefined || ownRows.has(row.id) || (prior === row && !aligned)) continue;
    if (holdsOtherAuthorsRowChange(prior, revision.author)) {
      return { ok: false, reason: 'unsupported-revision', detail: OTHER_AUTHORS_ROW_CHANGE };
    }
    const rowProperties = recordContainer('rowProperties', 'trPr', prior, row);
    const cells = tableChildren(row, 'tableCell');
    const priorCells = cells.map((cell) => elementOf(findNode(part, cell.id)));
    const current = cells.map((cell) => propertiesOf(cell, 'tcPr'));
    const cellRecord = (index: number, recordUnchanged = false): readonly OoxmlNode[] =>
      record(
        'cellProperties',
        propertiesOf(priorCells[index], 'tcPr'),
        current[index]!,
        recordUnchanged
      );
    const next = cells.map((cell, index) => {
      const was = priorCells[index];
      if (!was || wmlChildNamed(was, 'tcPr') !== wmlChildNamed(cell, 'tcPr')) {
        return cellRecord(index);
      }
      // An untouched cell sheds a record of this author's that no longer records a change,
      // and otherwise stays exactly as it is.
      const shed = cellRecord(index);
      return hasRecord(current[index]!, 'tcPrChange') && !hasRecord(shed, 'tcPrChange')
        ? shed
        : current[index]!;
    });
    if (
      aligned ||
      hasRecord(rowProperties, 'trPrChange') ||
      hasRecord(propertiesOf(row, 'tblPrEx'), 'tblPrExChange') ||
      next.some((children) => hasRecord(children, 'tcPrChange'))
    ) {
      for (const [index, children] of next.entries()) {
        if (!hasRecord(children, 'tcPrChange')) next[index] = cellRecord(index, true);
      }
    }
    for (const [index, cell] of cells.entries()) {
      if (next[index] === current[index]) continue;
      replace(cell, withCellProperties(cell, next[index]!, nextId));
    }
  }
  if (edits.length === 0) return result;
  return fromEdit(applyEdits(result.part, edits, options), {
    ...result.effect,
    dirty: [...dirty],
  });
}
