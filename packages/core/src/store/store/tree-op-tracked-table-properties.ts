// Tracked TABLE property changes: `w:tcPrChange` (§17.13.5.37), `w:trPrChange`, `w:tblPrChange`
// and `w:tblGridChange`, the table twins of `w:rPrChange`. The table ops write each container
// as an untracked write does; this folds a change record into every container that write
// replaced (`withPropertyChangeRecord`), and `resolveRevisions` decides the records.

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
import {
  rowPropertyRecords,
  TABLE_RECORDS,
  type TableContainer,
} from './tree-op-tracked-table-records.ts';
import { tableHeightAnchor, wordHeightRows } from './revision-table-implicit-height.ts';
import { heightRevisionOf } from './revision-table-row-anchor.ts';
import { rowMarker } from './tree-op-tracked-table-deletion.ts';
import {
  addsPlainHeight,
  heightAnchorOf,
  rowAnchorEdit,
  type HeightAnchor,
} from './tree-op-tracked-table-row-height.ts';
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

/** The same child nodes by identity: the record left the container as the write made it. */
const sameChildren = (left: readonly OoxmlNode[], right: readonly OoxmlNode[]): boolean =>
  left.length === right.length && left.every((child, index) => child === right[index]);

/**
 * The `detail` of a write refused by `holdsOtherAuthorsRowChange`. The surface names the
 * refusal for the user with it (`surface-tracked-table-refusal.ts`).
 */
export const OTHER_AUTHORS_ROW_CHANGE = 'other-authors-row-change';

/**
 * Whether the row holds a pending property change of someone other than `author`. A record
 * with no `@w:author` is nobody's, so it is someone else's too.
 *
 * A row holds one author's decision at a time. A second author's write would take the first
 * author's records over on the cells it changed and leave them on the rest; rejecting both
 * changes would then reset those cells. So a write that changes such a row is refused until
 * that change is accepted or rejected. The author's own records take a second change.
 */
function holdsOtherAuthorsRowChange(row: OoxmlElement, author: string): boolean {
  return rowPropertyRecords(row).some((record) => wmlAttributeValue(record, 'author') !== author);
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
  const inserted = rowMarker(row, 'ins');
  if (inserted) {
    return wmlAttributeValue(inserted, 'author') === author;
  }

  return insideOwnInsertion(part, row.id, author);
}

/**
 * The author's own pending cell record in a row the write changed: the decision the write
 * joins. The first one, in document order; undefined when there is none.
 *
 * A cell holds one `w:tcPrChange`, so a write over the author's own pending cell record keeps
 * that record's `@w:id` and `@w:date` for all it records, as the ordered model keeps the cell's
 * one record. A height on a row with no `w:trPr` lives in its cells' records and its anchor
 * alone (`revision-table-implicit-height.ts`): a fill there under a new id would cut the height
 * off its record, and Reject All would keep the height.
 */
function ownRowDecision(
  part: OoxmlPart,
  rows: readonly OoxmlElement[],
  author: string
): OoxmlElement | undefined {
  for (const row of rows) {
    const prior = elementOf(findNode(part, row.id));
    if (prior === undefined || prior === row) continue;
    for (const cell of tableChildren(prior, 'tableCell')) {
      for (const child of propertiesOf(cell, 'tcPr')) {
        if (
          child.kind !== 'textValue' &&
          isWmlNamed(child, 'tcPrChange') &&
          wmlAttributeValue(child, 'author') === author
        ) {
          return child;
        }
      }
    }
  }

  return undefined;
}

/** What recording one table write shares between its steps. */
interface TableWrite {
  /** The part before the write. */
  readonly part: OoxmlPart;
  readonly revision: RevisionAttributionInput;
  readonly nextId: () => string;
  /** The op's one revision id (`withTablePropertyChanges`). */
  readonly mintRevisionId: () => string;
  /** The edits that write the records, applied in order to the written part. */
  readonly edits: ((current: OoxmlPart) => ReturnType<typeof replaceNode>)[];
  readonly dirty: Set<string>;
  readonly options: EditOptions | undefined;
}

/** A container's children after a write from `prior` to `next`, with the write recorded. */
function recorded(
  write: TableWrite,
  container: TableContainer,
  prior: readonly OoxmlNode[],
  next: readonly OoxmlNode[],
  recordUnchanged = false
): readonly OoxmlNode[] {
  const children = withPropertyChangeRecord({
    container,
    prior,
    next,
    revision: write.revision,
    mint: write.nextId,
    nextRevisionId: write.mintRevisionId,
    recordUnchanged,
  });
  return container === 'tableGrid' ? asGridMarkup(prior, children) : children;
}

/** Queue `node`, inside `owner`, to replace the node of its id. */
function replaceLater(write: TableWrite, owner: OoxmlElement, node: OoxmlElement): void {
  write.dirty.add(owner.id);
  write.edits.push((latest) => replaceNode(latest, node.id, node, write.options));
}

/** The children of `owner`'s container `name`, with a record when the write replaced it. */
function recordContainer(
  write: TableWrite,
  container: TableContainer,
  name: string,
  prior: OoxmlElement,
  owner: OoxmlElement
): readonly OoxmlNode[] {
  const now = wmlChildNamed(owner, name);
  const was = wmlChildNamed(prior, name);
  if (!now && was) {
    // The write emptied the container away; the record still needs one to stand in, placed
    // after a row's `w:tblPrEx` as `CT_Row` orders them.
    const children = recorded(write, container, was.children, []);
    if (children.length === 0) return children;
    const at = owner.children.findIndex((child) => isWmlNamed(child, 'tblPrEx')) + 1;
    const restored = build(write.nextId(), 'generic', name, [], children);
    replaceLater(write, owner, {
      ...owner,
      children: [...owner.children.slice(0, at), restored, ...owner.children.slice(at)],
    } as OoxmlElement);
    return children;
  }

  if (!now || now === was) return now?.children ?? [];
  const children = recorded(write, container, was?.children ?? [], now.children);
  if (!sameChildren(children, now.children)) {
    replaceLater(write, owner, { ...now, children } as OoxmlElement);
  }

  return children;
}

/**
 * Record the table's own properties and grid, which reset no cell. Whether the write aligns
 * the table: a table ALIGNMENT records every cell too, the form Word writes for it. Word reads
 * a lone `w:tblPrChange` that only adds a `w:jc` as unused history, not a decision
 * (`unboundTableHistories`), so on its own the proposal would offer nothing to review.
 */
function recordTableContainers(
  write: TableWrite,
  before: OoxmlElement,
  after: OoxmlElement,
  justification: string | null | undefined
): boolean {
  const tableProperties = recordContainer(write, 'tableProperties', 'tblPr', before, after);
  recordContainer(write, 'tableGrid', 'tblGrid', before, after);
  return justification !== undefined && hasRecord(tableProperties, 'tblPrChange');
}

/**
 * Record a row the write changed: its `w:trPr` and each cell's `w:tcPr`. Whether the row took a
 * height in Word's plain form, with no row record (`addsPlainHeight`).
 *
 * Word reads the records in a row as one formatting decision, and rejecting it resets every
 * cell of the row that has no record to default formatting (`restoreUnrecordedCells`). So while
 * a row records a change, to one of its cells or to the row itself, every cell of that row
 * records its properties, changed or not (`recordUnchanged`).
 */
function recordRow(
  write: TableWrite,
  table: OoxmlElement,
  prior: OoxmlElement,
  row: OoxmlElement,
  anchor: HeightAnchor,
  aligned: boolean
): boolean {
  const plain = addsPlainHeight(table, prior, row, write.revision.author);
  const rowProperties = plain
    ? propertiesOf(row, 'trPr')
    : recordContainer(write, 'rowProperties', 'trPr', prior, row);
  const cells = tableChildren(row, 'tableCell');
  const priorCells = cells.map((cell) => elementOf(findNode(write.part, cell.id)));
  const current = cells.map((cell) => propertiesOf(cell, 'tcPr'));
  const cellRecord = (index: number, recordUnchanged = false): readonly OoxmlNode[] =>
    recorded(
      write,
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
    plain ||
    hasRecord(rowProperties, 'trPrChange') ||
    hasRecord(propertiesOf(row, 'tblPrEx'), 'tblPrExChange') ||
    next.some((children) => hasRecord(children, 'tcPrChange'))
  ) {
    for (const [index, children] of next.entries()) {
      if (!hasRecord(children, 'tcPrChange')) next[index] = cellRecord(index, true);
    }
  }

  // The row's own anchor goes in before its cells' records, which replace cells inside it
  if (plain && anchor === 'rows') {
    write.dirty.add(row.id);
    write.edits.push(
      rowAnchorEdit(row.id, write.revision, write.nextId, write.mintRevisionId, write.options)
    );
  }

  for (const [index, cell] of cells.entries()) {
    if (next[index] === current[index]) continue;
    replaceLater(write, cell, withCellProperties(cell, next[index]!, write.nextId));
  }

  return plain;
}

/**
 * Move each height that leaned on the table's record onto its row's own record, when the write
 * takes that record over (Word's rule: the latest author's).
 */
function moveLeaningHeights(write: TableWrite, before: OoxmlElement, after: OoxmlElement): void {
  if (!hasRecord(propertiesOf(after, 'tblPr'), 'tblPrChange')) return;
  if (wmlChildNamed(after, 'tblPr') === wmlChildNamed(before, 'tblPr')) return;
  const held = tableHeightAnchor(before);
  for (const height of wordHeightRows(before)) {
    if (height.rowAnchor || height.anchor !== held) continue;
    const revision = heightRevisionOf(height);
    if (revision === undefined) continue;
    write.dirty.add(height.row.id);
    write.edits.push(
      rowAnchorEdit(height.row.id, revision, write.nextId, undefined, write.options)
    );
  }
}

/** The table's anchor for the plain heights: its properties and grid, recorded unchanged. */
function anchorPlainHeights(write: TableWrite, table: OoxmlElement): void {
  for (const [container, name] of [
    ['tableProperties', 'tblPr'],
    ['tableGrid', 'tblGrid'],
  ] as const) {
    const now = wmlChildNamed(table, name)!;
    if (hasRecord(now.children, TABLE_RECORDS[container].changeName)) continue;
    const children = recorded(write, container, now.children, now.children, true);
    replaceLater(write, table, { ...now, children } as OoxmlElement);
  }
}

/**
 * An untracked table write, recorded as tracked format changes: the record holds the ORIGINAL
 * properties, a second change keeps it, and a write that lands back on it leaves none.
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
  const asked = op.revision;
  if (!result.ok || asked === undefined) return result;
  const before = findNode(part, op.tableId);
  const after = findNode(result.part, op.tableId);
  if (before?.kind !== 'table' || after?.kind !== 'table') return result;
  const rows = tableChildren(after, 'tableRow');
  const ownRows = new Set(
    rows
      .filter((row) => insideOwnRow(part, elementOf(findNode(part, row.id)) ?? row, asked.author))
      .map((row) => row.id)
  );

  // The author's own pending decision in a row the write changes, which it joins
  const joined = ownRowDecision(
    part,
    rows.filter((row) => !ownRows.has(row.id)),
    asked.author
  );
  const joinedId = joined && wmlAttributeValue(joined, 'id');
  const joinedDate = joined && wmlAttributeValue(joined, 'date');
  const revision: RevisionAttributionInput = joined
    ? { author: asked.author, ...(joinedDate === undefined ? {} : { date: joinedDate }) }
    : asked;

  // ONE id for the whole op, taken LAZILY: the decision it joins, else the transaction's when it
  // lends one (the rule `setRunProperties` follows), so one gesture over many cells, rows and
  // the grid is one card.
  let ownId: string | null = null;
  const write: TableWrite = {
    part,
    revision,
    nextId: createNodeIdAllocator(result.part),
    mintRevisionId: () =>
      joinedId ?? options?.revisionIds?.() ?? (ownId ??= nextRevisionId(result.part)()),
    edits: [],
    dirty: new Set(result.effect.dirty),
    options,
  };

  // A table this author proposed whole goes, properties and grid too, when that is rejected.
  let aligned = false;
  if (rows.some((row) => !ownRows.has(row.id))) {
    aligned = recordTableContainers(write, before, after, op.justification);
  }

  const plainRows: string[] = [];
  const anchor = heightAnchorOf(before);
  for (const row of rows) {
    const prior = elementOf(findNode(part, row.id));
    if (prior === undefined || ownRows.has(row.id) || (prior === row && !aligned)) continue;
    if (holdsOtherAuthorsRowChange(prior, revision.author)) {
      return { ok: false, reason: 'unsupported-revision', detail: OTHER_AUTHORS_ROW_CHANGE };
    }

    if (recordRow(write, before, prior, row, anchor, aligned)) plainRows.push(row.id);
  }

  if (anchor === 'held') moveLeaningHeights(write, before, after);
  if (plainRows.length > 0 && anchor === 'table') anchorPlainHeights(write, after);
  if (write.edits.length === 0) return result;
  return fromEdit(applyEdits(result.part, write.edits, options), {
    ...result.effect,
    dirty: [...write.dirty],
  });
}
