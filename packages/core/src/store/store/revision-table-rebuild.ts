import { implicitGridRows, restoreImplicitRowGrid } from './revision-table-implicit-grid.ts';
import { parentNodeOf } from '../package/ooxml-edit.ts';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import type { RevisionSite } from './tree-op-revisions.ts';
import { recordedProperties } from './tree-op-tracked-properties.ts';
import { tableChildren, replaceTableChildren } from './revision-table-children.ts';
import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import { isWmlNamed } from './tree-op-tracked.ts';
import { revisionAttribute } from './revision-table-plan.ts';
import { readTwipsMeasure } from '../units.ts';

function elements(node: OoxmlElement): OoxmlElement[] {
  const result: OoxmlElement[] = [];
  for (const entry of node.children) if (entry.kind !== 'textValue') result.push(entry);
  return result;
}
function child(node: OoxmlElement, name: string): OoxmlElement | undefined {
  return elements(node).find((n) => isWmlNamed(n, name));
}
function number(node: OoxmlElement | undefined, name: string, fallback: number): number {
  const value = node && Number(revisionAttribute(node, name));
  return value !== undefined && Number.isFinite(value) && value > 0 ? value : fallback;
}
/** A positive `w:w` in whole units: a decimal truncates, as every width reader reads it. */
function positiveWidth(node: OoxmlElement | undefined, fallback: number): number {
  const value = readTwipsMeasure(node && revisionAttribute(node, 'w'));
  return value !== null && value > 0 ? value : fallback;
}
function element(
  source: OoxmlElement,
  name: string,
  attrs: Record<string, string>,
  mint: () => string
): OoxmlElement {
  return {
    ...source,
    id: mint(),
    kind: 'generic',
    localName: name,
    children: [],
    attributes: Object.entries(attrs).map(([localName, value]) => ({
      kind: localName === 'val' ? 'wmlVal' : 'genericExtension',
      namespaceUri: WML_NAMESPACE_URI,
      prefix: source.prefix,
      localName,
      value,
    })),
  } as OoxmlElement;
}
function properties(cell: OoxmlElement, mint: () => string): OoxmlElement {
  return child(cell, 'tcPr') ?? element(cell, 'tcPr', {}, mint);
}
function patch(
  cell: OoxmlElement,
  values: Record<string, Record<string, string> | null>,
  mint: () => string
): OoxmlElement {
  const pr = properties(cell, mint);
  const order = [
    'cnfStyle',
    'tcW',
    'gridSpan',
    'hMerge',
    'vMerge',
    'tcBorders',
    'shd',
    'noWrap',
    'tcMar',
    'textDirection',
    'tcFitText',
    'vAlign',
    'hideMark',
    'headers',
    'cellIns',
    'cellDel',
    'cellMerge',
    'tcPrChange',
  ];
  const children = pr.children.filter(
    (n) =>
      n.kind === 'textValue' || n.namespaceUri !== WML_NAMESPACE_URI || !(n.localName in values)
  );
  for (const [name, attrs] of Object.entries(values)) {
    if (!attrs) continue;
    const index = children.findIndex(
      (n) =>
        n.kind !== 'textValue' &&
        n.namespaceUri === WML_NAMESPACE_URI &&
        order.indexOf(n.localName) > order.indexOf(name)
    );
    const existing = child(pr, name);
    const replacement = element(existing ?? pr, name, attrs, mint);
    const updated = existing
      ? ({
          ...existing,
          attributes: [
            ...existing.attributes.filter(
              (attribute) =>
                attribute.namespaceUri !== WML_NAMESPACE_URI || !(attribute.localName in attrs)
            ),
            ...replacement.attributes,
          ],
        } as OoxmlElement)
      : replacement;
    children.splice(index < 0 ? children.length : index, 0, updated);
  }
  return {
    ...cell,
    children: [
      { ...pr, children } as OoxmlElement,
      ...cell.children.filter((n) => !isWmlNamed(n, 'tcPr')),
    ],
  } as OoxmlElement;
}
export function applyCellMerge(
  cell: OoxmlElement,
  value: string | undefined,
  mint: () => string
): OoxmlElement {
  return patch(
    cell,
    { vMerge: value === undefined ? null : { val: value === 'cont' ? 'continue' : 'restart' } },
    mint
  );
}

export function restoredCellRows(part: OoxmlPart, sites: readonly RevisionSite[]): string[] {
  const rows = new Set<string>();
  for (const site of sites) {
    if (
      site.refused ||
      !['trPrChange', 'tblPrExChange', 'tcPrChange'].includes(site.node.localName)
    )
      continue;
    let node = site.parent && parentNodeOf(part, site.parent.id);
    while (node && node.kind !== 'tableRow' && node.kind !== 'table')
      node = parentNodeOf(part, node.id);
    if (node?.kind === 'tableRow') rows.add(node.id);
  }
  const selected = new Set(sites.map((s) => s.node.id));
  const tables = new Map<string, OoxmlElement>();
  for (const id of rows) {
    let table = parentNodeOf(part, id);
    while (table && table.kind !== 'table') table = parentNodeOf(part, table.id);
    if (table) tables.set(table.id, table);
  }
  for (const table of tables.values())
    if (implicitGridRows(table, selected).size) rows.add(table.id);
  return [...rows];
}

/** A rejected row formatting decision restores unrecorded cells to default formatting. */
export function restoresRowCells(row: OoxmlElement, restored: ReadonlySet<string>): boolean {
  const containers = [
    child(row, 'trPr'),
    child(row, 'tblPrEx'),
    ...tableChildren(row, 'tableCell').map((cell) => child(cell, 'tcPr')),
  ];
  return containers.some((pr) => pr?.children.some((record) => restored.has(record.id)));
}
function hasExtensionData(node: OoxmlNode): boolean {
  return (
    node.kind !== 'textValue' &&
    (node.namespaceUri !== WML_NAMESPACE_URI ||
      node.attributes.some((a) => a.namespaceUri !== WML_NAMESPACE_URI) ||
      node.children.some(hasExtensionData))
  );
}
function unrecordedDefaults(cell: OoxmlElement): Record<string, Record<string, string> | null> {
  const values: Record<string, Record<string, string> | null> = { tcW: { w: '0', type: 'auto' } };
  for (const name of [
    'cnfStyle',
    'tcBorders',
    'shd',
    'noWrap',
    'tcMar',
    'textDirection',
    'tcFitText',
    'vAlign',
    'hideMark',
  ]) {
    const property = child(child(cell, 'tcPr') ?? cell, name);
    // Unknown extension payloads do not belong to the recorded formatting decision.
    if (!property || !hasExtensionData(property)) values[name] = null;
  }
  return values;
}
/** A cell with no record of its own, as a rejected row formatting decision leaves it. */
export function withUnrecordedCellDefaults(cell: OoxmlElement, mint: () => string): OoxmlElement {
  return patch(cell, unrecordedDefaults(cell), mint);
}
function restoreUnrecordedCells(
  original: OoxmlElement,
  row: OoxmlElement,
  restored: ReadonlySet<string>,
  mint: () => string
): OoxmlElement {
  if (!restoresRowCells(original, restored)) return row;
  const recorded = new Set(
    tableChildren(original, 'tableCell')
      .filter((cell) => {
        const pr = child(cell, 'tcPr');
        return pr && child(pr, 'tcPrChange');
      })
      .map((cell) => cell.id)
  );
  return replaceTableChildren(
    row,
    new Map(
      tableChildren(row, 'tableCell').map((cell) => [
        cell.id,
        recorded.has(cell.id) ? cell : withUnrecordedCellDefaults(cell, mint),
      ])
    )
  );
}

/**
 * The cell's pending `w:tcPrChange` takes the width and grid span a rejected new column gives back
 * when it holds the cell as it showed: a record taken beside a pending column holds the width the
 * column lent (Word writes no record for it), so a later reject of the record keeps what the
 * column gave back (a party's table or row property over its own new column). An accepted deleted
 * column's width is not lent: the record keeps the width from before, as Word keeps it.
 */
function withRecordTakingCell(
  cell: OoxmlElement,
  shown: { readonly width: OoxmlElement; readonly span: number },
  taken: { readonly width: number; readonly span: number },
  mint: () => string
): OoxmlElement {
  const pr = child(cell, 'tcPr');
  const record = pr && child(pr, 'tcPrChange');
  const inner = record && child(record, 'tcPr');
  const recorded = inner && child(inner, 'tcW');
  const type = revisionAttribute(shown.width, 'type');
  if (!pr || !record || !inner || !recorded || !type) return cell;
  if (
    revisionAttribute(recorded, 'type') !== type ||
    positiveWidth(recorded, 0) !== positiveWidth(shown.width, 0) ||
    number(child(inner, 'gridSpan'), 'val', 1) !== shown.span
  )
    return cell;
  return withOwnRecord(cell, 'tcPr', (own) =>
    patch(
      own,
      {
        tcW: { type, w: String(positiveWidth(shown.width, 0) + taken.width) },
        gridSpan: { val: String(shown.span + taken.span) },
      },
      mint
    )
  );
}

/**
 * The cell that takes each removed cell's grid space, by node id: the survivor before it, or at
 * the row's start the first survivor after it. A row with no survivor gives its space to none.
 */
export function spaceReceivers(
  cells: readonly OoxmlElement[],
  removed: ReadonlySet<string>
): Map<string, string> {
  const out = new Map<string, string>();
  const leading: string[] = [];
  let previous: string | undefined;
  for (const cell of cells) {
    if (!removed.has(cell.id)) {
      previous = cell.id;
      for (const id of leading.splice(0)) {
        out.set(id, previous);
      }
    } else if (previous === undefined) {
      leading.push(cell.id);
    } else {
      out.set(cell.id, previous);
    }
  }

  return out;
}

/** Word gives deleted cell space to the preceding survivor (or the first following cell). */
export function rebuildRevisionTable(
  original: OoxmlElement,
  rebuilt: OoxmlElement,
  removed: ReadonlySet<string>,
  mint: () => string,
  mergedCells: ReadonlyMap<string, OoxmlElement>,
  restoredProperties: ReadonlySet<string>
): OoxmlElement {
  const originalRows = new Map(tableChildren(original, 'tableRow').map((n) => [n.id, n]));
  let rows: OoxmlElement[] = tableChildren(rebuilt, 'tableRow');
  rows = rows.map((row) => {
    const prior = originalRows.get(row.id);
    if (!prior) return row;
    row = restoreUnrecordedCells(prior, row, restoredProperties, mint);
    const priorCells = tableChildren(prior, 'tableCell');
    if (!priorCells.some((c) => removed.has(c.id))) return row;
    const additions = new Map<
      string,
      { span: number; width: number; type: string | undefined; lent: boolean }
    >();
    const receivers = spaceReceivers(priorCells, removed);
    for (const cell of priorCells) {
      const receiver = receivers.get(cell.id);
      if (receiver === undefined) {
        continue;
      }

      const pr = child(cell, 'tcPr');
      const span = number(pr && child(pr, 'gridSpan'), 'val', 1);
      const widthNode = pr && child(pr, 'tcW');
      const width = positiveWidth(widthNode, 0);
      const type = widthNode && revisionAttribute(widthNode, 'type');
      const existing = additions.get(receiver) ?? { span: 0, width: 0, type, lent: true };
      additions.set(receiver, {
        span: existing.span + span,
        width: existing.width + width,
        type: existing.type === type ? type : undefined,
        // A pending new column's width was lent by its neighbour; a deleted one's is its own
        lent: existing.lent && pr !== undefined && child(pr, 'cellIns') !== undefined,
      });
    }

    const cells = tableChildren(row, 'tableCell').map((c) => {
      const extra = additions.get(c.id);
      if (!extra || c.kind !== 'tableCell') return c;
      const originalCell = priorCells.find((cell) => cell.id === c.id);
      const originalProperties = originalCell && child(originalCell, 'tcPr');
      const pr = child(c, 'tcPr');
      const width = pr && child(pr, 'tcW');
      // A rejected cell-property record already describes the pre-insertion geometry, unless it
      // holds the width the cell showed: a record taken beside a pending column holds the lent
      // width (Word writes no record for it), so the cell still takes the removed width back.
      const shown = originalProperties && child(originalProperties, 'tcW');
      const lent =
        width !== undefined &&
        shown !== undefined &&
        revisionAttribute(width, 'type') === revisionAttribute(shown, 'type') &&
        positiveWidth(width, 0) === positiveWidth(shown, 0);
      if (
        !lent &&
        originalProperties?.children.some((property) => restoredProperties.has(property.id))
      )
        return c;
      const currentSpan = number(pr && child(pr, 'gridSpan'), 'val', 1);
      const record = originalProperties && child(originalProperties, 'tcPrChange');
      const recorded = record && recordedProperties(record);
      const historical = recorded
        ? ({ ...originalProperties!, children: recorded } as OoxmlElement)
        : undefined;
      const absorbedSpan = historical
        ? Math.max(0, currentSpan - number(child(historical, 'gridSpan'), 'val', 1))
        : 0;
      // Horizontal merges may already express the deleted neighbour in the current span.
      if (absorbedSpan >= extra.span) return c;
      const historicalWidth = historical && child(historical, 'tcW');
      const absorbedWidth =
        absorbedSpan &&
        historicalWidth &&
        width &&
        revisionAttribute(historicalWidth, 'type') === extra.type
          ? Math.max(0, positiveWidth(width, 0) - positiveWidth(historicalWidth, 0))
          : 0;
      const added = Math.max(0, extra.width - absorbedWidth);
      const takesWidth =
        width &&
        extra.type &&
        ['dxa', 'pct'].includes(extra.type) &&
        revisionAttribute(width, 'type') === extra.type;
      const patched = patch(
        c,
        {
          gridSpan: { val: String(currentSpan + extra.span - absorbedSpan) },
          ...(takesWidth
            ? { tcW: { type: extra.type!, w: String(positiveWidth(width, 0) + added) } }
            : {}),
        },
        mint
      );
      return takesWidth && added > 0 && extra.lent
        ? withRecordTakingCell(
            patched,
            { width, span: currentSpan },
            { width: added, span: extra.span - absorbedSpan },
            mint
          )
        : patched;
    });
    return replaceTableChildren(row, new Map(cells.map((c) => [c.id, c])));
  });
  let above = new Set<string>();
  let promotedAbove = new Set<string>();
  const hasText = (node: OoxmlNode): boolean =>
    node.kind === 'textValue'
      ? false
      : ['text', 'deletedText'].includes(node.kind)
        ? node.children.some((child) => child.kind === 'textValue' && child.value.length > 0)
        : node.children.some(hasText);
  const continuations = rows.map((row) => {
    const result = new Set<string>();
    const pr = child(row, 'trPr');
    let at = number(pr && child(pr, 'gridBefore'), 'val', 0);
    for (const cell of tableChildren(row, 'tableCell')) {
      const pr = child(cell, 'tcPr');
      const end = at + number(pr && child(pr, 'gridSpan'), 'val', 1);
      const merge = pr && child(pr, 'vMerge');
      if (merge && revisionAttribute(merge, 'val') !== 'restart') result.add(`${at}:${end}`);
      at = end;
    }
    return result;
  });
  rows = rows.map((row, rowIndex) => {
    const active = new Set<string>();
    const promoted = new Set<string>();
    const rowPr = child(row, 'trPr');
    let column = number(rowPr && child(rowPr, 'gridBefore'), 'val', 0);
    const cells = tableChildren(row, 'tableCell').map((cell) => {
      if (cell.kind !== 'tableCell') return cell;
      const pr = child(cell, 'tcPr');
      const end = column + number(pr && child(pr, 'gridSpan'), 'val', 1);
      const key = `${column}:${end}`;
      column = end;
      const merge = pr && child(pr, 'vMerge');
      if (!merge) return cell;
      const continuation = revisionAttribute(merge, 'val') !== 'restart';
      if (continuation && !above.has(key)) {
        if (continuations[rowIndex + 1]?.has(key)) {
          active.add(key);
          promoted.add(key);
          return patch(cell, { vMerge: { val: 'restart' } }, mint);
        }
        return patch(cell, { vMerge: null }, mint);
      }
      // Word keeps nonempty legacy continuation content visible after its head is removed.
      if (continuation && promotedAbove.has(key) && !mergedCells.has(cell.id) && hasText(cell))
        return patch(cell, { vMerge: null }, mint);
      active.add(key);
      if (continuation && promotedAbove.has(key)) promoted.add(key);
      if (!continuation || !mergedCells.has(cell.id)) return cell;
      const paragraph = elements(cell).find((n) => n.kind === 'paragraph');
      const empty = paragraph
        ? ({
            ...paragraph,
            children: paragraph.children.filter((n) => isWmlNamed(n, 'pPr')),
          } as OoxmlElement)
        : ({ ...element(cell, 'p', {}, mint), kind: 'paragraph' } as OoxmlElement);
      return {
        ...cell,
        children: [...cell.children.filter((n) => isWmlNamed(n, 'tcPr')), empty],
      } as OoxmlElement;
    });
    above = active;
    promotedAbove = promoted;
    return replaceTableChildren(row, new Map(cells.map((c) => [c.id, c])));
  });
  const byId = new Map(rows.map((row) => [row.id, row]));
  return restoreImplicitRowGrid(
    original,
    replaceTableChildren(rebuilt, byId),
    restoredProperties,
    mint
  );
}

/** The columns a row skips before and after its cells. */
interface RowOffsets {
  readonly before: number;
  readonly after: number;
}

/**
 * A row (or its record) whose `w:trPr` skips `offsets` (`w:gridBefore`, `w:gridAfter`), first
 * among its children, as `patch` puts a cell's `w:tcPr`.
 */
function withRowOffsets(
  owner: OoxmlElement,
  offsets: RowOffsets,
  mint: () => string
): OoxmlElement {
  const pr = child(owner, 'trPr');
  const placed = [
    ...(offsets.before
      ? [element(pr ?? owner, 'gridBefore', { val: String(offsets.before) }, mint)]
      : []),
    ...(offsets.after
      ? [element(pr ?? owner, 'gridAfter', { val: String(offsets.after) }, mint)]
      : []),
  ];
  if (!pr && placed.length === 0) {
    return owner;
  }

  const properties =
    pr?.children.filter((c) => !isWmlNamed(c, 'gridBefore') && !isWmlNamed(c, 'gridAfter')) ?? [];
  const insertion = properties.findIndex((property) => !isWmlNamed(property, 'cnfStyle'));
  properties.splice(insertion < 0 ? properties.length : insertion, 0, ...placed);
  const next = { ...(pr ?? element(owner, 'trPr', {}, mint)), children: properties };
  return {
    ...owner,
    children: [next, ...owner.children.filter((c) => !isWmlNamed(c, 'trPr'))],
  } as OoxmlElement;
}

/** The grid placement a join changes in a container, with each value's default. */
const PLACEMENT = { tcPr: { gridSpan: 1 }, trPr: { gridBefore: 0, gridAfter: 0 } } as const;

const placementOf = (pr: OoxmlElement | undefined, container: keyof typeof PLACEMENT): string =>
  Object.entries(PLACEMENT[container])
    .map(([name, fallback]) => number(pr && child(pr, name), 'val', fallback))
    .join();

/** A row or a cell with its pending property record (`w:trPrChange`, `w:tcPrChange`) updated. */
function withOwnRecord(
  owner: OoxmlElement,
  container: keyof typeof PLACEMENT,
  update: (record: OoxmlElement) => OoxmlElement
): OoxmlElement {
  const pr = child(owner, container);
  const record = pr && child(pr, `${container}Change`);
  if (!pr || !record) {
    return owner;
  }

  const next = update(record);
  return {
    ...owner,
    children: owner.children.map((c) =>
      c === pr ? { ...pr, children: pr.children.map((r) => (r === record ? next : r)) } : c
    ),
  } as OoxmlElement;
}

/**
 * The owner with its pending property record placed by `place` too when the record holds the
 * placement the owner holds: it follows a grid join, as an unchanged grid record does. A record
 * of another placement (a cell's width and span from before a column went) keeps it.
 */
function withRecordFollowing(
  owner: OoxmlElement,
  container: keyof typeof PLACEMENT,
  place: (record: OoxmlElement) => OoxmlElement
): OoxmlElement {
  const shown = placementOf(child(owner, container), container);
  return withOwnRecord(owner, container, (record) =>
    placementOf(child(record, container), container) === shown ? place(record) : record
  );
}

/** The grid's widths, in order. */
const gridWidths = (grid: OoxmlElement): string[] =>
  elements(grid)
    .filter((n) => isWmlNamed(n, 'gridCol'))
    .map((col) => revisionAttribute(col, 'w') ?? '');

/**
 * The grid's pending record (`w:tblGridChange`) when it holds the grid as it is (a row height's
 * anchor, Word's form): it holds no other grid, so a join goes on and the record follows it.
 */
export function unchangedGridRecord(grid: OoxmlElement): OoxmlElement | undefined {
  const record = child(grid, 'tblGridChange');
  const old = record && child(record, 'tblGrid');
  if (!record || !old) {
    return undefined;
  }

  return gridWidths(old).join() === gridWidths(grid).join() ? record : undefined;
}

/** Whether a grid's pending record holds another grid: the grid stays until its decision. */
export const holdsOtherGrid = (grid: OoxmlElement): boolean =>
  child(grid, 'tblGridChange') !== undefined && unchangedGridRecord(grid) === undefined;

/** The grid with other columns; a record that holds the grid as it is holds them too. */
function withColumns(
  grid: OoxmlElement,
  columns: readonly OoxmlElement[],
  mint: () => string
): OoxmlElement {
  const unchanged = unchangedGridRecord(grid);
  const followed = (old: OoxmlNode): OoxmlNode =>
    old.kind !== 'textValue' && isWmlNamed(old, 'tblGrid')
      ? withColumns(
          old,
          columns.map((col) => ({ ...col, id: mint() })),
          mint
        )
      : old;
  const rest = grid.children
    .filter((c) => !isWmlNamed(c, 'gridCol'))
    .map((c) =>
      c === unchanged ? { ...unchanged, children: unchanged.children.map(followed) } : c
    );
  return { ...grid, children: [...columns, ...rest] } as OoxmlElement;
}

/**
 * Remove grid boundaries unused by any surviving cell, with their widths given to the column
 * before, as Word does beside other pending records. A pending grid record that holds another
 * grid keeps the grid; one that holds this grid follows the join.
 */
export function compactRevisionGrid(table: OoxmlElement, mint: () => string): OoxmlElement {
  const grid = child(table, 'tblGrid');
  if (!grid || holdsOtherGrid(grid)) {
    return table;
  }

  if (
    tableChildren(table, 'tableRow').some(
      (row) =>
        !table.children.some((n) => n.id === row.id) ||
        tableChildren(row, 'tableCell').some((cell) => !row.children.some((n) => n.id === cell.id))
    )
  )
    return table;
  const columns = elements(grid).filter((n) => isWmlNamed(n, 'gridCol'));
  if (
    !columns.length ||
    columns.some((col) =>
      col.attributes.some(
        (attr) => attr.namespaceUri !== WML_NAMESPACE_URI || attr.localName !== 'w'
      )
    )
  )
    return table;
  const boundaries = new Set([0]);
  const ranges = new Map<string, [number, number]>();
  const rowRanges = new Map<string, [number, number]>();
  for (const row of table.children) {
    if (row.kind !== 'tableRow') continue;
    const trPr = child(row, 'trPr');
    let at = number(trPr && child(trPr, 'gridBefore'), 'val', 0);
    boundaries.add(at);
    const start = at;
    for (const cell of row.children) {
      if (cell.kind !== 'tableCell') continue;
      const pr = child(cell, 'tcPr');
      const end = at + number(pr && child(pr, 'gridSpan'), 'val', 1);
      ranges.set(cell.id, [at, end]);
      boundaries.add(end);
      at = end;
    }
    rowRanges.set(row.id, [start, at]);
    boundaries.add(at + number(trPr && child(trPr, 'gridAfter'), 'val', 0));
  }
  const sorted = [...boundaries].sort((a, b) => a - b);
  if (sorted.at(-1)! > columns.length || sorted.length < 2) return table;
  if (sorted.length === columns.length + 1) return table;
  const newColumns = sorted.slice(1).map((end, i) =>
    element(
      columns[sorted[i]!]!,
      'gridCol',
      {
        w: String(
          columns.slice(sorted[i], end).reduce((sum, col) => sum + positiveWidth(col, 0), 0)
        ),
      },
      mint
    )
  );
  const positions = new Map(sorted.map((pos, index) => [pos, index]));
  return {
    ...table,
    children: table.children.map((n) => {
      if (n.id === grid.id) {
        return withColumns(grid, newColumns, mint);
      }

      if (n.kind !== 'tableRow') return n;
      const range = rowRanges.get(n.id)!;
      const offsets: RowOffsets = {
        before: positions.get(range[0])!,
        after: sorted.length - 1 - positions.get(range[1])!,
      };
      const followed = withRecordFollowing(n, 'trPr', (record) =>
        withRowOffsets(record, offsets, mint)
      );
      const row = withRowOffsets(followed, offsets, mint);
      return {
        ...row,
        children: row.children.map((c) => {
          const range = ranges.get(c.id);
          if (!range || c.kind !== 'tableCell') {
            return c;
          }

          const span = positions.get(range[1])! - positions.get(range[0])!;
          const value = { gridSpan: span === 1 ? null : { val: String(span) } };
          const cell = withRecordFollowing(c, 'tcPr', (record) => patch(record, value, mint));
          return patch(cell, value, mint);
        }),
      } as OoxmlElement;
    }),
  } as OoxmlElement;
}
