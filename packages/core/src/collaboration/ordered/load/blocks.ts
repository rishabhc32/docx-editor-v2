// A story's blocks as items: paragraphs, tables with their rows and cells, block controls,
// block custom XML, notes. Anything else between blocks (a section's properties, a block
// bookmark, a control around rows) is one closed item, kept whole.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import {
  BLOCK_TOKENS,
  type OrderedContainerToken,
  type OrderedItem,
  type OrderedPropertyRecord,
} from '../model/types.ts';
import {
  attributeOf,
  childNamed,
  isControlContent,
  isWml,
  shellOf,
  showsPlaceholder,
  withChildren,
} from '../ooxml-nodes.ts';
import { loadMoveRangeMarker } from './moves.ts';
import { loadParagraph } from './paragraph.ts';
import { containerShell } from './properties.ts';
import { wordRowHeights } from './row-heights.ts';
import {
  checkDepth,
  contentLabels,
  markerLabels,
  markPrompt,
  NO_LABELS,
  pushClose,
  pushOpen,
  pushZero,
  ROW_ROLES,
  type Labels,
  type LoadState,
} from './state.ts';

/** Load the blocks of a story root or a container. */
export function loadBlocks(
  state: LoadState,
  children: readonly OoxmlNode[],
  depth: number,
  labels: Labels = NO_LABELS
): void {
  checkDepth(state, depth);
  for (const child of children) loadBlock(state, child, labels, depth);
}

const isBlock = (child: OoxmlNode) => holdsBlocks(child, 0);

/**
 * A block: a paragraph, a table, or a control or custom XML element whose content holds blocks
 * (a repeating section holds its items as block controls). Bounded: the depth is the file's.
 */
function holdsBlocks(node: OoxmlNode, depth: number): boolean {
  if (node.kind === 'paragraph' || node.kind === 'table') return true;
  if (node.kind === 'textValue' || depth > MAX_BLOCK_NESTING) return false;
  if (node.kind === 'contentControl' || isWml(node, 'sdt')) {
    const content = node.children.find(isControlContent);
    return (
      content !== undefined &&
      content.kind !== 'textValue' &&
      content.children.some((c) => holdsBlocks(c, depth + 1))
    );
  }

  if (isWml(node, 'customXml')) return node.children.some((c) => holdsBlocks(c, depth + 1));
  return false;
}

/** How deep controls inside controls are looked into for blocks. */
const MAX_BLOCK_NESTING = 16;
const isRow = (child: OoxmlNode) => child.kind === 'tableRow' || isWml(child, 'tr');
const isCell = (child: OoxmlNode) => child.kind === 'tableCell' || isWml(child, 'tc');

function loadBlock(state: LoadState, node: OoxmlNode, labels: Labels, depth: number): void {
  if (node.kind === 'textValue') return;
  if (node.kind === 'paragraph') {
    loadParagraph(state, node, labels, depth + 1);
    return;
  }

  if (node.kind === 'table') {
    loadTable(state, node, labels, depth + 1);
    return;
  }

  if (isRow(node)) {
    loadRow(state, node, labels, depth + 1);
    return;
  }

  if (isCell(node)) {
    loadCell(state, node, labels, depth + 1);
    return;
  }

  if (node.kind === 'contentControl' || isWml(node, 'sdt')) {
    const content = node.children.find(isControlContent);
    if (content && content.kind !== 'textValue' && content.children.some(isBlock)) {
      loadBlockControl(state, node, content, labels, depth + 1);
      return;
    }
  }

  if (isWml(node, 'customXml') && node.children.some(isBlock)) {
    const { shell } = containerShell(state, node, (c) => isWml(c, 'customXmlPr'));
    const content = node.children.filter((c) => !isWml(c, 'customXmlPr'));
    loadBlockContainer(state, node, 'customXmlBlock', shell, content, labels, depth);
    return;
  }

  if (node.kind === 'note' || isWml(node, 'footnote') || isWml(node, 'endnote')) {
    loadBlockContainer(state, node, 'note', shellOf(node), node.children, labels, depth);
    return;
  }

  if (loadMoveRangeMarker(state, node, labels, { block: true })) return;
  pushZero(state, node, labels, { block: true });
}

/** A container of blocks: its open item, its blocks and its close item. */
function loadBlockContainer(
  state: LoadState,
  node: OoxmlElement,
  token: OrderedContainerToken,
  shell: OoxmlElement,
  children: readonly OoxmlNode[],
  labels: Labels,
  depth: number
): void {
  const open = pushOpen(state, node, token, shell, labels);
  loadBlocks(state, children, depth + 1, labels);
  pushClose(state, open, labels);
}

function loadBlockControl(
  state: LoadState,
  node: OoxmlElement,
  content: OoxmlElement,
  labels: Labels,
  depth: number
): void {
  const { shell } = containerShell(state, node, (c) => !isControlContent(c));
  const open = pushOpen(state, node, 'sdtBlock', shell, labels);
  const first = state.items.length;
  loadBlocks(state, content.children, depth + 1, labels);

  // A block control showing its prompt: the prompt's units are not content
  if (showsPlaceholder(node)) {
    markPrompt(state, first, (item) => item.kind === 'text' || item.kind === 'atom');

    // A prompt of two or more paragraphs: their marks are the prompt's too, and the paragraphs
    // that hold nothing else go while content replaces it (`change/settle-prompts.ts`)
    const marks = ownMarks(state.items, first);
    if (marks.length > 1) for (const k of marks) state.prompts.add(k);
  }

  pushClose(state, open, labels);
}

/** The indexes of the marks from `first` on that no container inside the control holds. */
function ownMarks(items: readonly OrderedItem[], first: number): number[] {
  const out: number[] = [];
  let depth = 0;
  for (let k = first; k < items.length; k += 1) {
    const item = items[k];
    if (!item) continue;
    if (item.token !== undefined && BLOCK_TOKENS.has(item.token)) {
      if (item.kind === 'open') depth += 1;
      else if (item.kind === 'close') depth -= 1;
    } else if (item.kind === 'mark' && depth === 0) out.push(k);
  }

  return out;
}

const isTableProperties = (c: OoxmlNode) =>
  c.kind === 'tableProperties' ||
  c.kind === 'tableGrid' ||
  isWml(c, 'tblPr') ||
  isWml(c, 'tblGrid');

/** A properties element's change record of `name` under the `w:id` `id`, with an author. */
function recordOf(properties: OoxmlNode, name: string, id: string): OoxmlElement | undefined {
  const found = properties.kind === 'textValue' ? undefined : childNamed(properties, name);
  return found && attributeOf(found, 'id') === id && attributeOf(found, 'author') !== undefined
    ? found
    : undefined;
}

/**
 * The record whose suggestion a table's `w:tblGridChange` joins: one of the table's own records
 * (its `w:tblPrChange`, a row's or a cell's) with the same `w:id` and an author. A grid record
 * is `CT_Markup`, an id alone. A resize writes its records under one revision and the engine
 * decides them as one gesture (its review card joins them the same way); a grid record under an
 * id of its own, as Word writes it, stays a suggestion of its own.
 */
function gridGestureOf(table: OoxmlElement): OoxmlElement | undefined {
  const grid = table.children.find((c) => isWml(c, 'tblGrid'));
  const record = grid && grid.kind !== 'textValue' ? childNamed(grid, 'tblGridChange') : undefined;
  const id = record && attributeOf(record, 'id');
  if (id === undefined) return undefined;
  for (const child of table.children) {
    if (isWml(child, 'tblPr')) {
      const found = recordOf(child, 'tblPrChange', id);
      if (found) return found;
    }

    if (child.kind === 'textValue' || !isRow(child)) continue;
    for (const part of child.children) {
      const found = part.kind === 'textValue' ? undefined : rowPartRecord(part, id);
      if (found) {
        return found;
      }
    }
  }

  return undefined;
}

/** The change record under `id` of a row's part: its `w:trPr`, its `w:tblPrEx` or a cell. */
function rowPartRecord(part: OoxmlElement, id: string): OoxmlElement | undefined {
  if (isWml(part, 'trPr')) {
    return recordOf(part, 'trPrChange', id);
  }

  if (isWml(part, 'tblPrEx')) {
    return recordOf(part, 'tblPrExChange', id);
  }

  if (isCell(part)) {
    return part.children
      .filter((c) => isWml(c, 'tcPr'))
      .map((c) => recordOf(c, 'tcPrChange', id))
      .find((c) => c !== undefined);
  }

  return undefined;
}

function loadTable(state: LoadState, node: OoxmlElement, labels: Labels, depth: number): void {
  checkDepth(state, depth);

  // Word's row heights with no row record: one suggestion over the table's and cells' records
  const heights = new Map<OoxmlElement, AddedHeight>();
  for (const change of wordRowHeights(node)) {
    const suggestion = state.revisions.suggestionForAll(
      change.records,
      REVISION_ROLE.tableProperties
    );
    for (const [row, trPr] of change.added) {
      // A height on its row's own anchor: the row's `w:tblPrEx` is there for the record alone,
      // and compose writes it again while the height is pending (`compose/row-heights.ts`)
      const tblPrEx = childNamed(row, 'tblPrEx');
      const record = tblPrEx && childNamed(tblPrEx, 'tblPrExChange');
      const own = record && change.records.includes(record) ? { tblPrEx } : {};
      heights.set(row, { suggestion, trPr, ...own });
    }
  }

  const { shell, props } = containerShell(
    state,
    node,
    (c) => !isRow(c) && isTableProperties(c),
    undefined,
    gridGestureOf(node)
  );
  const open = pushOpen(state, node, 'tbl', shell, labels, props);
  for (const child of node.children) {
    if (child.kind === 'textValue' || isTableProperties(child)) continue;
    if (isRow(child)) {
      loadRow(state, child, labels, depth + 1, heights.get(child));
    } else loadBlock(state, child, labels, depth);
  }

  pushClose(state, open, labels);
}

/** A height Word added with no row record (`wordRowHeights`), and its row's own anchor. */
interface AddedHeight {
  readonly suggestion: string;
  readonly trPr: OoxmlElement;
  /** The row's `w:tblPrEx` that holds the height's own record: left out of the row. */
  readonly tblPrEx?: OoxmlElement;
}

function loadRow(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  depth: number,
  height?: AddedHeight
): void {
  checkDepth(state, depth);
  const { own: rowLabels, content } = containerLabels(state, node, 'trPr', labels);
  const { shell, props } = containerShell(
    state,
    node,
    (c) =>
      !isCell(c) &&
      ((isWml(c, 'trPr') && c !== height?.trPr) || (isWml(c, 'tblPrEx') && c !== height?.tblPrEx)),
    'trPr'
  );
  const added = height ? [addedHeightRecord(state, height)] : [];
  const open = pushOpen(state, node, 'tr', shell, rowLabels, [...props, ...added]);
  for (const child of node.children) {
    if (child.kind === 'textValue' || isWml(child, 'trPr') || isWml(child, 'tblPrEx')) continue;
    if (isCell(child)) loadCell(state, child, content, depth + 1);
    else loadBlock(state, child, content, depth);
  }

  pushClose(state, open, rowLabels);
}

/**
 * A height Word added with no row record: the row had no `w:trPr`, and the height shows while
 * its suggestion is pending.
 */
function addedHeightRecord(state: LoadState, height: AddedHeight): OrderedPropertyRecord {
  return {
    kind: 'element',
    suggestion: height.suggestion,
    seq: state.revisions.seqOf(height.suggestion),
    target: 'trPr',
    // Without the row's own markers, which label the row
    element: withChildren(
      height.trPr,
      height.trPr.children.filter((c) => !isWml(c, 'ins') && !isWml(c, 'del'))
    ),
    untracked: true,
  };
}

/** A row's or a cell's labels from its own markers, and the labels of what it holds. */
function containerLabels(
  state: LoadState,
  node: OoxmlElement,
  name: 'trPr' | 'tcPr',
  labels: Labels
): { readonly own: Labels; readonly content: Labels } {
  const found = node.children.find((c) => isWml(c, name));
  const properties = found && found.kind !== 'textValue' ? found : undefined;
  return {
    own: markerLabels(state, properties, labels, ROW_ROLES),
    // What it holds goes with it: its own markers label what is inside
    content: contentLabels(labels, markerLabels(state, properties, NO_LABELS, ROW_ROLES)),
  };
}

function loadCell(state: LoadState, node: OoxmlElement, labels: Labels, depth: number): void {
  checkDepth(state, depth);
  const { own: cellLabels, content } = containerLabels(state, node, 'tcPr', labels);
  const { shell, props } = containerShell(state, node, (c) => isWml(c, 'tcPr'), 'tcPr');
  const open = pushOpen(state, node, 'tc', shell, cellLabels, props);
  loadBlocks(
    state,
    node.children.filter((c) => !isWml(c, 'tcPr')),
    depth + 1,
    content
  );
  pushClose(state, open, cellLabels);
}
