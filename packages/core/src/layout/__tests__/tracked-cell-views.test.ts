// A proposed column in the views that resolve tracked changes.
//
// Suggesting mode proposes deleting a column as `w:cellDel` on each of its cells, and inserting
// one as `w:cellIns`. No Markup and Simple Markup lay out the document as accepting every change
// would leave it, and Original as rejecting every change would; a tracked row already leaves
// those views, but its cells kept painting, so No Markup still drew a deleted column. Each view
// must now lay the table out exactly as the Accept or Reject it stands for: the same cells, at
// the same place, at the same width.

import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '@docx-editor.dev/core/store';
import { applyTreeOp } from '../../store/store/tree-ops.ts';
import type { TreeDocOp } from '../../store/store/tree-op-types.ts';
import { revisionAuthorFilter, type RevisionDisplayMode } from '../revision-projection.ts';
import { readTableStructure } from '../semantic-table.ts';
import { createFixedMeasurer, layoutSemanticDocument } from '../semantic-layout.ts';
import type { SemanticLayout } from '../semantic-records.ts';

const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const cell = (text: string, width: number): string =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(text)}</w:tc>`;
const row = (prefix: string): string =>
  `<w:tr>${cell(`${prefix}1`, 2000)}${cell(`${prefix}2`, 3000)}${cell(`${prefix}3`, 4000)}</w:tr>`;
/** Three columns of 2000, 3000 and 4000 twips over two rows, between two paragraphs. */
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/><w:gridCol w:w="4000"/></w:tblGrid>' +
  `${row('a')}${row('b')}</w:tbl>` +
  p('after');

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${WML_NAMESPACE_URI}"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function elements(root: OoxmlNode, matches: (node: OoxmlElement) => boolean): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (matches(node)) found.push(node);
    for (const child of node.children) visit(child);
  };
  visit(root);
  return found;
}

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op} refused: ${result.reason}`);
  return result.part;
}

/** The part with Bea proposing a column operation on grid column `index`. */
function propose(kind: 'delete' | 'insert', index: number): OoxmlPart {
  const part = load(BODY);
  const tableId = elements(part.root, (node) => node.kind === 'table')[0]!.id;
  const gridColumnId = elements(part.root, (node) => node.localName === 'gridCol')[index]!.id;
  const revision = { author: 'Bea', date: '2026-10-01T00:00:00Z' };
  return apply(
    part,
    kind === 'delete'
      ? { op: 'deleteTableColumn', tableId, gridColumnId, revision }
      : { op: 'insertTableColumn', tableId, gridColumnId, where: 'right', revision }
  );
}

const resolveAll = (part: OoxmlPart, action: 'accept' | 'reject'): OoxmlPart =>
  apply(part, { op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions' });

const measurer = createFixedMeasurer(6, 14);
const lay = (
  part: OoxmlPart,
  displayMode: RevisionDisplayMode,
  hiddenAuthors: readonly string[] = []
): SemanticLayout =>
  layoutSemanticDocument(part, 0, {
    measurer,
    displayMode,
    revisionAuthorFilter: revisionAuthorFilter(hiddenAuthors),
  });

/** Every laid-out cell, row by row: which cell it is, where it starts, how wide it is. */
function cellGeometry(layout: SemanticLayout) {
  return layout.pages
    .flatMap((page) => page.fragments)
    .flatMap((block) =>
      block.kind === 'table'
        ? block.rows.map((laidRow) =>
            laidRow.cells.map((laid) => ({
              id: laid.id,
              x: laid.box.x,
              width: laid.box.width,
              shading: laid.revisionShading,
            }))
          )
        : []
    );
}

const cellCount = (layout: SemanticLayout): number => cellGeometry(layout).flat().length;

describe('a proposed column deletion', () => {
  const proposed = propose('delete', 1);
  const accepted = resolveAll(proposed, 'accept');

  test('No Markup lays out the table Accept leaves; All Markup keeps the column, shaded', () => {
    expect(cellGeometry(lay(proposed, 'proposed'))).toEqual(
      cellGeometry(lay(accepted, 'all-markup'))
    );
    expect(cellCount(lay(proposed, 'proposed'))).toBe(4);
    const markup = cellGeometry(lay(proposed, 'all-markup'));
    expect(markup.flat()).toHaveLength(6);
    expect(markup.map((cells) => cells.map((laid) => laid.shading))).toEqual([
      [undefined, 'deleted', undefined],
      [undefined, 'deleted', undefined],
    ]);
  });

  test('a reviewer filter that hides the proposer resolves the deletion in All Markup', () => {
    expect(cellGeometry(lay(proposed, 'all-markup', ['Bea']))).toEqual(
      cellGeometry(lay(accepted, 'all-markup'))
    );
  });

  test('Original lays out the table Reject leaves, which is the authored one', () => {
    const rejected = resolveAll(proposed, 'reject');
    expect(cellGeometry(lay(proposed, 'original'))).toEqual(
      cellGeometry(lay(rejected, 'all-markup'))
    );
    // Guards the comparison above: the hidden view and the authored table differ.
    expect(cellGeometry(lay(proposed, 'proposed'))).not.toEqual(
      cellGeometry(lay(rejected, 'all-markup'))
    );
  });
});

describe('a proposed column insertion', () => {
  const proposed = propose('insert', 1);
  const rejected = resolveAll(proposed, 'reject');

  test('Original lays out the table Reject leaves; No Markup keeps the new column', () => {
    expect(cellGeometry(lay(proposed, 'original'))).toEqual(
      cellGeometry(lay(rejected, 'all-markup'))
    );
    expect(cellCount(lay(proposed, 'original'))).toBe(6);
    expect(cellGeometry(lay(proposed, 'proposed'))).toEqual(
      cellGeometry(lay(resolveAll(proposed, 'accept'), 'all-markup'))
    );
    const markup = cellGeometry(lay(proposed, 'all-markup'));
    expect(markup.map((cells) => cells.map((laid) => laid.shading))).toEqual([
      [undefined, undefined, 'inserted', undefined],
      [undefined, undefined, 'inserted', undefined],
    ]);
  });
});

describe('the resolved table is a view, not an edit', () => {
  test('the stored tree keeps every cell, and the view edits no column', () => {
    const proposed = propose('delete', 1);
    const before = JSON.stringify(proposed.root);
    const table = elements(proposed.root, (node) => node.kind === 'table')[0]!;
    const structure = readTableStructure(table, 468, 0, undefined, 'proposed')!;
    expect(structure.rows.map((laidRow) => laidRow.cells.length)).toEqual([2, 2]);
    expect(structure.rows.flatMap((laidRow) => laidRow.cells.map((c) => c.gridColumnId))).toEqual([
      undefined,
      undefined,
      undefined,
      undefined,
    ]);
    expect(JSON.stringify(proposed.root)).toBe(before);
    // All Markup reads the authored grid, so its columns stay addressable.
    const markup = readTableStructure(table, 468, 0, undefined, 'all-markup')!;
    expect(markup.rows[0]!.cells.every((c) => c.gridColumnId !== undefined)).toBe(true);
  });

  // Paragraph layout, its caches and incremental reflow key on node identity, so a view that
  // laid out copies would lay out every paragraph of the table again on each keystroke.
  const resolvingViews = [
    { view: 'No Markup over a column deletion', kind: 'delete', mode: 'proposed', hidden: [] },
    { view: 'Original over a column insertion', kind: 'insert', mode: 'original', hidden: [] },
    { view: 'a filter hiding the proposer', kind: 'delete', mode: 'all-markup', hidden: ['Bea'] },
  ] as const;
  for (const { view, kind, mode, hidden } of resolvingViews) {
    test(`${view} lays out the authored paragraphs, before and after an edit`, () => {
      const proposed = propose(kind, 1);
      const tableOf = (part: OoxmlPart) => elements(part.root, (node) => node.kind === 'table')[0]!;
      const filter = revisionAuthorFilter(hidden);
      const laidBlocks = (part: OoxmlPart) =>
        readTableStructure(tableOf(part), 468, 0, undefined, mode, filter)!.rows.flatMap(
          (laidRow) => laidRow.cells.flatMap((laid) => laid.blocks)
        );
      const authored = elements(tableOf(proposed), (node) => node.kind === 'paragraph');
      const before = laidBlocks(proposed);
      // Guards the check below: the view resolved the column away.
      expect(before.length).toBeLessThan(authored.length);
      expect(before.every((block) => authored.includes(block))).toBe(true);
      // An edit in one cell re-reads the table; every other paragraph is still the same node.
      const paragraphId = before[0]!.id;
      const after = laidBlocks(
        apply(proposed, { op: 'insertText', paragraphId, offset: 0, text: 'x' })
      );
      expect(after.map((block) => block.id)).toEqual(before.map((block) => block.id));
      expect(after.filter((block, index) => block !== before[index])).toEqual([after[0]!]);
    });
  }
});
