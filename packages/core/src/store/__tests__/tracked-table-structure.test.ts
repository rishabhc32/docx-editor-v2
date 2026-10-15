// Suggesting mode proposes column insertion and deletion in Word's tracked form.
//
// A tracked column edit marks the column's cells — `w:cellDel` or `w:cellIns` — and keeps the
// table's width: the left neighbour takes a deleted column's width on Accept, and lends a new
// column its width, so Reject restores the table exactly. Before, both column ops changed the
// document outright in suggesting mode, with no card to answer.
//
// The last part covers the review side: a decision that removes a row, cell or table takes its
// author's own pending changes inside with it, and stays refused for another author's.

import { describe, expect, test } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  readOoxmlPart,
  serializeOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { diffSemanticDigests, semanticDigest } from '../package/ooxml-digest.ts';
import { applyTreeOp, validateTreeOp } from '../store/tree-ops.ts';
import { TreeDocumentStore } from '../store/tree-store.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { planRevisionBatch } from '../store/revision-batch.ts';
import { revisionItemsOf } from '../store/review-reads.ts';
import { reviewItemKey } from '../store/review-items.ts';
import { holdsPendingChange } from '../store/revision-contained-changes.ts';

const W = WML_NAMESPACE_URI;
const BEA = { author: 'Bea', date: '2026-10-01T00:00:00Z' };
const STAMP = 'w:author="Bea" w:date="2026-10-01T00:00:00Z"';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function collectByKind(root: OoxmlNode, kind: OoxmlElement['kind']): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === kind) found.push(node);
    for (const child of node.children) visit(child);
  };
  visit(root);
  return found;
}

const p = (text: string): string => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const CELL = (text: string, width = 3000, extra = ''): string =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${extra}</w:tcPr>${p(text)}</w:tc>`;
const ROW = (...cells: string[]): string => `<w:tr>${cells.join('')}</w:tr>`;
const GRID = (widths: readonly number[], history = ''): string =>
  `<w:tblGrid>${widths.map((width) => `<w:gridCol w:w="${width}"/>`).join('')}${history}</w:tblGrid>`;
/** Three columns of 2000, 3000 and 4000 twips over two rows, between two paragraphs. */
const BODY = (history = ''): string =>
  p('before') +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${GRID([2000, 3000, 4000], history)}` +
  ROW(CELL('a1', 2000), CELL('a2', 3000), CELL('a3', 4000)) +
  ROW(CELL('b1', 2000), CELL('b2', 3000), CELL('b3', 4000)) +
  `</w:tbl>${p('after')}`;

const tableId = (part: OoxmlPart): string => collectByKind(part.root, 'table')[0]!.id;
const gridColumnId = (part: OoxmlPart, index: number): string =>
  collectByKind(part.root, 'generic').filter((node) => node.localName === 'gridCol')[index]!.id;
const tableXml = (part: OoxmlPart): string =>
  serializeOoxmlPart(part).match(/<w:tbl>.*<\/w:tbl>/)![0];
const widths = (xml: string, element: 'gridCol' | 'tcW'): number[] =>
  [...xml.matchAll(new RegExp(`<w:${element}\\b[^>]*w:w="(\\d+)"`, 'g'))].map((m) => Number(m[1]));

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op} refused: ${result.reason}`);
  return result.part;
}

const deleteColumn = (part: OoxmlPart, index: number, revision = BEA): TreeDocOp => ({
  op: 'deleteTableColumn',
  tableId: tableId(part),
  gridColumnId: gridColumnId(part, index),
  revision,
});
const insertColumn = (part: OoxmlPart, index: number, where: 'left' | 'right'): TreeDocOp => ({
  op: 'insertTableColumn',
  tableId: tableId(part),
  where,
  gridColumnId: gridColumnId(part, index),
  revision: BEA,
});

/** Resolve every pending change, as Accept All and Reject All do. */
const resolveAll = (part: OoxmlPart, action: 'accept' | 'reject'): OoxmlPart =>
  apply(part, { op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions' });

/** The document before and after a save and reopen read the same. */
function expectRoundTrip(part: OoxmlPart): void {
  const reopened = load(serializeOoxmlPart(part).replace(/^.*<w:body>|<\/w:body>.*$/g, ''));
  expect(canonicalOoxmlFingerprint(reopened.root)).toBe(canonicalOoxmlFingerprint(part.root));
  expect(diffSemanticDigests(semanticDigest([part]), semanticDigest([reopened]))).toEqual([]);
}

describe('a tracked column deletion marks the column', () => {
  test('every cell of the column gains w:cellDel under one id; text and grid stay', () => {
    const part = load(BODY());
    const next = apply(part, deleteColumn(part, 1));
    const xml = tableXml(next);
    expect(xml.match(new RegExp(`<w:cellDel ${STAMP} w:id="0"/>`, 'g'))).toHaveLength(2);
    expect(xml).not.toMatch(/<w:del\b/);
    expect(xml).toContain('<w:t>a2</w:t>');
    expect(widths(xml, 'gridCol')).toEqual([2000, 3000, 4000]);
    expectRoundTrip(next);
    // Without an attribution the op still removes the column outright, as editing mode does.
    const { revision: _revision, ...untracked } = deleteColumn(part, 1) as Extract<
      TreeDocOp,
      { op: 'deleteTableColumn' }
    >;
    expect(widths(tableXml(apply(part, untracked)), 'gridCol')).toEqual([2000, 4000]);
  });

  test('Reject restores the table exactly', () => {
    const part = load(BODY());
    const rejected = resolveAll(apply(part, deleteColumn(part, 1)), 'reject');
    expect(serializeOoxmlPart(rejected)).toBe(serializeOoxmlPart(part));
  });

  test('Accept removes the cells, and the left neighbour takes their width', () => {
    const part = load(BODY());
    const xml = tableXml(resolveAll(apply(part, deleteColumn(part, 1)), 'accept'));
    expect(xml).not.toContain('a2');
    expect(xml).not.toMatch(/<w:(cellDel|gridSpan)\b/);
    expect(widths(xml, 'gridCol')).toEqual([5000, 4000]);
    expect(widths(xml, 'tcW')).toEqual([5000, 4000, 5000, 4000]);
  });

  test('the first column gives its width to the column after it', () => {
    const part = load(BODY());
    const xml = tableXml(resolveAll(apply(part, deleteColumn(part, 0)), 'accept'));
    expect(widths(xml, 'gridCol')).toEqual([5000, 4000]);
    expect(xml).not.toContain('a1');
  });

  test('columns one gesture deletes share one id; the next press is a decision of its own', () => {
    const press = (store: TreeDocumentStore, ...columns: number[]): void => {
      const ops = columns.map((index) => deleteColumn(store.part, index));
      expect(store.transact((tx) => ops.forEach((op) => tx.apply(op))).ok).toBe(true);
    };
    // Delete Columns over a rectangle writes one op per column, in one transaction.
    const gesture = new TreeDocumentStore(load(BODY()));
    press(gesture, 0, 1);
    expect(tableXml(gesture.part).match(/<w:cellDel [^>]*w:id="0"\/>/g)).toHaveLength(4);
    expect(revisionItemsOf(gesture.part)).toHaveLength(1);
    // Two presses in the same second share their author and date, and are still two decisions.
    const presses = new TreeDocumentStore(load(BODY()));
    press(presses, 0);
    press(presses, 1);
    expect(revisionItemsOf(presses.part)).toHaveLength(2);
  });

  test('a cell already proposed for deletion keeps its decision', () => {
    const theirs = '<w:cellDel w:id="7" w:author="Ada"/>';
    const part = load(BODY().replace(CELL('a2', 3000), CELL('a2', 3000, theirs)));
    const xml = tableXml(apply(part, deleteColumn(part, 1)));
    expect(xml).toContain('<w:cellDel w:author="Ada" w:id="7"/>');
    expect(xml.match(new RegExp(`<w:cellDel ${STAMP}`, 'g'))).toHaveLength(1);
  });

  test('shapes a proposal cannot represent are refused, not approximated', () => {
    // Another author's pending cell insertion in the column.
    const theirs = '<w:cellIns w:id="7" w:author="Ada"/>';
    const pending = load(BODY().replace(CELL('b2', 3000), CELL('b2', 3000, theirs)));
    expect(validateTreeOp(pending, deleteColumn(pending, 1))).toBe('unsupported-revision');
    // Merged cells, which the untracked column ops refuse too.
    const merged = '<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc>';
    const spanned = load(BODY().replace(CELL('b1', 2000) + CELL('b2', 3000), merged));
    expect(validateTreeOp(spanned, deleteColumn(spanned, 1))).toBe('table-has-merge');
  });

  test('a row another author proposed inserting or deleting is refused, not held', () => {
    for (const kind of ['ins', 'del']) {
      const theirs = `<w:tr><w:trPr><w:${kind} w:id="7" w:author="Ada"/></w:trPr>`;
      const part = load(BODY().replace('</w:tr><w:tr>', `</w:tr>${theirs}`));
      // A mark in Ada's row held her decision: Reject of the insertion, or Accept of the
      // deletion, skipped it as an `incomplete-group`.
      expect(validateTreeOp(part, deleteColumn(part, 1))).toBe('unsupported-revision');
    }
  });

  test('pending deletions cannot keep the table alive for the last column', () => {
    const part = load(BODY());
    const two = apply(apply(part, deleteColumn(part, 0)), deleteColumn(part, 1));
    expect(widths(tableXml(two), 'gridCol')).toHaveLength(3);
    expect(validateTreeOp(two, deleteColumn(two, 2))).toBe('block-required');
  });
});

describe('a tracked column insertion lends the new column its width', () => {
  test('the new cells carry w:cellIns, and the left neighbour lends half its width', () => {
    const part = load(BODY());
    const next = apply(part, insertColumn(part, 1, 'right'));
    const xml = tableXml(next);
    expect(xml.match(new RegExp(`<w:cellIns ${STAMP} w:id="0"/>`, 'g'))).toHaveLength(2);
    expect(widths(xml, 'gridCol')).toEqual([2000, 1500, 1500, 4000]);
    expect(widths(xml, 'tcW')).toEqual([2000, 1500, 1500, 4000, 2000, 1500, 1500, 4000]);
    expectRoundTrip(next);
  });

  test('left of a column, the column before it lends; left of the first, the first does', () => {
    const part = load(BODY());
    expect(widths(tableXml(apply(part, insertColumn(part, 1, 'left'))), 'gridCol')).toEqual([
      1000, 1000, 3000, 4000,
    ]);
    expect(widths(tableXml(apply(part, insertColumn(part, 0, 'left'))), 'gridCol')).toEqual([
      1000, 1000, 3000, 4000,
    ]);
  });

  test.each([
    ['right of the middle column', 1, 'right'],
    ['left of the middle column', 1, 'left'],
    ['left of the first column', 0, 'left'],
    ['right of the last column', 2, 'right'],
  ] as const)('Reject restores the table exactly: %s', (_name, index, where) => {
    const part = load(BODY());
    const rejected = resolveAll(apply(part, insertColumn(part, index, where)), 'reject');
    expect(serializeOoxmlPart(rejected)).toBe(serializeOoxmlPart(part));
  });

  test('Accept keeps the column, unmarked, in a table of the same width', () => {
    const part = load(BODY());
    const xml = tableXml(resolveAll(apply(part, insertColumn(part, 1, 'right')), 'accept'));
    expect(xml).not.toMatch(/cellIns/);
    expect(widths(xml, 'gridCol')).toEqual([2000, 1500, 1500, 4000]);
  });

  test('a grid history keeps matching the grid, so Reject still restores it exactly', () => {
    // Google Docs writes this no-op record into every table it exports.
    const history = `<w:tblGridChange w:id="9">${GRID([2000, 3000, 4000])}</w:tblGridChange>`;
    const part = load(BODY(history));
    const next = apply(part, insertColumn(part, 1, 'right'));
    expect(tableXml(next)).toContain(
      `${GRID([2000, 1500, 1500, 4000]).replace('</w:tblGrid>', '')}<w:tblGridChange w:id="9">` +
        `${GRID([2000, 1500, 1500, 4000])}</w:tblGridChange></w:tblGrid>`
    );
    expect(serializeOoxmlPart(resolveAll(next, 'reject'))).toBe(
      serializeOoxmlPart(resolveAll(part, 'reject'))
    );
  });

  test('deleting a column the author proposed withdraws it, restoring the table', () => {
    const part = load(BODY());
    const proposed = apply(part, insertColumn(part, 1, 'right'));
    expect(tableXml(proposed)).toContain('<w:cellIns ');
    const withdrawn = apply(proposed, deleteColumn(proposed, 2));
    expect(serializeOoxmlPart(withdrawn)).toBe(serializeOoxmlPart(part));
  });

  test("a column crossing the author's own proposed row withdraws that cell, proposing the rest", () => {
    const own = `<w:cellIns w:id="1" ${STAMP}/>`;
    const proposedRow =
      `<w:tr><w:trPr><w:ins w:id="1" ${STAMP}/></w:trPr>` +
      `${CELL('b1', 2000, own)}${CELL('b2', 3000, own)}${CELL('b3', 4000, own)}</w:tr>`;
    const part = load(BODY().replace(/<w:tr>(?:(?!<w:tr>).)*b1.*?<\/w:tr>/, proposedRow));
    const xml = tableXml(apply(part, deleteColumn(part, 1)));
    expect(xml.match(new RegExp(`<w:cellDel ${STAMP}`, 'g'))).toHaveLength(1);
    expect(xml).toContain('a2');
    expect(xml).not.toContain('b2');
    // The withdrawn cell's width went back to the cell before it, as rejecting it does.
    expect(xml).toMatch(/<w:tcW w:type="dxa" w:w="5000"\/><w:gridSpan w:val="2"\/>/);
  });

  test("a proposed column holding another author's pending change is not withdrawn", () => {
    const part = load(BODY());
    const proposed = apply(part, insertColumn(part, 1, 'right'));
    const cell = collectByKind(proposed.root, 'tableRow')[0]!.children.filter(
      (child) => child.kind === 'tableCell'
    )[2]!;
    const paragraph = collectByKind(cell, 'paragraph')[0]!;
    const typed = apply(proposed, {
      op: 'insertText',
      paragraphId: paragraph.id,
      offset: 0,
      text: 'x',
      revision: { author: 'Ada' },
    });
    expect(validateTreeOp(typed, deleteColumn(typed, 2))).toBe('unsupported-revision');
  });

  test.each([
    ['a table without w:tblGrid', BODY().replace(/<w:tblGrid>.*<\/w:tblGrid>/, '')],
    [
      'a lender with a cell-property record',
      BODY().replace(
        CELL('a2', 3000),
        CELL('a2', 3000, `<w:tcPrChange w:id="7" w:author="Ada"><w:tcPr/></w:tcPrChange>`)
      ),
    ],
    [
      'a lender width with no type',
      BODY().replace('<w:tcW w:w="3000" w:type="dxa"/>', '<w:tcW w:w="3000"/>'),
    ],
  ])('%s is refused', (_name, body) => {
    const part = load(body);
    const grid = collectByKind(part.root, 'generic').some((node) => node.localName === 'gridCol');
    const reference = collectByKind(part.root, 'tableCell')[1]!.id;
    const op: TreeDocOp = grid
      ? insertColumn(part, 1, 'right')
      : {
          op: 'insertTableColumn',
          tableId: tableId(part),
          where: 'right',
          referenceCellId: reference,
          revision: BEA,
        };
    expect(validateTreeOp(part, op)).toBe('unsupported-revision');
  });
});

describe('a structural decision takes its author’s own changes inside with it', () => {
  // Bea proposed the row and typed into it.
  const proposedRow = (typist: string): string =>
    `<w:tr><w:trPr><w:ins w:id="1" ${STAMP}/></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns w:id="1" ${STAMP}/></w:tcPr>` +
    `<w:p><w:ins w:id="2" w:author="${typist}" w:date="2026-10-01T00:00:05Z"><w:r><w:t>typed</w:t></w:r></w:ins></w:p></w:tc></w:tr>`;
  const body = (typist: string): string =>
    `<w:tbl>${GRID([2000])}${ROW(CELL('kept', 2000))}${proposedRow(typist)}</w:tbl>${p('after')}`;
  const rowCard = (part: OoxmlPart) =>
    revisionItemsOf(part).find((item) => item.revisionKind === 'structural')!;

  test('rejecting a proposed row also rejects its author’s typing; accepting leaves it', () => {
    const part = load(body('Bea'));
    expect(revisionItemsOf(part)).toHaveLength(2);
    const rejected = planRevisionBatch(part, 'reject', [reviewItemKey(rowCard(part))]);
    expect(rejected.result.skipped).toEqual([]);
    expect(rejected.result.resolved).toHaveLength(2);
    const resolved = rejected.ops.reduce(apply, part);
    expect(serializeOoxmlPart(resolved)).not.toContain('typed');
    expect(revisionItemsOf(resolved)).toHaveLength(0);
    // Accepting removes nothing, so the typing stays a decision of its own.
    const accepted = planRevisionBatch(part, 'accept', [reviewItemKey(rowCard(part))]);
    expect(accepted.result.resolved).toHaveLength(1);
    const kept = accepted.ops.reduce(apply, part);
    expect(revisionItemsOf(kept).map((item) => item.revisionKind)).toEqual(['insert']);
  });

  test("only another author's typing holds the decision, and says why", () => {
    const own = load(body('Bea'));
    expect(planRevisionBatch(own, 'reject', [reviewItemKey(rowCard(own))]).result.skipped).toEqual(
      []
    );
    expect(holdsPendingChange(own, rowCard(own), 'reject')).toBe(false);
    const theirs = load(body('Ada'));
    const plan = planRevisionBatch(theirs, 'reject', [reviewItemKey(rowCard(theirs))]);
    expect(plan.result.skipped.map((entry) => entry.reason)).toEqual(['incomplete-group']);
    expect(holdsPendingChange(theirs, rowCard(theirs), 'reject')).toBe(true);
  });

  test('accepting a proposed column deletion also accepts its author’s changes in it', () => {
    const part = load(BODY());
    const deleted = apply(part, deleteColumn(part, 1));
    const cell = collectByKind(deleted.root, 'tableCell')[1]!;
    const typed = apply(deleted, {
      op: 'insertText',
      paragraphId: collectByKind(cell, 'paragraph')[0]!.id,
      offset: 0,
      text: 'x',
      revision: BEA,
    });
    const card = revisionItemsOf(typed).find((item) => item.revisionKind === 'structural')!;
    const plan = planRevisionBatch(typed, 'accept', [reviewItemKey(card)]);
    expect(plan.result.skipped).toEqual([]);
    const xml = tableXml(plan.ops.reduce(apply, typed));
    expect(xml).not.toMatch(/a2|x<|cellDel|<w:ins\b/);
  });
});
