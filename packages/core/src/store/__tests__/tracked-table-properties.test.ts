// Suggesting mode records row and table property changes in Word's tracked form.
//
// A row height or header row keeps the new `w:trPr` and records the old one in a
// `w:trPrChange`; a table style records `w:tblPrChange`; a column resize records the old grid in
// a `w:tblGridChange`, the old cell widths in `w:tcPrChange`, and the old layout in a
// `w:tblPrChange`. Before, each landed in suggesting mode as a permanent edit with nothing to
// accept or reject.

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
import { planRevisionBatch, revisionItemsOf } from '../index.ts';
import { reviewItemKey } from '../store/review-items.ts';
import { applyTreeOp, validateTreeOp } from '../store/tree-ops.ts';
import { OTHER_AUTHORS_ROW_CHANGE } from '../store/tree-op-tracked-table-properties.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = WML_NAMESPACE_URI;
const BEA = { author: 'Bea', date: '2026-10-01T00:00:00Z' };
const STAMP = 'w:author="Bea" w:date="2026-10-01T00:00:00Z"';
const TC_W = '<w:tcW w:w="2400" w:type="dxa"/>';
const NAME = '/word/document.xml';
const CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: NAME,
    contentType: CONTENT_TYPE,
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

const CELL = (text: string, properties = TC_W): string =>
  `<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const ROW = (...cells: string[]): string => `<w:tr>${cells.join('')}</w:tr>`;
const ROW_WITH = (properties: string, ...cells: string[]): string =>
  `<w:tr><w:trPr>${properties}</w:trPr>${cells.join('')}</w:tr>`;
/** A table with one grid column per cell of its first row, and a paragraph after it. */
const TABLE = (properties: string, ...rows: string[]): string =>
  `<w:tbl><w:tblPr>${properties}</w:tblPr><w:tblGrid>` +
  `${'<w:gridCol w:w="2400"/>'.repeat(rows[0]!.split('<w:tc>').length - 1)}` +
  `</w:tblGrid>${rows.join('')}</w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>`;
const AUTO_WIDTH = '<w:tblW w:w="0" w:type="auto"/>';
const TWO_BY_TWO = TABLE(
  AUTO_WIDTH,
  ROW(CELL('a1'), CELL('a2')),
  ROW_WITH('<w:cantSplit/>', CELL('b1'), CELL('b2'))
);

const tableId = (part: OoxmlPart): string => collectByKind(part.root, 'table')[0]!.id;
const rowIds = (part: OoxmlPart): string[] =>
  collectByKind(part.root, 'tableRow').map((row) => row.id);
const gridColumnIds = (part: OoxmlPart): string[] =>
  collectByKind(part.root, 'generic')
    .filter((node) => node.localName === 'gridCol')
    .map((node) => node.id);

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op} refused: ${result.reason}`);
  return result.part;
}

const untracked = (op: TreeDocOp): TreeDocOp => ({ ...op, revision: undefined });
const fingerprint = (part: OoxmlPart): string => canonicalOoxmlFingerprint(part.root);
const tableXml = (part: OoxmlPart): string =>
  serializeOoxmlPart(part).match(/<w:tbl>.*<\/w:tbl>/)![0];
/** Every serialized `w:trPr`, `w:tblPr` or `w:tblGrid` of the table, in document order. */
const containers = (part: OoxmlPart, name: 'trPr' | 'tblPr' | 'tblGrid'): string[] =>
  tableXml(part).match(new RegExp(`<w:${name}>.*?</w:${name}>(?=<w:(tc|tblGrid|tr)>)`, 'g')) ?? [];
/** Every serialized `<w:tcPr>…</w:tcPr>` in the part, in document order. */
const cellProperties = (part: OoxmlPart): string[] =>
  tableXml(part).match(/<w:tcPr>.*?<\/w:tcPr>(?=<w:p)/g) ?? [];

const rowHeight = (part: OoxmlPart, row = 1, heightTwips = 800): TreeDocOp => ({
  op: 'setTableRowHeight',
  tableId: tableId(part),
  rowId: rowIds(part)[row]!,
  heightTwips,
  revision: BEA,
});
const columnWidths = (part: OoxmlPart, left = 3000, right = 1800): TreeDocOp => ({
  op: 'setTableColumnWidths',
  tableId: tableId(part),
  leftGridColumnId: gridColumnIds(part)[0]!,
  rightGridColumnId: gridColumnIds(part)[1]!,
  leftWidthTwips: left,
  rightWidthTwips: right,
  revision: BEA,
});
const rightEdge = (part: OoxmlPart): TreeDocOp => ({
  op: 'setTableRightEdgeWidth',
  tableId: tableId(part),
  gridColumnId: gridColumnIds(part)[1]!,
  columnWidthTwips: 3000,
  tableWidthTwips: 5400,
  revision: BEA,
});
const tableStyle = (part: OoxmlPart): TreeDocOp => ({
  op: 'setTableProperties',
  tableId: tableId(part),
  styleId: 'TableGrid',
  revision: BEA,
});
const headerRows = (part: OoxmlPart, headerRowCount: number): TreeDocOp => ({
  op: 'setTableProperties',
  tableId: tableId(part),
  headerRowCount,
  revision: BEA,
});

/** A serialized change record of this author's, holding `inner` as the old properties. */
const record = (name: string, inner: string, id = '0'): string =>
  `<w:${name}Change ${STAMP} w:id="${id}">` +
  (inner ? `<w:${name}>${inner}</w:${name}>` : `<w:${name}/>`) +
  `</w:${name}Change>`;
const PLAIN_CELL = '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>';
const SNAPSHOT_CELL = `<w:tcPr><w:tcW w:type="dxa" w:w="2400"/>${record('tcPr', '<w:tcW w:type="dxa" w:w="2400"/>')}</w:tcPr>`;

describe('tracked row properties write w:trPrChange', () => {
  test('a row height keeps the new w:trPr and records the old one as its last child', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, rowHeight(part));
    expect(containers(next, 'trPr')).toEqual([
      '<w:trPr><w:cantSplit/><w:trHeight w:hRule="exact" w:val="800"/>' +
        `${record('trPr', '<w:cantSplit/>')}</w:trPr>`,
    ]);
    // Rejecting a row record resets a cell of that row without one, so every cell of the row
    // records its unchanged properties. The other row is not part of the decision.
    expect(cellProperties(next)).toEqual([PLAIN_CELL, PLAIN_CELL, SNAPSHOT_CELL, SNAPSHOT_CELL]);
  });

  test("a row's own revision marker stays on the row, out of the record", () => {
    const marker = '<w:del w:id="5" w:author="Ada"/>';
    const part = load(TABLE(AUTO_WIDTH, ROW(CELL('a')), ROW_WITH(marker, CELL('b'))));
    const next = apply(part, rowHeight(part));
    expect(containers(next, 'trPr')).toEqual([
      '<w:trPr><w:trHeight w:hRule="exact" w:val="800"/><w:del w:author="Ada" w:id="5"/>' +
        `${record('trPr', '', '6')}</w:trPr>`,
    ]);
  });

  test("Reject keeps the formatting of the row's cells, which a record-less cell would lose", () => {
    const shaded = `${TC_W}<w:shd w:val="clear" w:fill="FFFF00"/><w:vAlign w:val="bottom"/>`;
    const part = load(TABLE(AUTO_WIDTH, ROW(CELL('a', shaded), CELL('b', shaded))));
    const rejected = apply(apply(part, rowHeight(part, 0)), { op: 'rejectAllRevisions' });
    expect(tableXml(rejected).replace('<w:tr><w:trPr/>', '<w:tr>')).toBe(tableXml(part));
  });

  test('a second change keeps the ORIGINAL as the record', () => {
    const part = load(TWO_BY_TWO);
    const once = apply(part, rowHeight(part));
    const twice = apply(once, rowHeight(once, 1, 1200));
    expect(containers(twice, 'trPr')).toEqual([
      '<w:trPr><w:cantSplit/><w:trHeight w:hRule="exact" w:val="1200"/>' +
        `${record('trPr', '<w:cantSplit/>', '1')}</w:trPr>`,
    ]);
  });

  test('a change back to the original drops the records of the whole row', () => {
    const exact = '<w:trHeight w:val="400" w:hRule="exact"/>';
    const part = load(TABLE(AUTO_WIDTH, ROW(CELL('a')), ROW_WITH(exact, CELL('b'))));
    const changed = apply(part, rowHeight(part));
    expect(serializeOoxmlPart(changed)).toContain('<w:trPrChange');
    expect(serializeOoxmlPart(apply(changed, rowHeight(changed, 1, 400)))).toBe(
      serializeOoxmlPart(part)
    );
  });

  test('one op is one revision across the row and its cells', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, rowHeight(part));
    const ids = collectByKind(next.root, 'generic')
      .filter((node) => node.localName === 'trPrChange' || node.localName === 'tcPrChange')
      .map((node) => node.attributes.find((attribute) => attribute.localName === 'id')?.value);
    expect(ids).toEqual(['0', '0', '0']);
  });

  test("a row the author proposed inserting needs no record; another author's does", () => {
    const own = `<w:ins w:id="5" ${STAMP}/>`;
    const mine = load(TABLE(AUTO_WIDTH, ROW(CELL('a')), ROW_WITH(own, CELL('b'))));
    expect(serializeOoxmlPart(apply(mine, rowHeight(mine)))).not.toMatch(/PrChange/);
    const theirs = load(
      TABLE(AUTO_WIDTH, ROW(CELL('a')), ROW_WITH('<w:ins w:id="5" w:author="Ada"/>', CELL('b')))
    );
    expect(serializeOoxmlPart(apply(theirs, rowHeight(theirs)))).toContain('<w:trPrChange');
  });

  test('header rows record each row whose w:tblHeader changed, and only those', () => {
    const header = '<w:tblHeader w:val="1"/>';
    const part = load(
      TABLE(AUTO_WIDTH, ROW_WITH(header, CELL('a')), ROW(CELL('b')), ROW(CELL('c')))
    );
    const next = apply(part, headerRows(part, 2));
    expect(containers(next, 'trPr')).toEqual([
      '<w:trPr><w:tblHeader w:val="1"/></w:trPr>',
      `<w:trPr><w:tblHeader w:val="1"/>${record('trPr', '')}</w:trPr>`,
    ]);
    expect(cellProperties(next)).toEqual([PLAIN_CELL, SNAPSHOT_CELL, PLAIN_CELL]);
  });
});

describe('tracked table properties write w:tblPrChange and w:tblGridChange', () => {
  test('a table style records the old w:tblPr, with no change element inside it', () => {
    const part = load(TWO_BY_TWO);
    const once = apply(part, tableStyle(part));
    const twice = apply(once, { ...tableStyle(once), styleId: 'LightList' });
    expect(containers(twice, 'tblPr')).toEqual([
      '<w:tblPr><w:tblStyle w:val="LightList"/><w:tblW w:type="auto" w:w="0"/>' +
        `${record('tblPr', '<w:tblW w:type="auto" w:w="0"/>', '1')}</w:tblPr>`,
    ]);
    // A table record resets no cell, so no cell records anything.
    expect(cellProperties(twice)).toEqual([PLAIN_CELL, PLAIN_CELL, PLAIN_CELL, PLAIN_CELL]);
  });

  test('a column resize records the old grid, cell widths and layout under one revision', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, columnWidths(part));
    expect(containers(next, 'tblPr')).toEqual([
      '<w:tblPr><w:tblW w:type="auto" w:w="0"/><w:tblLayout w:type="fixed"/>' +
        `${record('tblPr', '<w:tblW w:type="auto" w:w="0"/>')}</w:tblPr>`,
    ]);
    // `CT_TblGridChange` is `CT_Markup`: an id and nothing else.
    expect(containers(next, 'tblGrid')).toEqual([
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="1800"/><w:tblGridChange w:id="0">' +
        '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        '</w:tblGridChange></w:tblGrid>',
    ]);
    const widened = (width: number): string =>
      `<w:tcPr><w:tcW w:type="dxa" w:w="${width}"/>${record('tcPr', '<w:tcW w:type="dxa" w:w="2400"/>')}</w:tcPr>`;
    expect(cellProperties(next)).toEqual([
      widened(3000),
      widened(1800),
      widened(3000),
      widened(1800),
    ]);
  });

  test('a layout that is already fixed needs no table record', () => {
    const part = load(TABLE('<w:tblLayout w:type="fixed"/>', ROW(CELL('a'), CELL('b'))));
    const xml = serializeOoxmlPart(apply(part, columnWidths(part)));
    expect(xml).toContain('<w:tblGridChange');
    expect(xml).not.toContain('<w:tblPrChange');
  });

  test('a no-op grid record, as Google writes in every table, is taken over and then shed', () => {
    const grid =
      '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/><w:tblGridChange w:id="0">' +
      '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
    const part = load(
      TABLE('<w:tblLayout w:type="fixed"/>', ROW(CELL('a'), CELL('b'))).replace(
        /<w:tblGrid>.*<\/w:tblGrid>/,
        grid
      )
    );
    expect(revisionItemsOf(part)).toHaveLength(0);
    const resized = apply(part, columnWidths(part));
    expect(containers(resized, 'tblGrid')[0]).toBe(
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="1800"/><w:tblGridChange w:id="1">' +
        '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        '</w:tblGridChange></w:tblGrid>'
    );
    const back = apply(resized, columnWidths(resized, 2400, 2400));
    expect(serializeOoxmlPart(back)).not.toMatch(/Change/);
  });

  test('the right edge records the old grid, widths and table width', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, rightEdge(part));
    expect(containers(next, 'tblPr')[0]).toBe(
      '<w:tblPr><w:tblW w:type="dxa" w:w="5400"/><w:tblLayout w:type="fixed"/>' +
        `${record('tblPr', '<w:tblW w:type="auto" w:w="0"/>')}</w:tblPr>`
    );
    expect(containers(next, 'tblGrid')[0]).toContain('<w:tblGridChange w:id="0">');
  });

  test('a table the author proposed whole records nothing', () => {
    const own = `<w:ins w:id="5" ${STAMP}/>`;
    const part = load(TABLE(AUTO_WIDTH, ROW_WITH(own, CELL('a'), CELL('b'))));
    for (const op of [columnWidths(part), tableStyle(part), rowHeight(part, 0)]) {
      expect(serializeOoxmlPart(apply(part, op))).not.toMatch(/PrChange|GridChange/);
    }
  });
});

const GESTURES: readonly (readonly [string, (part: OoxmlPart) => TreeDocOp])[] = [
  ['row height', (part) => rowHeight(part)],
  ['header rows', (part) => headerRows(part, 1)],
  ['table style', tableStyle],
  ['column resize', (part) => columnWidths(part)],
  ['right edge', rightEdge],
];

describe.each(GESTURES)('%s', (_name, gesture) => {
  test('is one review card, which Accept and Reject answer whole', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, gesture(part));
    const items = revisionItemsOf(next);
    expect(items.map((item) => item.revisionKind)).toEqual(['format']);
    const resolve = (action: 'accept' | 'reject'): OoxmlPart => {
      const plan = planRevisionBatch(next, action, [reviewItemKey(items[0]!)]);
      expect(plan.result.skipped).toEqual([]);
      return plan.ops.reduce(apply, next);
    };
    expect(fingerprint(resolve('accept'))).toBe(fingerprint(apply(part, untracked(gesture(part)))));
    // Reject restores every row and cell, the ones the gesture left alone included. A row that
    // had no `w:trPr` keeps the empty one the record needed, as a rejected row deletion does.
    expect(tableXml(resolve('reject')).replace('<w:tr><w:trPr/>', '<w:tr>')).toBe(tableXml(part));
  });

  test('a save keeps the proposal, and reopening reads the same tree', () => {
    const part = load(TWO_BY_TWO);
    const next = apply(part, gesture(part));
    const reopened = readOoxmlPart(serializeOoxmlPart(next), {
      name: NAME,
      contentType: CONTENT_TYPE,
    });
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(revisionItemsOf(reopened.part).map((item) => item.revisionKind)).toEqual(['format']);
    expect(fingerprint(reopened.part)).toBe(fingerprint(next));
    expect(diffSemanticDigests(semanticDigest([next]), semanticDigest([reopened.part]))).toEqual(
      []
    );
  });

  test('an attribution without an author is refused', () => {
    const part = load(TWO_BY_TWO);
    expect(validateTreeOp(part, { ...gesture(part), revision: { author: ' ' } } as TreeDocOp)).toBe(
      'invalid-property-value'
    );
  });
});

test("rejecting one author's row leaves another author's row and every other cell alone", () => {
  const part = load(TWO_BY_TWO);
  const cy = (from: OoxmlPart): TreeDocOp => ({
    ...rowHeight(from, 1, 1200),
    revision: { author: 'Cy' },
  });
  const both = apply(apply(part, rowHeight(part, 0)), cy(part));
  const cards = revisionItemsOf(both);
  expect(cards.map((card) => card.author)).toEqual(['Bea', 'Cy']);
  const rejected = planRevisionBatch(both, 'reject', [reviewItemKey(cards[0]!)]).ops.reduce(
    apply,
    both
  );
  const withoutIds = (xml: string): string =>
    xml.replace(/ w:id="\d+"/g, '').replace('<w:tr><w:trPr/>', '<w:tr>');
  expect(withoutIds(tableXml(rejected))).toBe(withoutIds(tableXml(apply(part, cy(part)))));
});

// A ROW HOLDS ONE AUTHOR'S PENDING DECISION AT A TIME. Bea's change to a row with Bob's pending
// change took his records over on the cells it changed. Rejecting hers and then his reset those
// cells to `tcW auto`, so their widths were lost. Bea's change is refused instead.

const BOB = { author: 'Bob', date: '2026-10-01T00:00:00Z' };
const REFUSED = { ok: false, reason: 'unsupported-revision', detail: OTHER_AUTHORS_ROW_CHANGE };
const fill = (part: OoxmlPart, row: number, revision = BEA, hex = 'FF0000'): TreeDocOp => ({
  op: 'setTableCellFill',
  tableId: tableId(part),
  cellIds: [collectByKind(part.root, 'tableCell')[row * 2]!.id],
  color: { kind: 'hex', value: hex },
  revision,
});
/** Bea's changes that reach the first row; a column or right-edge resize reaches every row. */
const firstRowChanges = (part: OoxmlPart): TreeDocOp[] => [
  fill(part, 0),
  rowHeight(part, 0),
  columnWidths(part),
  rightEdge(part),
];
const authors = (part: OoxmlPart): unknown[] => revisionItemsOf(part).map((card) => card.author);
const resolveAll = (part: OoxmlPart, action: 'accept' | 'reject'): OoxmlPart => {
  const plan = planRevisionBatch(part, action);
  expect(plan.result.skipped).toEqual([]);
  return plan.ops.reduce(apply, part);
};
/** Every cell's `w:tcW`, as `type:width`. */
const widths = (part: OoxmlPart): string[] =>
  [...tableXml(part).matchAll(/<w:tcW w:type="(\w+)" w:w="(\d+)"\/>/g)].map(
    ([, type, width]) => `${type}:${width}`
  );

describe("a row with another author's pending change", () => {
  /** Bob's pending fill in the first row, which records on each of its cells. */
  const bobsFill = (): OoxmlPart => {
    const part = load(TWO_BY_TWO);
    return apply(part, fill(part, 0, BOB, '00FF00'));
  };

  test('refuses a change to it, and takes one to another row', () => {
    const held = bobsFill();
    for (const op of firstRowChanges(held)) expect(applyTreeOp(held, op)).toEqual(REFUSED);
    for (const op of [fill(held, 1), rowHeight(held, 1)]) {
      expect(authors(apply(held, op))).toEqual(['Bob', 'Bea']);
    }
    // Editing mode writes the row as it did.
    expect(applyTreeOp(held, untracked(columnWidths(held))).ok).toBe(true);
  });

  test("takes its author's second change, whose Reject restores the original", () => {
    const held = bobsFill();
    expect(applyTreeOp(held, fill(held, 0))).toEqual(REFUSED);
    const again = apply(apply(held, fill(held, 0, BOB, '0000FF')), {
      ...columnWidths(held),
      revision: BOB,
    });
    expect(authors(again)).toEqual(['Bob']);
    expect(tableXml(resolveAll(again, 'reject'))).toBe(tableXml(load(TWO_BY_TWO)));
  });

  test.each(['accept', 'reject'] as const)(
    'takes a change once its pending change is answered (%s), and no width is lost',
    (action) => {
      const held = bobsFill();
      expect(applyTreeOp(held, columnWidths(held))).toEqual(REFUSED);
      const answered = resolveAll(held, action);
      const resized = apply(answered, columnWidths(answered));
      expect(authors(resized)).toEqual(['Bea']);
      const resizedWidths = ['dxa:3000', 'dxa:1800', 'dxa:3000', 'dxa:1800'];
      expect(widths(resolveAll(resized, 'accept'))).toEqual(resizedWidths);
      expect(widths(resolveAll(resized, 'reject'))).toEqual(Array(4).fill('dxa:2400'));
    }
  );
});

// Records another producer wrote: Word, Google Docs, or an earlier copy of this editor.
const HELD_ROWS: readonly (readonly [string, string])[] = [
  [
    'a row record',
    ROW_WITH(
      '<w:trHeight w:val="400"/><w:trPrChange w:id="7" w:author="Ada"><w:trPr/></w:trPrChange>',
      CELL('a1'),
      CELL('a2')
    ),
  ],
  [
    "one cell's record",
    ROW(
      CELL('a1'),
      CELL(
        'a2',
        `${TC_W}<w:shd w:val="clear" w:fill="FFFF00"/>` +
          `<w:tcPrChange w:id="7" w:author="Ada"><w:tcPr>${TC_W}</w:tcPr></w:tcPrChange>`
      )
    ),
  ],
  [
    'a record with no author',
    ROW(CELL('a1'), CELL('a2', `${TC_W}<w:tcPrChange w:id="7"><w:tcPr/></w:tcPrChange>`)),
  ],
  [
    'a table exception record',
    '<w:tr><w:tblPrEx><w:jc w:val="center"/>' +
      '<w:tblPrExChange w:id="7" w:author="Ada"><w:tblPrEx/></w:tblPrExChange></w:tblPrEx>' +
      `${CELL('a1')}${CELL('a2')}</w:tr>`,
  ],
];

describe.each(HELD_ROWS)('a row holding %s', (_name, held) => {
  test('refuses a change to it, and takes one to another row', () => {
    const part = load(TABLE(AUTO_WIDTH, held, ROW(CELL('b1'), CELL('b2'))));
    for (const op of firstRowChanges(part)) expect(applyTreeOp(part, op)).toEqual(REFUSED);
    expect(applyTreeOp(part, rowHeight(part, 1)).ok).toBe(true);
  });
});
