// Columns in a table with merged cells, as Word inserts and deletes them. Each expected table is
// the one Word for Mac 16 saved on 2026-10-29 from the same table and the same command (the
// samples are listed in the ordered collaboration design, D17): Word inserts a cell beside the
// first row's cell that covers the column in every row, and deletes the cell that covers it in
// every row, a spanning cell whole.

import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { applyTreeOp, validateTreeOp } from '../store/tree-ops.ts';
import { wmlAttributeValue, wmlChildNamed } from '../store/tree-op-table-shared.ts';

const W = WML_NAMESPACE_URI;
const cell = (text: string, props = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${props.includes('gridSpan w:val="3"') ? 9360 : props.includes('gridSpan') ? 6240 : 3120}" w:type="dxa"/>${props}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (...cells: string[]) => `<w:tr>${cells.join('')}</w:tr>`;
const span = (n: number) => `<w:gridSpan w:val="${n}"/>`;
const table = (...rows: string[]) =>
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>' +
  rows.join('') +
  '</w:tbl><w:p/>';

/** Word's sample: a cell spanning two columns in row 1, a vertical merge in column 3. */
const WORD_BASE = table(
  row(cell('R1C1R1C2', span(2)), cell('R1C3')),
  row(cell('R2C1'), cell('R2C2'), cell('R2C3R3C3', '<w:vMerge w:val="restart"/>')),
  row(
    cell('R3C1'),
    cell('R3C2'),
    `<w:tc><w:tcPr><w:tcW w:w="3120" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>`
  )
);
/** The span in row 2 instead. */
const ROW2_SPAN = table(
  row(cell('R1C1'), cell('R1C2'), cell('R1C3')),
  row(cell('R2C1R2C2', span(2)), cell('R2C3')),
  row(cell('R3C1'), cell('R3C2'), cell('R3C3'))
);
/** A title row: one cell over the whole table. */
const TITLE = table(
  row(cell('R1C1R1C2R1C3', span(3))),
  row(cell('R2C1'), cell('R2C2'), cell('R2C3')),
  row(cell('R3C1'), cell('R3C2'), cell('R3C3'))
);

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function firstTable(part: OoxmlPart): OoxmlElement {
  const stack: OoxmlNode[] = [part.root];
  for (let node = stack.pop(); node; node = stack.pop()) {
    if (node.kind === 'table') return node;
    if (node.kind !== 'textValue') stack.push(...node.children);
  }
  throw new Error('no table');
}

const elements = (node: OoxmlElement, name: string): OoxmlElement[] =>
  node.children.filter((c): c is OoxmlElement => c.kind !== 'textValue' && c.localName === name);

const gridIds = (part: OoxmlPart): string[] =>
  elements(wmlChildNamed(firstTable(part), 'tblGrid')!, 'gridCol').map((c) => c.id);

const textOf = (node: OoxmlNode): string =>
  node.kind === 'textValue' ? node.value : node.children.map(textOf).join('');

/** The table as Word's samples read: the grid, and each row as `text:width/sN/v` cells. */
function shape(part: OoxmlPart): string[] {
  const t = firstTable(part);
  const grid = elements(wmlChildNamed(t, 'tblGrid')!, 'gridCol').map((c) =>
    wmlAttributeValue(c, 'w')
  );
  return [
    `grid ${grid.join(',')}`,
    ...elements(t, 'tr').map((tr) => {
      const props = wmlChildNamed(tr, 'trPr');
      const after = props && wmlChildNamed(props, 'gridAfter');
      const cells = elements(tr, 'tc').map((tc) => {
        const p = wmlChildNamed(tc, 'tcPr');
        const width = p && wmlChildNamed(p, 'tcW');
        const s = p && wmlChildNamed(p, 'gridSpan');
        const v = p && wmlChildNamed(p, 'vMerge');
        return (
          `${textOf(tc) || '_'}:${width ? wmlAttributeValue(width, 'w') : '-'}` +
          (s ? `/s${wmlAttributeValue(s, 'val')}` : '') +
          (v ? `/v${wmlAttributeValue(v, 'val') ?? 'c'}` : '')
        );
      });
      return cells.join(' | ') + (after ? ` +after${wmlAttributeValue(after, 'val')}` : '');
    }),
  ];
}

function apply(part: OoxmlPart, op: Parameters<typeof applyTreeOp>[1]): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

const insertRight = (part: OoxmlPart, column: number) =>
  apply(part, {
    op: 'insertTableColumn',
    tableId: firstTable(part).id,
    gridColumnId: gridIds(part)[column]!,
    where: 'right',
  });

const deleteColumn = (part: OoxmlPart, column: number) =>
  apply(part, {
    op: 'deleteTableColumn',
    tableId: firstTable(part).id,
    gridColumnId: gridIds(part)[column]!,
  });

describe('a column inserted in a table with merged cells, as Word inserts it', () => {
  test('right of the second column: after the spanning cell in row 1, after C1 below', () => {
    expect(shape(insertRight(load(WORD_BASE), 1))).toEqual([
      'grid 3120,3120,3120,3120',
      'R1C1R1C2:6240/s2 | _:3120 | R1C3:3120',
      'R2C1:3120 | _:3120 | R2C2:3120 | R2C3R3C3:3120/vrestart',
      'R3C1:3120 | _:3120 | R3C2:3120 | _:3120/vc',
    ]);
  });

  test('right of the third column: the new cells stand before the vertical merge below', () => {
    expect(shape(insertRight(load(WORD_BASE), 2))).toEqual([
      'grid 3120,3120,3120,3120',
      'R1C1R1C2:6240/s2 | R1C3:3120 | _:3120',
      'R2C1:3120 | R2C2:3120 | _:3120 | R2C3R3C3:3120/vrestart',
      'R3C1:3120 | R3C2:3120 | _:3120 | _:3120/vc',
    ]);
  });

  test('the first row decides the index: a span in row 2 gets its new cell after R2C3', () => {
    expect(shape(insertRight(load(ROW2_SPAN), 1))).toEqual([
      'grid 3120,3120,3120,3120',
      'R1C1:3120 | R1C2:3120 | _:3120 | R1C3:3120',
      'R2C1R2C2:6240/s2 | R2C3:3120 | _:3120',
      'R3C1:3120 | R3C2:3120 | _:3120 | R3C3:3120',
    ]);
  });

  test('a title row over the whole table: the rows below end short of it, padded after', () => {
    // Word gives the title row's new cell the title's width (9360) and pads the rows below;
    // the engine gives every new cell the column's width, so no row ends short
    expect(shape(insertRight(load(TITLE), 1))).toEqual([
      'grid 3120,3120,3120,3120',
      'R1C1R1C2R1C3:9360/s3 | _:3120',
      'R2C1:3120 | _:3120 | R2C2:3120 | R2C3:3120',
      'R3C1:3120 | _:3120 | R3C2:3120 | R3C3:3120',
    ]);
  });
});

describe('a column deleted from a table with merged cells, as Word deletes it', () => {
  test('the column through the spanning cell: the cell goes whole, the row is padded after', () => {
    expect(shape(deleteColumn(load(WORD_BASE), 1))).toEqual([
      'grid 3120,3120',
      'R1C3:3120 +after1',
      'R2C1:3120 | R2C3R3C3:3120/vrestart',
      'R3C1:3120 | _:3120/vc',
    ]);
  });

  test('the column of the vertical merge goes with every cell of the merge', () => {
    expect(shape(deleteColumn(load(WORD_BASE), 2))).toEqual([
      'grid 3120,3120',
      'R1C1R1C2:6240/s2',
      'R2C1:3120 | R2C2:3120',
      'R3C1:3120 | R3C2:3120',
    ]);
  });

  test('a span in row 2: its cell goes, the cells after it move left', () => {
    expect(shape(deleteColumn(load(ROW2_SPAN), 1))).toEqual([
      'grid 3120,3120',
      'R1C1:3120 | R1C3:3120',
      'R2C3:3120 +after1',
      'R3C1:3120 | R3C3:3120',
    ]);
  });

  test('a title row left with no cell goes', () => {
    expect(shape(deleteColumn(load(TITLE), 1))).toEqual([
      'grid 3120,3120',
      'R2C1:3120 | R2C3:3120',
      'R3C1:3120 | R3C3:3120',
    ]);
  });

  test('a column whose deletion empties every row is the whole table: refused', () => {
    const part = load(table(row(cell('All', span(3)))));
    const op = {
      op: 'deleteTableColumn' as const,
      tableId: firstTable(part).id,
      gridColumnId: gridIds(part)[1]!,
    };
    expect(validateTreeOp(part, op)).toBe('block-required');
  });
});

describe('a drag in a table with merged cells', () => {
  test('a vertical merge keeps its cells; a span takes the sum of its columns', () => {
    const part = load(WORD_BASE);
    const [left, right] = gridIds(part).slice(1);
    const dragged = apply(part, {
      op: 'setTableColumnWidths',
      tableId: firstTable(part).id,
      leftGridColumnId: left!,
      rightGridColumnId: right!,
      leftWidthTwips: 4000,
      rightWidthTwips: 2240,
    });
    expect(shape(dragged)).toEqual([
      'grid 3120,4000,2240',
      'R1C1R1C2:7120/s2 | R1C3:2240',
      'R2C1:3120 | R2C2:4000 | R2C3R3C3:2240/vrestart',
      'R3C1:3120 | R3C2:4000 | _:2240/vc',
    ]);
  });
});

describe('a proposed column in a table with merged cells', () => {
  const ana = { author: 'Ana', date: '2026-10-29T00:00:00Z' };
  const decide = (part: OoxmlPart, action: 'accept' | 'reject') =>
    apply(part, { op: action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions' });

  test("an insertion: Word's cells, each lent half its row neighbour's width", () => {
    const part = load(WORD_BASE);
    const proposed = apply(part, {
      op: 'insertTableColumn',
      tableId: firstTable(part).id,
      gridColumnId: gridIds(part)[1]!,
      where: 'right',
      revision: ana,
    });
    expect(shape(proposed)).toEqual([
      'grid 1560,1560,3120,3120',
      'R1C1R1C2:3120/s2 | _:3120 | R1C3:3120',
      'R2C1:1560 | _:1560 | R2C2:3120 | R2C3R3C3:3120/vrestart',
      'R3C1:1560 | _:1560 | R3C2:3120 | _:3120/vc',
    ]);
    // Rejected, the table is Word's again; accepted, it keeps its width
    expect(shape(decide(proposed, 'reject'))).toEqual(shape(part));
    const accepted = decide(proposed, 'accept');
    expect(shape(accepted)).toEqual(shape(proposed));
  });

  test('a deletion marks the cell that covers the column in each row, a spanning cell whole', () => {
    const part = load(WORD_BASE);
    const proposed = apply(part, {
      op: 'deleteTableColumn',
      tableId: firstTable(part).id,
      gridColumnId: gridIds(part)[1]!,
      revision: ana,
    });
    expect(shape(proposed)).toEqual(shape(part));
    const marked = elements(firstTable(proposed), 'tr').map((tr) =>
      elements(tr, 'tc')
        .filter((tc) => wmlChildNamed(wmlChildNamed(tc, 'tcPr')!, 'cellDel'))
        .map(textOf)
    );
    expect(marked).toEqual([['R1C1R1C2'], ['R2C2'], ['R3C2']]);
    expect(shape(decide(proposed, 'reject'))).toEqual(shape(part));
    const accepted = shape(decide(proposed, 'accept'));
    expect(accepted.join('\n')).not.toContain('R1C1R1C2');
    expect(accepted.join('\n')).not.toContain('R2C2');
  });
});
