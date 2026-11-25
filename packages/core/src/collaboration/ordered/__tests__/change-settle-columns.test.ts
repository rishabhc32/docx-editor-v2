// A decided column is written into its table (`change/settle-columns.ts`): rows made after the
// decision are made on the table the decision left, and the live list, its saved file and every
// later decision agree on that table. Found by the random session `sessions-paste` seed 5.
//
// A decision joins the grid once, as it is applied, and writes the joined table into the items,
// so a save or a fold cannot lose it. The live list once joined the grid at every drawing while
// the table held a decided label, which a save does not keep: a decision that took away the last
// row over a grid line was joined live only (`sessions-two` seed 5), and so was a later direct
// row deletion, which the editor alone and a fold never join.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { addressOf, apply, BODY, DIRECT, EDIT, opens } from './change-apply-fixture.ts';
import { folded, mediaStore } from './fold-check.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, packageOf, p, t } from './word-fixtures.ts';

const cell = (body: string, props: string) => `<w:tc><w:tcPr>${props}</w:tcPr>${body}</w:tc>`;
const width = '<w:tcW w:w="3000" w:type="dxa"/>';
// Two columns, the first merged down two rows
const FILE = docx(
  p(t('Before.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(p(t('Alpha')), `${width}<w:vMerge w:val="restart"/>`)}${cell(p(t('Fee')), width)}</w:tr>` +
    `<w:tr>${cell('<w:p/>', `${width}<w:vMerge/>`)}${cell(p(t('30 days')), width)}</w:tr>` +
    '</w:tbl>' +
    p(t('After.'))
);

/** The table's grid and each row's cell widths and spans, as compose draws them. */
function tableOf(doc: OrderedDocument, pending: 'show' | 'reject' = 'show'): string[] {
  const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY, { pending }).part);
  // Without the change records: a row height records the grid unchanged (Word's form)
  const table = (xml.match(/<w:tbl>.*<\/w:tbl>/)?.[0] ?? '').replace(
    /<w:(tcPrChange|tblGridChange)\b.*?<\/w:\1>/g,
    ''
  );
  const grid = [
    ...(table.match(/<w:tblGrid>.*?<\/w:tblGrid>/)?.[0] ?? '').matchAll(/w:w="(\d+)"/g),
  ];
  const rows = (table.match(/<w:tr>.*?<\/w:tr>/g) ?? []).map((row) =>
    [...row.matchAll(/<w:tcW[^>]*w:w="(\d+)"[^>]*\/>(<w:gridSpan w:val="(\d+)"\/>)?/g)]
      .map((m) => (m[3] ? `${m[1]}x${m[3]}` : m[1]))
      .join(',')
  );
  return [`grid ${grid.map((m) => m[1]).join(',')}`, ...rows];
}

/** Ana fills the merged cell and deletes the second column; Olive accepts the deletion. */
function decided(): OrderedDocument {
  const doc = loadOrderedPackage(packageOf(FILE));
  const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } } as const;
  apply(doc, 'ana:1', {
    kind: 'table-property',
    ...EDIT,
    on: [addressOf(opens(doc, 'tc')[2]!)],
    op: fill,
  });
  apply(doc, 'ana:2', { kind: 'delete-column', ...EDIT, cell: addressOf(opens(doc, 'tc')[1]!) });
  const seen = orderedCardSeen(doc, 'S:ana:2.1') ?? [];
  apply(
    doc,
    'olive:1',
    { kind: 'decide', cards: [{ suggestions: seen }], status: 'accepted' },
    'Olive'
  );
  return doc;
}

describe('a decided column', () => {
  test('is written into the table: compose draws it from the items alone', () => {
    const doc = decided();
    const joined = ['grid 6000', '6000', '6000'];
    expect(tableOf(doc)).toEqual(joined);
    expect(tableOf(doc, 'reject')).toEqual(joined);
    // The removed cells show nowhere and are taken out no more
    const cells = [...(doc.stories.get(BODY)?.tree.items() ?? [])].filter(
      (item) => item.kind === 'open' && item.token === 'tc'
    );
    expect(cells.map((open) => open.gone)).toEqual([false, true, false, true]);
  });

  test("a row made after the decision keeps the table's grid, live and after a reopen", () => {
    const doc = decided();
    const rows = opens(doc, 'tr');
    apply(doc, 'ana:3', {
      kind: 'table-property',
      ...EDIT,
      on: [addressOf(rows[1]!)],
      op: { op: 'setTableRowHeight', heightTwips: 600 },
    });
    apply(
      doc,
      'olive:2',
      { kind: 'insert-row', ...EDIT, row: addressOf(rows[1]!), where: 'below' },
      'Olive'
    );
    expect(tableOf(doc)).toEqual(['grid 6000', '6000', '6000', '6000']);
    // Every later decision keeps it: the row rejected, or accepted
    for (const status of ['rejected', 'accepted'] as const) {
      const { record } = withUndoRecord(doc, () => {
        for (const id of ['S:ana:3.1', 'S:olive:2.1']) setStatus(doc, id, status);
      });
      expect({ status, table: tableOf(doc) }).toEqual({
        status,
        table: ['grid 6000', '6000', '6000', ...(status === 'accepted' ? ['6000'] : [])],
      });
      takeBack(doc, record);
    }
    checkSaveAndLoad(doc, 'row after a decided column', packageOf(FILE));
  });
});

// Two columns: a first row over both, and a second row of two cells
const SPANNED = docx(
  p(t('Before.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(p(t('Fee and rate')), '<w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/>')}</w:tr>` +
    `<w:tr>${cell(p(t('100')), width)}${cell(p(t('5')), width)}</w:tr>` +
    '</w:tbl>' +
    p(t('After.'))
);

/** Olive decides the card of `suggestion`. */
function decide(doc: OrderedDocument, id: string, suggestion: string, status: Status): void {
  const seen = orderedCardSeen(doc, suggestion) ?? [];
  apply(doc, id, { kind: 'decide', cards: [{ suggestions: seen }], status }, 'Olive');
}

type Status = 'accepted' | 'rejected';

/** Ana's fill of the first row's cell, accepted. */
function acceptedFill(doc: OrderedDocument): void {
  const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } } as const;
  const on = [addressOf(opens(doc, 'tc')[0]!)];
  apply(doc, 'ana:1', { kind: 'table-property', ...EDIT, on, op: fill });
  decide(doc, 'olive:1', 'S:ana:1.1', 'accepted');
}

describe('a decided column that takes every cell of a row', () => {
  // Ana deletes the first column from the first row's cell, which covers both: Olive accepts
  test('takes the row too, as the engine does, live and after a reopen', () => {
    const doc = loadOrderedPackage(packageOf(SPANNED));
    apply(doc, 'ana:1', { kind: 'delete-column', ...EDIT, cell: addressOf(opens(doc, 'tc')[0]!) });
    decide(doc, 'olive:1', 'S:ana:1.1', 'accepted');
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    expect({ emptyRow: xml.includes('<w:tr/>'), table: tableOf(doc) }).toEqual({
      emptyRow: false,
      table: ['grid 6000', '6000'],
    });
    checkSaveAndLoad(doc, 'a decided column that took a row', packageOf(SPANNED));
  });
});

describe('a decided row', () => {
  // Olive accepts Ana's fill of the first row; Ana adds a row of two cells, and Olive deletes the
  // other one directly: the new row is the last over the grid line
  test('rejected, the last over a grid line, joins the grid live and after a reopen', () => {
    const doc = loadOrderedPackage(packageOf(SPANNED));
    acceptedFill(doc);
    const row = addressOf(opens(doc, 'tr')[1]!);
    apply(doc, 'ana:2', { kind: 'insert-row', ...EDIT, row, where: 'below' });
    apply(doc, 'olive:2', { kind: 'delete-row', ...DIRECT, row }, 'Olive');

    // A drawing as if every suggestion were rejected joins it as the reject does
    expect(tableOf(doc, 'reject')).toEqual(['grid 6000', '6000']);
    checkSaveAndLoad(doc, 'a rejected row, the last over a grid line', packageOf(SPANNED));
    decide(doc, 'olive:3', 'S:ana:2.1', 'rejected');
    expect(tableOf(doc)).toEqual(['grid 6000', '6000']);
  });

  // Bob's pending height on the first row records the grid unchanged on the table (Word's form):
  // it holds the joined grid too, live and after a reopen
  test('rejected beside a pending row height: the grid joins live and after a reopen', () => {
    const first = openOrderedWordFile(SPANNED, { file: 1 });
    const doc = first.document;
    acceptedFill(doc);
    const rows = opens(doc, 'tr');
    apply(doc, 'ana:2', { kind: 'insert-row', ...EDIT, row: addressOf(rows[1]!), where: 'below' });
    const height = { op: 'setTableRowHeight', heightTwips: 600 } as const;
    apply(
      doc,
      'bob:1',
      { kind: 'table-property', ...EDIT, on: [addressOf(rows[0]!)], op: height },
      'Bob'
    );
    apply(doc, 'olive:2', { kind: 'delete-row', ...DIRECT, row: addressOf(rows[1]!) }, 'Olive');
    const again = openOrderedWordFile(orderedPackageBytes(doc, first.base), { file: 2 }).document;
    const tables = [doc, again].map((d) => {
      decide(d, 'olive:3', opens(d, 'tr').at(-1)!.by!, 'rejected');
      return tableOf(d);
    });
    expect(tables).toEqual([
      ['grid 6000', '6000'],
      ['grid 6000', '6000'],
    ]);
  });
});

describe('a direct row deletion after a decision on the table', () => {
  // The editor alone joins no grid line on a direct edit: the live list after the decision, and
  // its fold, delete the row the same way
  for (const [name, decided] of [
    ['an accepted fill', acceptedFill],
    [
      'a rejected column insertion',
      (doc: OrderedDocument) => {
        const cell = addressOf(opens(doc, 'tc')[2]!);
        apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, cell, where: 'right' });
        decide(doc, 'olive:1', 'S:ana:1.1', 'rejected');
      },
    ],
  ] as const) {
    test(`${name}: no grid line goes, live as after a fold`, async () => {
      const first = openOrderedWordFile(SPANNED, { file: 1 });
      decided(first.document);
      const again = await folded(first.document, first.base, 2, mediaStore(first.base));
      for (const doc of [first.document, again.document]) {
        const row = addressOf(opens(doc, 'tr')[1]!);
        apply(doc, 'olive:2', { kind: 'delete-row', ...DIRECT, row }, 'Olive');
      }

      const kept = ['grid 3000,3000', '6000x2'];
      expect({ live: tableOf(first.document), folded: tableOf(again.document) }).toEqual({
        live: kept,
        folded: kept,
      });
    });
  }
});
