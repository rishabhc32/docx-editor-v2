// A decided column is written into its table (`change/settle-columns.ts`): rows made after the
// decision are made on the table the decision left, and the live list, its saved file and every
// later decision agree on that table. Found by the random session `sessions-paste` seed 5.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { addressOf, apply, BODY, EDIT, opens } from './change-apply-fixture.ts';
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
