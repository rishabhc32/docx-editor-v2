// The records a table press leaves, live, in a saved file reopened, and after one and two folds.
//
// In a row its author proposed (a new row, or a row of a new table) the engine writes no change
// record: a reject of the insertion takes the row away, properties and all (#495). So the press
// is part of that insertion, with no card of its own, and the file holds its properties with no
// record and no row anchor. Found by the acceptance run: a row height and a cell alignment in a
// party's own new table read back after a fold as a row anchor and a grid record of nobody.
//
// A press records only the containers the engine wrote: another person's records elsewhere in the
// table stay theirs. A grid record (`w:tblGridChange`) writes its `w:id` alone, so the
// checkpoint names its suggestion, and a fold keeps who made it.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedCards } from '../decide.ts';
import type { OrderedTableOp } from '../edits/tables.ts';
import type { OrderedDocument, OrderedItem } from '../model/types.ts';
import { addressOf, apply, APPLY_FILE, at, BODY, EDIT, opens } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, p, t } from './word-fixtures.ts';

/** A suggested table press on a row or a cell. */
const press = (doc: OrderedDocument, id: string, on: OrderedItem, op: OrderedTableOp, by: string) =>
  apply(doc, id, { kind: 'table-property', ...EDIT, on: [addressOf(on)], op }, by);

const HEIGHT = { op: 'setTableRowHeight', heightTwips: 480 } as const;
const ALIGN = { op: 'setTableCellVerticalAlignment', alignment: 'center' } as const;

/** Each table property record: its container, properties, suggestion and whether it shows bare. */
const records = (doc: OrderedDocument) =>
  ['tbl', 'tr', 'tc'].flatMap((token) =>
    opens(doc, token).flatMap((open) =>
      open.props.map((r) => ({
        token,
        target: r.target,
        suggestion: r.suggestion,
        untracked: r.kind === 'element' && r.untracked === true,
      }))
    )
  );

const tableXml = (doc: OrderedDocument, index: number): string =>
  serializeOoxmlPart(composeOrderedPart(doc, BODY).part).match(/<w:tbl>.*?<\/w:tbl>/g)?.[index] ??
  '';

/** Check the room, its saved file reopened, and its folded copies. */
async function checked(doc: OrderedDocument, base: OoxmlPackage, label: string) {
  checkSaveAndLoad(doc, label, base);
  await checkFold(doc, base, mediaStore(base), label);
}

describe("a table press in its author's own new row or table", () => {
  for (const [label, insert] of [
    [
      'a new row',
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-row',
          ...EDIT,
          row: addressOf(opens(doc, 'tr')[0]!),
          where: 'below',
        });
        return { row: opens(doc, 'tr')[1]!, cell: opens(doc, 'tc')[2]!, table: 0 };
      },
    ],
    [
      'a new table',
      (doc: OrderedDocument) => {
        apply(doc, 'ana:1', {
          kind: 'insert-table',
          ...EDIT,
          place: at(doc, 9, 0),
          rows: 2,
          columns: 2,
          columnWidth: 2400,
        });
        return { row: opens(doc, 'tr').at(-2)!, cell: opens(doc, 'tc').at(-2)!, table: 1 };
      },
    ],
  ] as const) {
    test(`a row height and a cell alignment in ${label}`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const { row, cell, table } = insert(doc);
      const insertion = row.by ?? '';
      press(doc, 'ana:2', row, HEIGHT, 'Ana');
      press(doc, 'ana:3', cell, ALIGN, 'Ana');

      // Both are the insertion's, shown with no record: its one card
      expect(records(doc)).toEqual([
        { token: 'tr', target: 'trPr', suggestion: insertion, untracked: true },
        { token: 'tc', target: 'tcPr', suggestion: insertion, untracked: true },
      ]);
      expect(orderedCards(doc).map((card) => card.suggestions)).toEqual([[insertion]]);

      // Word's form: the properties in the new row, with no record and no anchor
      const xml = tableXml(doc, table);
      expect(xml).toContain('<w:trHeight w:hRule="exact" w:val="480"/>');
      expect(xml).toContain('<w:vAlign w:val="center"/>');
      expect(xml).not.toMatch(/PrChange|PrExChange|GridChange/);
      await checked(doc, base, label);
    });
  }
});

describe('a table press leaves the records it did not write', () => {
  for (const [label, first] of [
    ['a fill', { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } }],
    ['a row height', HEIGHT],
  ] as const) {
    test(`${label} on the first row, then an alignment on the second`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const on = first.op === 'setTableRowHeight' ? opens(doc, 'tr')[0]! : opens(doc, 'tc')[0]!;
      press(doc, 'ana:1', on, first, 'Ana');
      const ana = records(doc);
      press(doc, 'bob:1', opens(doc, 'tc')[3]!, ALIGN, 'Bob');

      // Ana's records stay hers; Bob's alignment records the second row's cells
      const [bob] = orderedCards(doc)
        .flatMap((card) => card.suggestions)
        .filter((id) => !ana.some((record) => record.suggestion === id));
      expect(records(doc)).toEqual([
        ...ana,
        { token: 'tc', target: 'tcPr', suggestion: bob, untracked: false },
        { token: 'tc', target: 'tcPr', suggestion: bob, untracked: false },
      ]);
      expect(orderedCards(doc)).toHaveLength(2);
      await checked(doc, base, label);
    });
  }
});

const cell = (text: string, width: number) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** A table whose grid (3000, 3000) differs from its cells' widths (4000, 2000). */
const UNEVEN = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell('Fee', 4000)}${cell('Rate', 2000)}</w:tr>` +
    '</w:tbl>' +
    p(t('Signed here.'))
);

describe('a fold of a grid record', () => {
  test('a drag that changes no cell keeps its maker', async () => {
    const { document: doc, base } = openOrderedWordFile(UNEVEN, { file: 1 });
    press(
      doc,
      'ana:1',
      opens(doc, 'tc')[0]!,
      { op: 'setTableColumnWidths', leftWidthTwips: 4000, rightWidthTwips: 2000 },
      'Ana'
    );
    expect(tableXml(doc, 0)).toMatch(/<w:tblGridChange w:id="\d+">/);

    // A Word file names nobody for it: only the checkpoint does
    await checkFold(doc, base, mediaStore(base), 'a lone grid record');
  });
});
