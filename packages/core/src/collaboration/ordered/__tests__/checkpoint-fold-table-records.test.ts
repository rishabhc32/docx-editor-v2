// A fold of a table's records changes nothing a decision gives (design D14): records apply in the
// order they were made, and the organiser's direct columns are the table's own, live as after a
// fold and as in the editor alone.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedCards } from '../decide.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DIRECT,
  EDIT,
  opens,
} from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

const ANA_1 = '2026-10-28T09:01:00Z';

// Records apply in the order they were made, after a fold too: a record of an older suggestion
// made after another's (a height on the organiser's own new row, a direct height, a second
// height of one person) still shows on top of it
describe('row heights made in turn', () => {
  const height = (doc: OrderedDocument, row: number, heightTwips: number) =>
    ({
      kind: 'table-property',
      on: [addressOf(opens(doc, 'tr')[row]!)],
      op: { op: 'setTableRowHeight', heightTwips },
    }) as const;

  test("a party's height on the organiser's new row, then the organiser's", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const row = addressOf(opens(doc, 'tr')[0]!);
    apply(doc, 'olive:1', { kind: 'insert-row', ...EDIT, row, where: 'below' }, 'Olive');
    apply(doc, 'ana:1', { ...height(doc, 1, 600), ...EDIT }, 'Ana', ANA_1);
    apply(doc, 'olive:2', { ...height(doc, 1, 400), ...EDIT }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'a height on a new row of its maker');
  });

  test("a party's height, the organiser's direct one, then the party's on another row", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ana:1', { ...height(doc, 1, 700), ...EDIT }, 'Ana', ANA_1);
    apply(doc, 'olive:1', { ...height(doc, 1, 400), ...DIRECT }, 'Olive');
    apply(doc, 'ana:2', { ...height(doc, 0, 500), ...EDIT }, 'Ana', ANA_1);
    await checkFold(doc, base, mediaStore(base), 'a direct height between two');
  });

  test("a party's height, the organiser's direct column, then the party's height again", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ana:1', { ...height(doc, 1, 600), ...EDIT }, 'Ana', ANA_1);
    const cell = addressOf(opens(doc, 'tc')[2]!);
    apply(doc, 'olive:1', { kind: 'insert-column', ...DIRECT, cell, where: 'right' }, 'Olive');
    apply(doc, 'ana:2', { ...height(doc, 1, 500), ...EDIT }, 'Ana', ANA_1);
    await checkFold(doc, base, mediaStore(base), 'a height again over a direct column');
  });

  // A height on its row's own anchor is Word's form only on a row Word reads as the height: a
  // row with a record of its own, or a cell with none, keeps no anchor that a fold would read
  // as a record of the row
  const style = (doc: OrderedDocument) =>
    ({
      kind: 'table-property',
      ...EDIT,
      on: [addressOf(opens(doc, 'tbl')[0]!)],
      op: { op: 'setTableProperties', styleId: 'TableGrid', headerRowCount: 1 },
    }) as const;

  test("a party's height, then the party's table style with a header row", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ana:1', { ...height(doc, 0, 500), ...EDIT }, 'Ana', ANA_1);
    apply(doc, 'ana:2', style(doc), 'Ana', ANA_1);
    await checkFold(doc, base, mediaStore(base), 'a height under a header row');
  });

  test("a party's height and table style, then the organiser's direct column", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ana:1', { ...height(doc, 1, 700), ...EDIT }, 'Ana', ANA_1);
    apply(doc, 'ana:2', style(doc), 'Ana', ANA_1);
    const cell = addressOf(opens(doc, 'tc')[2]!);
    apply(doc, 'olive:1', { kind: 'insert-column', ...DIRECT, cell, where: 'right' }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'a height and a direct column');
  });

  // A column its author took back holds cells until compose removes them, which hold no record of
  // a height: the height's anchor reads the cells the table keeps, live as after a fold
  test("a party's column taken back, then another party's height and widths", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const right = { cell: addressOf(opens(doc, 'tc')[2]!), where: 'right' } as const;
    apply(doc, 'ben:1', { kind: 'insert-column', ...EDIT, ...right }, 'Ben');
    const added = addressOf(opens(doc, 'tc')[1]!);
    apply(doc, 'ben:2', { kind: 'delete-column', ...EDIT, cell: added }, 'Ben');
    apply(doc, 'ana:1', { ...height(doc, 1, 500), ...EDIT }, 'Ana', ANA_1);
    const on = [addressOf(opens(doc, 'tc')[0]!)];
    const op = { op: 'setTableColumnWidths', leftWidthTwips: 3600, rightWidthTwips: 2400 } as const;
    apply(doc, 'ana:2', { kind: 'table-property', ...EDIT, on, op }, 'Ana', ANA_1);
    await checkFold(doc, base, mediaStore(base), 'a height beside a column taken back');
  });

  // An accepted height keeps the records it took from its author's table style after a fold, on
  // a row whose cells keep every record (another party's column lent them width)
  test("a party's table style and height, another's column, then the height accepted", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ben:1', style(doc), 'Ben');
    apply(doc, 'ben:2', { ...height(doc, 0, 500), ...EDIT }, 'Ben', ANA_1);
    const cell = addressOf(opens(doc, 'tc')[0]!);
    apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, cell, where: 'right' }, 'Ana', ANA_1);
    const suggestions = orderedCardSeen(doc, 'S:ben:2.1')!;
    const status = 'accepted';
    apply(doc, 'olive:1', { kind: 'decide', cards: [{ suggestions }], status }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'an accepted height that took a style');
  });

  // A pending record taken from another keeps the place it was made in too, before the
  // organiser's column and direct deletion that came after it
  test("a party's height, a column, a direct deletion, then the party's cell alignment", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    apply(doc, 'ben:1', { ...height(doc, 0, 500), ...EDIT }, 'Ben');
    const right = { cell: addressOf(opens(doc, 'tc')[0]!), where: 'right' } as const;
    apply(doc, 'olive:1', { kind: 'insert-column', ...EDIT, ...right }, 'Olive');
    const deleted = addressOf(opens(doc, 'tc')[2]!);
    apply(doc, 'olive:2', { kind: 'delete-column', ...DIRECT, cell: deleted }, 'Olive');
    const on = [addressOf(opens(doc, 'tc')[0]!)];
    const op = { op: 'setTableCellVerticalAlignment', alignment: 'center' } as const;
    apply(doc, 'ben:2', { kind: 'table-property', ...EDIT, on, op }, 'Ben', ANA_1);
    await checkFold(doc, base, mediaStore(base), 'a taken record before a direct deletion');
  });
});

// A table record the load read from the file's change record, and a column made after a fold
// that changes its original: the record holds the changed original after another fold, as live,
// and a reject of the record keeps the column lent once it is accepted
test.each([false, true])(
  "a party's row height and widths, a fold, then the organiser's column (accepted: %p)",
  async (accepted) => {
    const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const blobs = mediaStore(first.base);
    const row = (doc: OrderedDocument, k: number) => addressOf(opens(doc, 'tr')[k]!);
    const height = (doc: OrderedDocument, k: number, heightTwips: number) =>
      ({
        kind: 'table-property',
        on: [row(doc, k)],
        op: { op: 'setTableRowHeight', heightTwips },
      }) as const;
    apply(
      first.document,
      'ben:1',
      { kind: 'delete-row', ...EDIT, row: row(first.document, 0) },
      'Ben'
    );
    apply(first.document, 'ana:1', { ...height(first.document, 0, 700), ...EDIT }, 'Ana');
    const on = [addressOf(opens(first.document, 'tc')[0]!)];
    const op = { op: 'setTableColumnWidths', leftWidthTwips: 3600, rightWidthTwips: 2400 } as const;
    apply(first.document, 'ana:2', { kind: 'table-property', ...EDIT, on, op }, 'Ana', ANA_1);
    const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
    apply(doc, 'olive:1', { ...height(doc, 1, 400), ...DIRECT }, 'Olive');
    const right = { cell: addressOf(opens(doc, 'tc')[2]!), where: 'right' } as const;
    apply(doc, 'olive:2', { kind: 'insert-column', ...EDIT, ...right }, 'Olive');
    if (accepted) {
      const suggestions = orderedCardSeen(doc, 'S:olive:2.1')!;
      apply(
        doc,
        'olive:3',
        { kind: 'decide', cards: [{ suggestions }], status: 'accepted' },
        'Olive'
      );
    }

    await checkFold(doc, base, blobs, 'a column over a file record of widths');
  }
);

// A table of two columns and two rows, and a paragraph after it
const TABLE_FILE = docx(
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    ['Fee', 'Late fee']
      .map((text) => {
        const cell = (body: string) =>
          `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(body))}</w:tc>`;
        return `<w:tr>${cell(text)}${cell('Net 30 days')}</w:tr>`;
      })
      .join('') +
    '</w:tbl>' +
    p(t('Signed.'))
);

// A party's new row joined to their split after the table: the marks of the row's cells take the
// row's `w:id` and the split's mark its own, live as after a fold
test("a party's split after a table, then the party's row joined to it", async () => {
  const { document: doc, base } = openOrderedWordFile(TABLE_FILE, { file: 1 });
  apply(doc, 'ben:1', { kind: 'split', ...EDIT, place: at(doc, 4, 3) }, 'Ben');
  const row = addressOf(opens(doc, 'tr')[0]!);
  const into = 'S:ben:1.1';
  apply(doc, 'ben:2', { kind: 'insert-row', ...EDIT, into, row, where: 'below' }, 'Ben');
  await checkFold(doc, base, mediaStore(base), 'a row joined to a split after its table');
});

// A column a take-back or a reject removed, under a later column: the later column's records of
// the grid and of the widths were taken while it showed, and a fold writes them as the table
// shows without it
describe('a column removed under a later one', () => {
  const cell = (doc: OrderedDocument, k: number) => addressOf(opens(doc, 'tc')[k]!);
  const column = (doc: OrderedDocument, k: number) =>
    ({ kind: 'insert-column', ...EDIT, cell: cell(doc, k), where: 'right' }) as const;

  test.each(['taken back', 'rejected'])(
    "a party's column %s, after the party's column beside another",
    async (how) => {
      const { document: doc, base } = openOrderedWordFile(TABLE_FILE, { file: 1 });
      apply(doc, 'ben:1', column(doc, 0), 'Ben');
      apply(doc, 'ben:2', column(doc, 2), 'Ben');
      if (how === 'taken back') {
        apply(doc, 'ben:3', { kind: 'delete-column', ...EDIT, cell: cell(doc, 1) }, 'Ben');
      } else {
        const suggestions = orderedCardSeen(doc, 'S:ben:1.1')!;
        const status = 'rejected';
        apply(doc, 'olive:1', { kind: 'decide', cards: [{ suggestions }], status }, 'Olive');
      }

      await checkFold(doc, base, mediaStore(base), `a column ${how} under a later grid`);
    }
  );

  // The later column took its width from the same cell, and the removed column's width goes to
  // the later column's new cell before it
  test("a party's column taken back, after the party's column at the same cell", async () => {
    const { document: doc, base } = openOrderedWordFile(TABLE_FILE, { file: 1 });
    apply(doc, 'ben:1', column(doc, 0), 'Ben');
    apply(doc, 'ben:2', column(doc, 0), 'Ben');
    apply(doc, 'ben:3', { kind: 'delete-column', ...EDIT, cell: cell(doc, 2) }, 'Ben');
    await checkFold(doc, base, mediaStore(base), 'a column taken back beside a later cell');
  });

  /** The grid's widths and each row's cell widths, as compose draws them. */
  const widths = (doc: OrderedDocument): string[] =>
    (
      serializeOoxmlPart(composeOrderedPart(doc, BODY).part).match(
        /<w:(tblGrid|tr)>.*?<\/w:\1>/g
      ) ?? []
    ).map((row) =>
      [...row.matchAll(/<w:(?:gridCol|tcW)[^>]*w:w="(\d+)"/g)].map((m) => m[1]).join(',')
    );
  const decide = (
    doc: OrderedDocument,
    id: string,
    suggestion: string,
    status: 'accepted' | 'rejected'
  ) =>
    apply(
      doc,
      id,
      { kind: 'decide', cards: [{ suggestions: orderedCardSeen(doc, suggestion)! }], status },
      'Olive'
    );

  // A take-back is its author's undo: the editor alone withdraws its own pending column by
  // rejecting its cells, so the cell the column took width from gets back what is left of the
  // column's cell after a later column took part of it, as a reject gives it back
  test.each(['taken back', 'rejected'])(
    "a party's column %s, after the party's column made from its cell",
    async (how) => {
      const { document: doc } = openOrderedWordFile(TABLE_FILE, { file: 1 });
      apply(doc, 'ben:1', column(doc, 0), 'Ben');
      apply(doc, 'ben:2', column(doc, 1), 'Ben');
      if (how === 'taken back') {
        apply(doc, 'ben:3', { kind: 'delete-column', ...EDIT, cell: cell(doc, 1) }, 'Ben');
      } else {
        decide(doc, 'olive:1', 'S:ben:1.1', 'rejected');
      }

      expect(widths(doc)).toEqual(['2250,750,3000', '2250,750,3000', '2250,750,3000']);
      expect(orderedCards(doc).map((card) => card.suggestions)).toEqual([['S:ben:2.1']]);
      decide(doc, 'olive:2', 'S:ben:2.1', 'rejected');
      expect(widths(doc)).toEqual(['3000,3000', '3000,3000', '3000,3000']);
    }
  );
});

describe("the organiser's direct columns at a decision", () => {
  // The organiser's direct deletion of a column, after a party's row height and the organiser's
  // own pending column beside it: a reject of the height gives the height's grid back and fits
  // it to the cells, as the editor alone does; the cells the direct deletion removed are gone
  // after a fold, so live draws none of them either
  test("a row height, a column, then the organiser's direct deletion of its neighbour", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const on = [addressOf(opens(doc, 'tr')[0]!)];
    const op = { op: 'setTableRowHeight', heightTwips: 500 } as const;
    apply(doc, 'ben:1', { kind: 'table-property', ...EDIT, on, op }, 'Ben');
    const right = { cell: addressOf(opens(doc, 'tc')[0]!), where: 'right' } as const;
    apply(doc, 'olive:1', { kind: 'insert-column', ...EDIT, ...right }, 'Olive');
    const cell = addressOf(opens(doc, 'tc')[0]!);
    apply(doc, 'olive:2', { kind: 'delete-column', ...DIRECT, cell }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'a direct deletion beside a pending column');
  });

  // A party's deletion of the organiser's direct column: an accept gives the column's width to
  // its neighbour, as the editor alone does; after a fold the direct cells are the file's
  test("the organiser's direct column, then a party's deletion of it", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const right = { cell: addressOf(opens(doc, 'tc')[0]!), where: 'right' } as const;
    apply(doc, 'olive:1', { kind: 'insert-column', ...DIRECT, ...right }, 'Olive');
    const cell = addressOf(opens(doc, 'tc')[1]!);
    apply(doc, 'ana:1', { kind: 'delete-column', ...EDIT, cell }, 'Ana');
    await checkFold(doc, base, mediaStore(base), "a party's deletion of a direct column");
  });

  // The same in a party's pending new table: its cells and the direct column's are the table's
  test("a party's new table, the organiser's direct column, then the party's deletion of it", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const table = { place: at(doc, 0, 0), rows: 2, columns: 2, columnWidth: 2400 };
    apply(doc, 'ben:1', { kind: 'insert-table', ...EDIT, ...table }, 'Ben');
    const right = { cell: addressOf(opens(doc, 'tc')[0]!), where: 'right' } as const;
    apply(doc, 'olive:1', { kind: 'insert-column', ...DIRECT, ...right }, 'Olive');
    const cell = addressOf(opens(doc, 'tc')[1]!);
    apply(doc, 'ben:2', { kind: 'delete-column', ...EDIT, cell }, 'Ben');
    await checkFold(doc, base, mediaStore(base), 'a deletion of a direct column in a new table');
  });

  // The organiser's direct widths over a pending cell fill in the same row: a reject of the fill
  // resets the row's cells with no record of their own (the direct column's new cell), after a
  // fold as live: a fold keeps which suggestions were direct edits
  test("a pending cell fill, then the organiser's direct column and widths", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const cell = (k: number) => addressOf(opens(doc, 'tc')[k]!);
    const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } } as const;
    apply(doc, 'olive:1', { kind: 'table-property', ...EDIT, on: [cell(0)], op: fill }, 'Olive');
    const right = { cell: cell(1), where: 'right' } as const;
    apply(doc, 'olive:2', { kind: 'insert-column', ...DIRECT, ...right }, 'Olive');
    const op = { op: 'setTableColumnWidths', leftWidthTwips: 3600, rightWidthTwips: 2400 } as const;
    apply(doc, 'olive:3', { kind: 'table-property', ...DIRECT, on: [cell(1)], op }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'direct widths over a pending fill');
  });
});
