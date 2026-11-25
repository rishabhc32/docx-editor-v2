// A row's property records and another person's new column. Word records a row height on the
// row and on every cell of it (`w:tcPrChange` holding the cell's properties as they show), and
// its reject gives every recorded cell its record back. So each record holds what the cell
// shows under the pending column (its lent width, which Word writes no record for), and a
// reject of the row height leaves that column whole. Found by the acceptance run (quick
// pass 103, cy-1:51). The other order, a column in a row with a pending row height, was
// refused; now the record lends the same width, and the new cell gets the row's record too, so
// both orders write one file.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart, type OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { planRevisionBatch } from '../../../store/store/revision-batch.ts';
import { revisionItemsOf } from '../../../store/index.ts';
import { reviewItemKey } from '../../../store/store/review-items.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedCards } from '../decide.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { addressOf, apply, BODY, EDIT, opens } from './change-apply-fixture.ts';
import { readOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { liveOf } from './sessions/reopen.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import {
  wordCreatedExistingRowHeight,
  wordCreatedNewRowHeight,
} from '../../../store/__tests__/fixtures/word-created-row-height.ts';
import {
  wordTwoAuthorRowHeights,
  wordTwoAuthorRowHeightsAnaRejected,
} from '../../../store/__tests__/fixtures/word-row-height-authors.ts';
import { docx, packageOf, p, t } from './word-fixtures.ts';

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;
/** One row of two cells, with the row's own properties. */
const oneRow = (trPr = '') =>
  docx(
    '<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>' +
      `<w:tr>${trPr}${cell('A1')}${cell('B1')}</w:tr></w:tbl>` +
      p(t('After.'))
  );
const FILE = oneRow();

const tableXml = (part: OoxmlPart): string =>
  serializeOoxmlPart(part).match(/<w:tbl>.*<\/w:tbl>/)?.[0] ?? '';

/** Each cell's width as it shows, and as a reject of its `w:tcPrChange` gives it back. */
function widths(xml: string): { shown: string[]; recorded: string[] } {
  const cells = [...xml.matchAll(/<w:tc>(.*?)<\/w:tc>/g)].map((m) => m[1] ?? '');
  const width = (tcPr: string) => /<w:tcW [^>]*w:w="(\d+)"/.exec(tcPr)?.[1] ?? '';
  const record = (c: string) => /<w:tcPrChange\b.*?<\/w:tcPrChange>/.exec(c)?.[0];
  return {
    shown: cells.map((c) => width(c.replace(/<w:tcPrChange\b.*?<\/w:tcPrChange>/, ''))),
    recorded: cells.map((c) => width(record(c) ?? c)),
  };
}

const shownWidths = (doc: OrderedDocument) =>
  widths(tableXml(composeOrderedPart(doc, BODY).part)).shown;

type Order = 'column first' | 'height first';

/** Ana suggests a column right of A1; Cy suggests a row height on the row. */
function suggested(order: Order = 'column first'): OrderedDocument {
  const doc = loadOrderedPackage(packageOf(FILE));
  const column = () =>
    apply(doc, 'ana:1', {
      kind: 'insert-column',
      ...EDIT,
      cell: addressOf(opens(doc, 'tc')[0]!),
      where: 'right',
    });
  const height = () =>
    apply(
      doc,
      'cy:1',
      {
        kind: 'table-property',
        ...EDIT,
        on: [addressOf(opens(doc, 'tr')[0]!)],
        op: { op: 'setTableRowHeight', heightTwips: 480 },
      },
      'Cy'
    );
  if (order === 'height first') height();
  column();
  expect(shownWidths(doc)).toEqual(['1170', '1170', '2340']);
  if (order === 'column first') height();
  return doc;
}

/** The engine's decision of one author's revisions in a saved part. */
function engineDecides(part: OoxmlPart, author: string, action: 'accept' | 'reject'): OoxmlPart {
  const keys = revisionItemsOf(part)
    .filter((item) => item.author === author)
    .map(reviewItemKey);
  const plan = planRevisionBatch(part, action, keys);
  expect(plan.result.skipped).toEqual([]);
  for (const op of plan.ops) {
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  return part;
}

/** The table's markup with no revision ids, dates or paragraph ids. */
const plainTable = (part: OoxmlPart) =>
  tableXml(part).replace(/ w:(?:id|date)="[^"]*"| w14:\w+="[^"]*"|<w:trPr\/>/g, '');

describe("a party's row height over another party's new column", () => {
  test('records each cell as it shows, under its own author', () => {
    const doc = suggested();
    const xml = tableXml(composeOrderedPart(doc, BODY).part);
    // Word's reject of the row height gives each cell its shown width: Ana's column stays whole
    expect(widths(xml)).toEqual({
      shown: ['1170', '1170', '2340'],
      recorded: ['1170', '1170', '2340'],
    });
    const authors = (name: string) =>
      [...xml.matchAll(new RegExp(`<w:${name} [^>]*w:author="(\\w+)"`, 'g'))].map((m) => m[1]);
    // Word's form for a height on a row with no properties: no row record, the table's
    expect(authors('trPrChange')).toEqual([]);
    expect(authors('tblPrChange')).toEqual(['Cy']);
    expect(authors('tcPrChange')).toEqual(['Cy', 'Cy', 'Cy']);
    expect(authors('cellIns')).toEqual(['Ana']);
  });

  test('keeps its own card: the row reaches past the new cell', () => {
    const doc = suggested();
    expect(orderedCards(doc).map((card) => [card.author, card.suggestions])).toEqual([
      ['Ana', ['S:ana:1.1']],
      ['Cy', ['S:cy:1.1']],
    ]);
  });

  test("a reject of the row height, in the room or by the engine on the file, keeps Ana's column", () => {
    const doc = suggested();
    const { record } = withUndoRecord(doc, () => setStatus(doc, 'S:cy:1.1', 'rejected'));
    expect(shownWidths(doc)).toEqual(['1170', '1170', '2340']);
    takeBack(doc, record);
    // The engine rejects Cy's revision alone in the saved part
    const part = composeOrderedPart(doc, BODY).part;
    expect(revisionItemsOf(part).filter((item) => item.author === 'Cy')).toHaveLength(1);
    expect(widths(tableXml(engineDecides(part, 'Cy', 'reject'))).shown).toEqual([
      '1170',
      '1170',
      '2340',
    ]);
    // Then Ana's column, rejected: the table as the file had it
    const { record: both } = withUndoRecord(doc, () => {
      setStatus(doc, 'S:cy:1.1', 'rejected');
      setStatus(doc, 'S:ana:1.1', 'rejected');
    });
    expect(shownWidths(doc)).toEqual(['2340', '2340']);
    takeBack(doc, both);
  });

  test('saves and reopens with the same view and decisions', () => {
    checkSaveAndLoad(suggested(), 'row height over a new column', packageOf(FILE));
  });
});

describe("a party's new column over another party's pending row height", () => {
  test('writes the same file as the other order: records as shown, each under its author', () => {
    const height = plainTable(composeOrderedPart(suggested('height first'), BODY).part);
    expect(height).toBe(plainTable(composeOrderedPart(suggested('column first'), BODY).part));
    expect(widths(height)).toEqual({
      shown: ['1170', '1170', '2340'],
      recorded: ['1170', '1170', '2340'],
    });
  });

  test('keeps two cards', () => {
    // Cy's height records the table first, as Word's form does, so its card comes first
    expect(
      orderedCards(suggested('height first')).map((card) => [card.author, card.suggestions])
    ).toEqual([
      ['Cy', ['S:cy:1.1']],
      ['Ana', ['S:ana:1.1']],
    ]);
  });

  test.each([
    ['Cy', 'reject', 'Ana', 'reject', ['2340', '2340']],
    ['Ana', 'accept', 'Cy', 'reject', ['1170', '1170', '2340']],
    ['Cy', 'accept', 'Ana', 'reject', ['2340', '2340']],
  ] as const)(
    'decided in the room and by the engine alike: %s %s, then %s %s',
    (first, how, second, then, end) => {
      for (const order of ['height first', 'column first'] as const) {
        const doc = suggested(order);
        let part = composeOrderedPart(doc, BODY).part;
        const sid = (author: string) => (author === 'Ana' ? 'S:ana:1.1' : 'S:cy:1.1');
        for (const [author, action] of [
          [first, how],
          [second, then],
        ] as const) {
          part = engineDecides(part, author, action);
          setStatus(doc, sid(author), action === 'accept' ? 'accepted' : 'rejected');
          expect(plainTable(composeOrderedPart(doc, BODY).part)).toBe(plainTable(part));
        }
        expect(widths(tableXml(part)).shown).toEqual([...end]);
      }
    }
  );

  test('Reject All in the editor alone gives the table as the file had it', () => {
    for (const order of ['height first', 'column first'] as const) {
      const result = applyTreeOp(composeOrderedPart(suggested(order), BODY).part, {
        op: 'rejectAllRevisions',
      });
      if (!result.ok) throw new Error(result.reason);
      expect(plainTable(result.part)).toBe(
        plainTable(composeOrderedPart(loadOrderedPackage(packageOf(FILE)), BODY).part)
      );
    }
  });

  test("the organiser's reject of the column first gives the table its widths back", () => {
    // The engine holds this decision while Cy's record is on Ana's cell (`incomplete-group`)
    const doc = suggested('height first');
    setStatus(doc, 'S:ana:1.1', 'rejected');
    expect(widths(tableXml(composeOrderedPart(doc, BODY).part))).toEqual({
      shown: ['2340', '2340'],
      recorded: ['2340', '2340'],
    });
  });

  test('saves and reopens with the same view and decisions', () => {
    checkSaveAndLoad(suggested('height first'), 'new column over a row height', packageOf(FILE));
  });
});

// The engine lends a new column's width and grid column from the original a pending record holds
// too (`withRecordLending`, `lentGrid`), so a reject of that record keeps the column lent, whether
// the column is pending or accepted then
describe("a party's new column over another party's pending widths", () => {
  // The height is the row's own record on a row with properties (Word's form)
  const file = oneRow('<w:trPr><w:cantSplit/></w:trPr>');

  /**
   * Cy suggests a row height and A1 wider (the engine lends from a recorded width under a row
   * record only); Ana suggests a column right of A1, which takes half of its width.
   */
  function widened(): OrderedDocument {
    const doc = loadOrderedPackage(packageOf(file));
    const row = [addressOf(opens(doc, 'tr')[0]!)];
    const height = { op: 'setTableRowHeight', heightTwips: 480 } as const;
    apply(doc, 'cy:1', { kind: 'table-property', ...EDIT, on: row, op: height }, 'Cy');
    const on = [addressOf(opens(doc, 'tc')[0]!)];
    const op = { op: 'setTableColumnWidths', leftWidthTwips: 3000, rightWidthTwips: 1680 } as const;
    apply(doc, 'cy:2', { kind: 'table-property', ...EDIT, on, op }, 'Cy');
    const cell = addressOf(opens(doc, 'tc')[0]!);
    apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, cell, where: 'right' });
    return doc;
  }

  test("writes the engine's form: the grid record's original split as the engine splits it", () => {
    const xml = tableXml(composeOrderedPart(widened(), BODY).part);
    const grid = (part: string) => [...part.matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => m[1]);
    const history = /<w:tblGridChange\b.*?<\/w:tblGridChange>/.exec(xml)?.[0] ?? '';
    expect(grid(history)).toEqual(['1170', '1170', '2340']);
    expect(widths(xml)).toEqual({
      shown: ['1500', '1500', '1680'],
      recorded: ['840', '1500', '2340'],
    });
  });

  test.each([
    ['Cy', 'reject', 'Ana', 'accept', ['840', '1500', '2340']],
    ['Ana', 'accept', 'Cy', 'reject', ['840', '1500', '2340']],
    ['Cy', 'reject', 'Ana', 'reject', ['2340', '2340']],
    ['Cy', 'accept', 'Ana', 'reject', ['3000', '1680']],
  ] as const)(
    'decided in the room and by the engine alike: %s %s, then %s %s',
    (first, how, second, then, end) => {
      const doc = widened();
      let part = composeOrderedPart(doc, BODY).part;
      const sids = (author: string) =>
        author === 'Ana' ? ['S:ana:1.1'] : ['S:cy:1.1', 'S:cy:2.1'];
      for (const [author, action] of [
        [first, how],
        [second, then],
      ] as const) {
        part = engineDecides(part, author, action);
        for (const sid of sids(author)) {
          setStatus(doc, sid, action === 'accept' ? 'accepted' : 'rejected');
        }

        expect(plainTable(composeOrderedPart(doc, BODY).part)).toBe(plainTable(part));
      }

      expect(widths(tableXml(part)).shown).toEqual([...end]);
    }
  );

  test('saves and reopens with the same view and decisions', () => {
    checkSaveAndLoad(widened(), 'new column over pending widths', packageOf(file));
  });
});

// Found by the acceptance run (quick pass 101, cy-1:135 and cy-1:138): a party's new column,
// then the same party's table or row property. The record holds what each cell shows (the width
// its own column lent); a take-back of both must still give the table as the file had it
describe("a party's table or row property over its own new column", () => {
  const ops = {
    table: { op: 'setTableProperties', justification: 'center' },
    row: { op: 'setTableRowHeight', heightTwips: 480 },
  } as const;
  const original = () =>
    plainTable(composeOrderedPart(loadOrderedPackage(packageOf(FILE)), BODY).part);

  function own(what: 'table' | 'row'): OrderedDocument {
    const doc = loadOrderedPackage(packageOf(FILE));
    apply(
      doc,
      'cy:1',
      { kind: 'insert-column', ...EDIT, cell: addressOf(opens(doc, 'tc')[0]!), where: 'right' },
      'Cy'
    );
    const on = opens(doc, what === 'table' ? 'tbl' : 'tr')[0]!;
    apply(
      doc,
      'cy:2',
      { kind: 'table-property', ...EDIT, on: [addressOf(on)], op: ops[what] },
      'Cy'
    );
    expect(shownWidths(doc)).toEqual(['1170', '1170', '2340']);
    return doc;
  }

  /** The engine's decision of one review item in a saved part: the column or the property. */
  function engineDecidesOne(part: OoxmlPart, column: boolean, action: 'accept' | 'reject') {
    const keys = revisionItemsOf(part)
      .filter((item) => (item.revisionKind === 'structural') === column)
      .map(reviewItemKey);
    expect(keys).toHaveLength(1);
    const plan = planRevisionBatch(part, action, keys);
    expect(plan.result.skipped).toEqual([]);
    for (const op of plan.ops) {
      const result = applyTreeOp(part, op);
      if (!result.ok) throw new Error(result.reason);
      part = result.part;
    }
    return part;
  }

  for (const what of ['table', 'row'] as const) {
    test(`${what}: a take-back of both gives the table as the file had it`, () => {
      const doc = own(what);
      const part = composeOrderedPart(doc, BODY).part;
      expect(plainTable(engineDecides(part, 'Cy', 'reject'))).toBe(original());
      const result = applyTreeOp(part, { op: 'rejectAllRevisions' });
      if (!result.ok) throw new Error(result.reason);
      expect(plainTable(result.part)).toBe(original());
      setStatus(doc, 'S:cy:1.1', 'rejected');
      setStatus(doc, 'S:cy:2.1', 'rejected');
      expect(plainTable(composeOrderedPart(doc, BODY).part)).toBe(original());
    });

    test(`${what}: each decision order gives one table, in the room and by the engine`, () => {
      for (const [first, how, then] of [
        ['column', 'reject', 'reject'],
        ['property', 'reject', 'reject'],
        ['column', 'accept', 'reject'],
        ['property', 'accept', 'reject'],
        ['column', 'reject', 'accept'],
        ['property', 'reject', 'accept'],
      ] as const) {
        const doc = own(what);
        let part = composeOrderedPart(doc, BODY).part;
        const steps = [
          [first === 'column', how],
          [first !== 'column', then],
        ] as const;
        for (const [column, action] of steps) {
          part = engineDecidesOne(part, column, action);
          setStatus(
            doc,
            column ? 'S:cy:1.1' : 'S:cy:2.1',
            action === 'accept' ? 'accepted' : 'rejected'
          );
          // While the record is pending the engine keeps the grid it was taken in (a joined
          // cell spans two grid columns, open item D): the cells show the same widths
          const room = tableXml(composeOrderedPart(doc, BODY).part);
          expect({ first, how, then, room: widths(room).shown }).toEqual({
            first,
            how,
            then,
            room: widths(tableXml(part)).shown,
          });
        }
        expect({ first, how, then, room: plainTable(composeOrderedPart(doc, BODY).part) }).toEqual({
          first,
          how,
          then,
          room: plainTable(part),
        });
        const columnRejected = (first === 'column' ? how : then) === 'reject';
        expect(widths(tableXml(part)).shown).toEqual(
          columnRejected ? ['2340', '2340'] : ['1170', '1170', '2340']
        );
      }
    });
  }

  test('saves and reopens with the same view and decisions', () => {
    for (const what of ['table', 'row'] as const)
      checkSaveAndLoad(own(what), `${what} over its own column`, packageOf(FILE));
  });
});

// Word's own row height (a file Word saved: `store/__tests__/fixtures/word-created-row-height.ts`):
// a `w:tblPrChange` and a `w:tblGridChange` with nothing changed, a `w:tcPrChange` on each cell of
// the row, and the old height in a `w:trPrChange` only when the row had a `w:trPr`. Each record
// has a `w:id` of its own; the engine reviews them as one change, and its reject of the form with
// no row record takes the height away. The room read each record as a suggestion of its own, and
// rejecting all of them kept the height
describe('a row height Word saved', () => {
  for (const [name, fixture, rejected] of [
    ['on a row with no row properties', wordCreatedNewRowHeight, undefined],
    ['on a row with a height', wordCreatedExistingRowHeight, '400'],
  ] as const) {
    test(`${name}: one suggestion, decided as the engine decides it`, () => {
      const file = docx(fixture.replace(/ \/>/g, '/>') + p(t('After.')));
      const heightOf = (part: OoxmlPart) =>
        /<w:trHeight w:val="(\d+)"/.exec(
          serializeOoxmlPart(part).replace(/<w:trPrChange\b.*?<\/w:trPrChange>/g, '')
        )?.[1];
      for (const action of ['accept', 'reject'] as const) {
        const doc = loadOrderedPackage(packageOf(file));
        expect(orderedCards(doc).map((card) => card.author)).toEqual(['Ada']);
        const [only] = doc.suggestions.keys();
        const part = composeOrderedPart(doc, BODY).part;
        expect(heightOf(part)).toBe('800');
        expect(serializeOoxmlPart(part).includes('<w:trPrChange')).toBe(rejected !== undefined);
        const engine = engineDecides(part, 'Ada', action);
        setStatus(doc, only!, action === 'accept' ? 'accepted' : 'rejected');
        const room = composeOrderedPart(doc, BODY).part;
        expect(plainTable(room)).toBe(plainTable(engine));
        expect(heightOf(room)).toBe(action === 'accept' ? '800' : rejected);
        expect(serializeOoxmlPart(room)).not.toMatch(/<w:(tblPr|tblGrid|trPr|tcPr)Change\b/);
      }
      checkSaveAndLoad(
        loadOrderedPackage(packageOf(file)),
        `Word's row height ${name}`,
        packageOf(file)
      );
    });
  }
});

// Two authors' row heights in one table, as Word for Mac saved them: the table holds Ana's one
// `w:tblPrChange`, and nx7194's height is an untracked `w:trHeight` and nx7194's own
// `w:tcPrChange` on the row's cells. Word's reject of Ana's alone moves nx7194's part into a
// row-level `w:tblPrExChange` (D17). The room reads each author's height as one suggestion and
// decides it as the engine decides the file.
describe("two authors' row heights Word saved", () => {
  const ANA_REJECTED = wordTwoAuthorRowHeightsAnaRejected;
  const heights = (part: OoxmlPart) =>
    [...serializeOoxmlPart(part).matchAll(/<w:trHeight w:val="(\d+)"/g)].map((m) => m[1]);
  const sidOf = (doc: OrderedDocument, author: string) =>
    [...doc.suggestions.values()].find((s) => s.attribution.author === author)!.id;
  const fileOf = (table: string) => docx(table + p(t('After.')));

  test('one suggestion for each author, decided in any order as the engine decides them', () => {
    const file = fileOf(wordTwoAuthorRowHeights);
    for (const [first, how, second, then] of [
      ['Ana', 'reject', 'nx7194', 'reject'],
      ['Ana', 'accept', 'nx7194', 'reject'],
      ['nx7194', 'reject', 'Ana', 'reject'],
      ['nx7194', 'accept', 'Ana', 'accept'],
      ['Ana', 'reject', 'nx7194', 'accept'],
    ] as const) {
      const doc = loadOrderedPackage(packageOf(file));
      expect(orderedCards(doc).map((card) => card.author)).toEqual(['Ana', 'nx7194']);
      let part = composeOrderedPart(doc, BODY).part;
      expect(heights(part)).toEqual(['800', '1000']);
      for (const [author, action] of [
        [first, how],
        [second, then],
      ] as const) {
        part = engineDecides(part, author, action);
        setStatus(doc, sidOf(doc, author), action === 'accept' ? 'accepted' : 'rejected');
        const room = composeOrderedPart(doc, BODY).part;
        expect({ author, action, room: plainTable(room) }).toEqual({
          author,
          action,
          room: plainTable(part),
        });
      }
    }
    checkSaveAndLoad(
      loadOrderedPackage(packageOf(file)),
      "Word's two row heights",
      packageOf(file)
    );
  });

  test("Ana's reject gives Word's row-level form; the rest decides as the engine decides it", () => {
    const doc = loadOrderedPackage(packageOf(fileOf(wordTwoAuthorRowHeights)));
    setStatus(doc, sidOf(doc, 'Ana'), 'rejected');
    const room = composeOrderedPart(doc, BODY).part;
    const word = composeOrderedPart(loadOrderedPackage(packageOf(fileOf(ANA_REJECTED))), BODY).part;
    expect(plainTable(room)).toBe(plainTable(word));
    for (const action of ['accept', 'reject'] as const) {
      const file = fileOf(ANA_REJECTED);
      const reopened = loadOrderedPackage(packageOf(file));
      expect(orderedCards(reopened).map((card) => card.author)).toEqual(['nx7194']);
      const engine = engineDecides(composeOrderedPart(reopened, BODY).part, 'nx7194', action);
      setStatus(reopened, sidOf(reopened, 'nx7194'), action === 'accept' ? 'accepted' : 'rejected');
      expect(plainTable(composeOrderedPart(reopened, BODY).part)).toBe(plainTable(engine));
      expect(heights(engine)).toEqual(action === 'accept' ? ['1000'] : []);
      checkSaveAndLoad(
        loadOrderedPackage(packageOf(file)),
        "Word's row-level height",
        packageOf(file)
      );
    }
  });
});

// Word keeps one `w:tcPrChange` on a cell, so a later change of the row's cells takes the
// height's cell records over (policy 2). The height has no record of its own: Word's file ties
// it to its row's cell records, so a reopened file gives it to the change that took them. The
// live room kept it under the earlier suggestion, which a reject then took away alone
// (random session seed 19, two people).
describe("a later change that takes over a row height's cell records", () => {
  const TWO_ROWS = docx(
    '<w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2340"/><w:gridCol w:w="2340"/></w:tblGrid>' +
      `<w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl>` +
      p(t('After.'))
  );
  const later = {
    fill: { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } },
    drag: { op: 'setTableColumnWidths', leftWidthTwips: 2000, rightWidthTwips: 2680 },
  } as const;

  function suggestedTwice(what: keyof typeof later, author: string): OrderedDocument {
    const doc = loadOrderedPackage(packageOf(TWO_ROWS));
    const row = opens(doc, 'tr')[1]!;
    apply(
      doc,
      'ana:1',
      {
        kind: 'table-property',
        ...EDIT,
        on: [addressOf(row)],
        op: { op: 'setTableRowHeight', heightTwips: 600 },
      },
      'Ana',
      '2026-10-28T00:00:04Z'
    );
    const cells = opens(doc, 'tc');
    apply(
      doc,
      `${author.toLowerCase()}:2`,
      {
        kind: 'table-property',
        ...EDIT,
        on: (what === 'fill' ? cells.slice(2) : cells.slice(0, 1)).map(addressOf),
        op: later[what],
      },
      author,
      '2026-10-28T00:01:00Z'
    );
    return doc;
  }

  /** The body with `ids` rejected; the document is restored. */
  function rejected(doc: OrderedDocument, ids: readonly string[]): string {
    const { record } = withUndoRecord(doc, () => {
      for (const id of ids) setStatus(doc, id, 'rejected');
    });
    const table = plainTable(composeOrderedPart(doc, BODY).part);
    takeBack(doc, record);
    return table;
  }

  // Another person's fill or drag over those cell records is the engine's refusal, as in the
  // editor alone; another person's table change leaves the height theirs (D17, rule 4)
  test("another person's fill or drag over the height's cells is refused", () => {
    for (const what of ['fill', 'drag'] as const)
      expect(() => suggestedTwice(what, 'Cy')).toThrow('unsupported-revision');
  });

  for (const what of ['fill', 'drag'] as const)
    test(`${what}: each live reject equals the reopened file's`, () => {
      const doc = suggestedTwice(what, 'Ana');
      const read = readOoxmlPackage(orderedPackageBytes(doc, packageOf(TWO_ROWS)), {});
      if (!read.ok) throw new Error(read.reason);
      const again = loadOrderedPackage(read.package);
      const live = liveOf(doc, again);
      for (const id of [...doc.suggestions.keys()]) {
        const reopened = [...again.suggestions.keys()].filter((r) => live.get(r) === id);
        expect({ id, table: rejected(doc, [id]) }).toEqual({
          id,
          table: rejected(again, reopened),
        });
      }
      checkSaveAndLoad(doc, what, packageOf(TWO_ROWS));
    });
});
