// Two authors' row heights in one table, in Word's own form (fixtures from Word for Mac): each
// author's change is one card, a decision of either one takes only its own height, and a
// decision of the author whose `w:tblPrChange` the table holds moves the other's part into a
// row-level `w:tblPrExChange`, as Word's reject does.
import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  planRevisionBatch,
  readOoxmlPart,
  reviewItemKey,
  revisionItemsOf,
  serializeOoxmlPart,
  type OoxmlPart,
} from '../index.ts';
import {
  wordTwoAuthorRowHeights,
  wordTwoAuthorRowHeightsAnaRejected,
} from './fixtures/word-row-height-authors.ts';

const load = (body: string): OoxmlPart => {
  const read = readOoxmlPart(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw Error(read.reason);
  return read.part;
};

/** The part after deciding the cards of `author` (every card when undefined). */
function decide(part: OoxmlPart, action: 'accept' | 'reject', author?: string): OoxmlPart {
  const keys = revisionItemsOf(part)
    .filter((item) => author === undefined || item.author === author)
    .map(reviewItemKey);
  const plan = planRevisionBatch(part, action, author === undefined ? undefined : keys);
  expect(plan.result.skipped).toEqual([]);
  let next = part;
  for (const op of plan.ops) {
    const result = applyTreeOp(next, op);
    if (!result.ok) throw Error(result.reason);
    next = result.part;
  }
  return next;
}

/** The table's XML with no ids and dates, and self-closing elements in one form. */
const normalized = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>') + 8)
    .replace(/ w:(id|date)="[^"]*"/g, '')
    .replace(/\s*\/>/g, '/>');

const heights = (part: OoxmlPart): string[] =>
  [...serializeOoxmlPart(part).matchAll(/<w:trHeight[^>]*w:val="(\d+)"/g)].map((m) => m[1]!);

const records = (part: OoxmlPart): string[] =>
  [...serializeOoxmlPart(part).matchAll(/<w:(\w+Change)\b([^>]*)>/g)].map((m) => {
    const author = /w:author="([^"]*)"/.exec(m[2]!)?.[1];
    return author === undefined ? m[1]! : `${m[1]}:${author}`;
  });

test("Word's two row heights by two authors are two cards", () => {
  const items = revisionItemsOf(load(wordTwoAuthorRowHeights));
  expect(items.map((item) => item.author)).toEqual(['Ana', 'nx7194']);
});

test("rejecting the table record's author alone gives Word's row-level form", () => {
  const part = decide(load(wordTwoAuthorRowHeights), 'reject', 'Ana');
  expect(heights(part)).toEqual(['1000']);
  expect(normalized(serializeOoxmlPart(part))).toBe(
    normalized(serializeOoxmlPart(load(wordTwoAuthorRowHeightsAnaRejected)))
  );
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['nx7194']);
});

test("accepting the table record's author alone keeps the other's card", () => {
  const part = decide(load(wordTwoAuthorRowHeights), 'accept', 'Ana');
  expect(heights(part)).toEqual(['800', '1000']);
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['nx7194']);
  expect(heights(decide(part, 'reject'))).toEqual(['800']);
});

test('rejecting the second author alone takes only their height', () => {
  const part = decide(load(wordTwoAuthorRowHeights), 'reject', 'nx7194');
  expect(heights(part)).toEqual(['800']);
  expect(records(part)).toEqual([
    'tblPrChange:Ana',
    'tblGridChange',
    'tcPrChange:Ana',
    'tcPrChange:Ana',
  ]);
  expect(heights(decide(part, 'reject'))).toEqual([]);
});

test('accepting the second author alone keeps both heights and the first card', () => {
  const part = decide(load(wordTwoAuthorRowHeights), 'accept', 'nx7194');
  expect(heights(part)).toEqual(['800', '1000']);
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['Ana']);
  expect(heights(decide(part, 'reject'))).toEqual(['1000']);
});

for (const [action, expected] of [
  ['accept', ['800', '1000']],
  ['reject', []],
] as const)
  test(`${action} all decides both heights`, () => {
    const part = decide(load(wordTwoAuthorRowHeights), action);
    expect(heights(part)).toEqual([...expected]);
    expect(records(part)).toEqual([]);
  });

test("Word's row-level form is one card, and its reject takes the height", () => {
  const part = load(wordTwoAuthorRowHeightsAnaRejected);
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['nx7194']);
  const rejected = decide(part, 'reject');
  expect(heights(rejected)).toEqual([]);
  expect(records(rejected)).toEqual([]);
  expect(heights(decide(part, 'accept'))).toEqual(['1000']);
});

// THE ENGINE WRITES WORD'S FORM. The base is Word's table with its records and heights taken
// out; the engine's ops then give the file Word saved, ids and dates apart.

/** Word's table before either height. */
const base = wordTwoAuthorRowHeights
  .replace(/<w:(tblPrChange|tblGridChange|tcPrChange)\b[^>]*>.*?<\/w:\1>/g, '')
  .replace(/<w:trPr><w:trHeight [^>]*\/><\/w:trPr>/g, '');
const ANA = { author: 'Ana', date: '2026-10-30T13:38:00Z' };
const NX = { author: 'nx7194', date: '2026-10-30T13:38:00Z' };

function tracked(part: OoxmlPart, row: number, heightTwips: number, revision = ANA): OoxmlPart {
  const table = part.root.children.find((n) => n.kind === 'body')!;
  const tbl = table.kind !== 'textValue' ? table.children.find((n) => n.kind === 'table') : null;
  const rows =
    tbl && tbl.kind !== 'textValue' ? tbl.children.filter((n) => n.kind === 'tableRow') : [];
  const result = applyTreeOp(part, {
    op: 'setTableRowHeight',
    tableId: tbl!.id,
    rowId: rows[row]!.id,
    heightTwips,
    revision,
  });
  if (!result.ok) throw Error(`${result.reason} ${result.detail ?? ''}`);
  return result.part;
}

/** `normalized`, with Word's "at least" rule for the engine's exact one. */
const asWord = (part: OoxmlPart): string =>
  normalized(serializeOoxmlPart(part)).replace(/ w:hRule="exact"/g, '');
const word = (body: string): string => normalized(serializeOoxmlPart(load(body)));

test("the first author's height is Word's form: table records, a plain height, cell records", () => {
  const part = tracked(load(base), 0, 800);
  expect(asWord(part)).toBe(
    word(
      wordTwoAuthorRowHeights
        .replace(/<w:tcPrChange w:id="[45]"[^>]*>.*?<\/w:tcPrChange>/g, '')
        .replace('<w:trPr><w:trHeight w:val="1000"/></w:trPr>', '')
    )
  );
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['Ana']);
  expect(heights(decide(part, 'reject'))).toEqual([]);
});

test("a second author's height leans on the first author's table record, as Word writes it", () => {
  const part = tracked(tracked(load(base), 0, 800), 2, 1000, NX);
  expect(asWord(part)).toBe(word(wordTwoAuthorRowHeights));
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['Ana', 'nx7194']);
});

test("rejecting the engine's first author alone gives Word's row-level form", () => {
  const part = decide(tracked(tracked(load(base), 0, 800), 2, 1000, NX), 'reject', 'Ana');
  expect(asWord(part)).toBe(word(wordTwoAuthorRowHeightsAnaRejected));
  expect(heights(decide(part, 'reject'))).toEqual([]);
});

for (const [action, author, kept, after] of [
  ['reject', 'nx7194', ['800'], []],
  ['accept', 'nx7194', ['800', '1000'], ['1000']],
  ['accept', 'Ana', ['800', '1000'], ['800']],
] as const)
  test(`${action} of ${author} alone in the engine's form, then reject of the rest`, () => {
    const part = decide(tracked(tracked(load(base), 0, 800), 2, 1000, NX), action, author);
    expect(heights(part)).toEqual([...kept]);
    expect(heights(decide(part, 'reject'))).toEqual([...after]);
  });

test("the engine's form saves and reopens with the same cards", () => {
  const part = tracked(tracked(load(base), 0, 800), 2, 1000, NX);
  const xml = serializeOoxmlPart(part);
  const reopened = load(xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>') + 8));
  expect(revisionItemsOf(reopened).map((item) => item.author)).toEqual(['Ana', 'nx7194']);
  expect(heights(decide(reopened, 'reject', 'nx7194'))).toEqual(['800']);
});

test("an author's own second height in the table is decided with the table record or alone", () => {
  const later = { ...ANA, date: '2026-10-30T13:40:00Z' };
  const part = tracked(tracked(load(base), 0, 800), 2, 1000, later);
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['Ana', 'Ana']);
  expect(heights(decide(part, 'reject'))).toEqual([]);
  const [first] = revisionItemsOf(part);
  const plan = planRevisionBatch(part, 'reject', [reviewItemKey(first!)]);
  const rejected = plan.ops.reduce((next, op) => {
    const result = applyTreeOp(next, op);
    if (!result.ok) throw Error(result.reason);
    return result.part;
  }, part);
  expect(heights(rejected)).toEqual(['1000']);
  expect(heights(decide(rejected, 'reject'))).toEqual([]);
});

test('a height in a table whose record holds another change leans on its own row record', () => {
  const styled = base.replace(
    '<w:tblLayout w:type="fixed"/>',
    '<w:tblLayout w:type="fixed"/><w:tblPrChange w:id="9" w:author="Ana" w:date="2026-10-30T13:38:00Z"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>'
  );
  const part = tracked(load(styled), 2, 1000, NX);
  expect(serializeOoxmlPart(part)).toContain('<w:tblPrExChange');
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['Ana', 'nx7194']);
  expect(heights(decide(part, 'reject', 'nx7194'))).toEqual([]);
  expect(heights(decide(part, 'reject', 'Ana'))).toEqual(['1000']);
});

test("a row's own table exceptions stay in its anchor and its record", () => {
  const margins = '<w:tblCellMar><w:left w:type="dxa" w:w="400"/></w:tblCellMar>';
  const rows = base
    .replace(
      '<w:tblLayout w:type="fixed"/>',
      '<w:tblLayout w:type="fixed"/><w:tblPrChange w:id="9" w:author="Ana" w:date="2026-10-30T13:38:00Z"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange>'
    )
    .split('<w:tr>');
  rows[3] = `<w:tblPrEx>${margins}</w:tblPrEx>${rows[3]}`;
  const part = tracked(load(rows.join('<w:tr>')), 2, 1000, NX);
  const row = serializeOoxmlPart(part).split('<w:tr>')[3]!;
  const exceptions = row.slice(0, row.indexOf('<w:tblPrExChange'));
  const record = row.slice(row.indexOf('<w:tblPrExChange'), row.indexOf('</w:tblPrExChange>'));
  expect(exceptions).toContain(margins);
  expect(record).toContain(margins);
  // Reject takes the height and keeps the row's own margins
  const rejected = serializeOoxmlPart(decide(part, 'reject', 'nx7194'));
  expect(heights(decide(part, 'reject', 'nx7194'))).toEqual([]);
  expect(rejected.split('<w:tr>')[3]).toContain(`<w:tblPrEx>${margins}</w:tblPrEx>`);
});

// A second author's table change takes the table's one record over, as Word does
// (`~$c2-as.docx`); the first author's height then leans on its row's own record, so no
// author is lost (D17).
test("a table change over another author's height moves that height onto its row", () => {
  const before = tracked(load(base), 0, 800);
  const table = before.root.children.find((n) => n.kind === 'body');
  const tbl =
    table && table.kind !== 'textValue' ? table.children.find((n) => n.kind === 'table') : null;
  const result = applyTreeOp(before, {
    op: 'setTableProperties',
    tableId: tbl!.id,
    preferredWidth: { type: 'pct', value: 5000 },
    revision: NX,
  });
  if (!result.ok) throw Error(result.reason);
  const part = result.part;
  expect(records(part)).toEqual([
    'tblPrChange:nx7194',
    'tblGridChange',
    'tblPrExChange:Ana',
    'tcPrChange:Ana',
    'tcPrChange:Ana',
  ]);
  expect(revisionItemsOf(part).map((item) => item.author)).toEqual(['nx7194', 'Ana']);
  expect(heights(decide(part, 'reject', 'nx7194'))).toEqual(['800']);
  expect(heights(decide(part, 'reject', 'Ana'))).toEqual([]);
  expect(serializeOoxmlPart(decide(part, 'reject', 'Ana'))).toContain('w:type="pct"');
  const all = decide(part, 'reject');
  expect(heights(all)).toEqual([]);
  expect(records(all)).toEqual([]);
});
