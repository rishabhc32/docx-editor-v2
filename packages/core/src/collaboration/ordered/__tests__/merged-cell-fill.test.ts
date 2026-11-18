// A cell fill in a vertically merged cell (task 3.6). The benchmark (seed 3, step 69 of
// `bench/__tests__/translate.test.ts` on the 2% contract) had a party fill a cell of a row the
// organiser inserted inside a vertical merge, next to rows parties suggested there. A row
// inserted inside the merge repeats `w:vMerge` (the engine's rule), so that cell is one of the
// merged cell's cells, and the engine refuses a cell op that names only some of them
// (`engine-refused`, `setTableCellFill: invalidArgs`). A change naming one of them now means
// the merged cell: the fill covers all its cells.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { loadOrderedPackage } from '../load/load.ts';
import type { OrderedItem } from '../model/types.ts';
import { addressOf, apply, BODY, DIRECT, EDIT, opens, refusal } from './change-apply-fixture.ts';
import { docx, packageOf, t } from './word-fixtures.ts';

const cell = (text: string, properties = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${properties}</w:tcPr><w:p>${t(text)}</w:p></w:tc>`;

/** Three rows; the first column's cell merges down over the first two rows. */
const FILE = docx(
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell('Merged', '<w:vMerge w:val="restart"/>')}${cell('B1')}</w:tr>` +
    `<w:tr>${cell('', '<w:vMerge/>')}${cell('B2')}</w:tr>` +
    `<w:tr>${cell('A3')}${cell('B3')}</w:tr>` +
    '</w:tbl><w:p/>'
);

const FILL = { op: 'setTableCellFill', color: { kind: 'hex', value: 'E2EFDA' } } as const;

/** The first cell of each row, in order. */
function firstCells(doc: ReturnType<typeof loadOrderedPackage>): OrderedItem[] {
  const cells = opens(doc, 'tc');
  const rows = opens(doc, 'tr');
  const story = doc.stories.get(BODY);
  if (!story) throw new Error('no body');
  return rows.map((row) => {
    for (let item = story.tree.next(row); item; item = story.tree.next(item))
      if (cells.includes(item)) return item;
    throw new Error('a row without cells');
  });
}

/** How many cells of the body hold the fill. */
const filled = (doc: ReturnType<typeof loadOrderedPackage>): number =>
  serializeOoxmlPart(composeOrderedPart(doc, BODY).part).split('w:fill="E2EFDA"').length - 1;

describe('a cell fill in a vertically merged cell', () => {
  test("a party fills a cell of the organiser's row inside the merge, next to a party's row", () => {
    const doc = loadOrderedPackage(packageOf(FILE));
    const [first] = opens(doc, 'tr');
    if (!first) throw new Error('no row');
    // The organiser inserts a row below the first (inside the merge), a party one below hers
    apply(
      doc,
      'olive-1:1',
      { kind: 'insert-row', ...DIRECT, row: addressOf(first), where: 'below' },
      'Olive'
    );
    const olive = opens(doc, 'tr')[1];
    if (!olive) throw new Error('no new row');
    apply(
      doc,
      'cy-1:1',
      { kind: 'insert-row', ...EDIT, row: addressOf(olive), where: 'below' },
      'Cy'
    );
    const cells = firstCells(doc);
    expect(cells).toHaveLength(5);
    // The organiser's row's first cell continues the merge, as the engine inserts it
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    expect(xml.split('<w:vMerge/>').length - 1).toBe(3);
    // The party fills that one cell: the fill covers the merged cell's four cells
    const target = cells[1];
    if (!target) throw new Error('no cell');
    expect(
      refusal(
        doc,
        'bob-1:1',
        { kind: 'table-property', ...EDIT, on: [addressOf(target)], op: FILL },
        'Bob'
      )
    ).toBe('applied');
    expect(filled(doc)).toBe(4);
  });

  test('a fill named by the merged cell alone covers its cells; a plain cell only itself', () => {
    const doc = loadOrderedPackage(packageOf(FILE));
    const [restart, , plain] = firstCells(doc);
    if (!restart || !plain) throw new Error('no cells');
    apply(
      doc,
      'bob-1:1',
      { kind: 'table-property', ...EDIT, on: [addressOf(restart)], op: FILL },
      'Bob'
    );
    expect(filled(doc)).toBe(2);
    apply(
      doc,
      'bob-1:2',
      { kind: 'table-property', ...EDIT, on: [addressOf(plain)], op: FILL },
      'Bob'
    );
    expect(filled(doc)).toBe(3);
  });
});
