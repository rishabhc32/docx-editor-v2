import { expect, test } from 'bun:test';
import {
  applyTreeOp,
  readOoxmlPart,
  revisionItemsOf,
  reviewItemKey,
  serializeOoxmlPart,
  type OoxmlPart,
  type TreeDocOp,
} from '../index.ts';
import { planRevisionBatch } from '../store/revision-batch.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (properties: string, text: string) => `<w:tr>${properties}${cell(text)}</w:tr>`;

/** A lone row marker, a marker beside other row properties, and an untouched `<w:trPr/>`. */
const ROWS = [
  row('<w:trPr><w:ins w:id="1" w:author="Ada"/></w:trPr>', 'inserted'),
  row('<w:trPr><w:del w:id="2" w:author="Ada"/></w:trPr>', 'deleted'),
  row('<w:trPr><w:trHeight w:val="500"/><w:del w:id="3" w:author="Ada"/></w:trPr>', 'tall'),
  row('<w:trPr><w:cantSplit/><w:ins w:id="4" w:author="Ada"/></w:trPr>', 'unsplit'),
  row('<w:trPr/>', 'authored'),
];
/** The rows each action keeps. */
const KEPT = {
  reject: [
    row('', 'deleted'),
    row('<w:trPr><w:trHeight w:val="500"/></w:trPr>', 'tall'),
    row('<w:trPr/>', 'authored'),
  ],
  accept: [
    row('', 'inserted'),
    row('<w:trPr><w:cantSplit/></w:trPr>', 'unsplit'),
    row('<w:trPr/>', 'authored'),
  ],
};

function load(rows: readonly string[]): OoxmlPart {
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${rows.join('')}</w:tbl><w:p/></w:body></w:document>`,
    { name: '/word/document.xml', contentType: 'application/xml' }
  );
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}

function apply(part: OoxmlPart, ops: readonly TreeDocOp[]): OoxmlPart {
  for (const op of ops) {
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    part = result.part;
  }
  return part;
}

function rowsOf(part: OoxmlPart): string {
  const xml = serializeOoxmlPart(part);
  return xml.slice(xml.indexOf('</w:tblGrid>') + '</w:tblGrid>'.length, xml.indexOf('</w:tbl>'));
}

for (const action of ['accept', 'reject'] as const) {
  test(`${action}: one decision at a time drops the row properties only its marker held`, () => {
    let part = load(ROWS);
    // The review pane resolves one card per call and reads the queue again after each.
    const count = revisionItemsOf(part).length;
    for (let index = 0; index < count; index += 1) {
      const plan = planRevisionBatch(part, action, [reviewItemKey(revisionItemsOf(part)[0]!)]);
      expect(plan.result.skipped).toEqual([]);
      part = apply(part, plan.ops);
    }
    expect(revisionItemsOf(part)).toHaveLength(0);
    expect(rowsOf(part)).toBe(KEPT[action].join(''));
  });

  test(`${action}: the whole-story decision drops them the same way`, () => {
    const op = action === 'accept' ? 'acceptAllRevisions' : 'rejectAllRevisions';
    const part = apply(load(ROWS), [{ op }]);
    expect(revisionItemsOf(part)).toHaveLength(0);
    expect(rowsOf(part)).toBe(KEPT[action].join(''));
  });
}

test('a tracked row the editor wrote resolves back to a row without row properties', () => {
  for (const [op, action] of [
    ['deleteTableRow', 'reject'],
    ['insertTableRow', 'accept'],
  ] as const) {
    const source = load([row('', 'first'), row('', 'second')]);
    const body = source.root.children[0];
    const table = body?.kind === 'textValue' ? undefined : body?.children[0];
    if (table?.kind !== 'table') throw new Error('no table');
    const rowId = table.children.filter((child) => child.kind === 'tableRow')[1]!.id;
    const revision = { author: 'Ada' };
    const tracked = apply(source, [
      op === 'deleteTableRow'
        ? { op, tableId: table.id, rowId, revision }
        : { op, tableId: table.id, rowId, where: 'below', revision },
    ]);
    expect(rowsOf(tracked)).toContain('<w:trPr>');
    const plan = planRevisionBatch(tracked, action, revisionItemsOf(tracked).map(reviewItemKey));
    expect(plan.result.skipped).toEqual([]);
    expect(rowsOf(apply(tracked, plan.ops))).not.toContain('<w:trPr');
  }
});
