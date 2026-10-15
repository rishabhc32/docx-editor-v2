// Suggesting mode records table edits in Word's tracked form.
//
// A cell property write keeps the new `w:tcPr` and records the old one in a `w:tcPrChange`
// (§17.13.5.37), the cell twin of `w:rPrChange`. A whole-table deletion keeps the table and
// marks every row deleted, the form a tracked `deleteTableRow` writes for one row. Before,
// both landed in suggesting mode as permanent edits with nothing to accept or reject.

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
import { applyTreeOp, validateTreeOp } from '../store/tree-ops.ts';
import { OTHER_AUTHORS_ROW_CHANGE } from '../store/tree-op-tracked-table-properties.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = WML_NAMESPACE_URI;
const BEA = { author: 'Bea', date: '2026-10-01T00:00:00Z' };
const STAMP = 'w:author="Bea" w:date="2026-10-01T00:00:00Z"';
const TC_W = '<w:tcW w:w="2400" w:type="dxa"/>';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
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
/** A table with one grid column per cell of its first row, and a paragraph after it. */
const TABLE = (...rows: string[]): string =>
  `<w:tbl><w:tblGrid>${'<w:gridCol w:w="2400"/>'.repeat(rows[0]!.split('<w:tc>').length - 1)}` +
  `</w:tblGrid>${rows.join('')}</w:tbl><w:p><w:r><w:t>after</w:t></w:r></w:p>`;

const tableId = (part: OoxmlPart): string => collectByKind(part.root, 'table')[0]!.id;
const cellIds = (part: OoxmlPart): string[] =>
  collectByKind(part.root, 'tableCell').map((c) => c.id);

function apply(part: OoxmlPart, op: TreeDocOp): OoxmlPart {
  const result = applyTreeOp(part, op);
  if (!result.ok) throw new Error(`${op.op} refused: ${result.reason}`);
  return result.part;
}

/** Every serialized `<w:tcPr>…</w:tcPr>` in the part, in document order. */
function cellProperties(part: OoxmlPart): string[] {
  return serializeOoxmlPart(part).match(/<w:tcPr>.*?<\/w:tcPr>(?=<w:p)/g) ?? [];
}

const fill = (part: OoxmlPart, hex: string, cells = cellIds(part).slice(0, 1)): TreeDocOp => ({
  op: 'setTableCellFill',
  tableId: tableId(part),
  cellIds: cells,
  color: { kind: 'hex', value: hex },
  revision: BEA,
});

describe('tracked cell properties write w:tcPrChange', () => {
  test('a fill keeps the new w:tcPr and records the old one as its last child', () => {
    const part = load(TABLE(ROW(CELL('a1'), CELL('a2')), ROW(CELL('b1'), CELL('b2'))));
    const next = apply(part, fill(part, 'FF0000'));
    const record = `<w:tcPrChange ${STAMP} w:id="0"><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr></w:tcPrChange>`;
    const plain = '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>';
    // Rejecting a row's records resets a cell of that row without one, so the neighbour
    // records its unchanged properties too. The other row is not part of the decision.
    expect(cellProperties(next)).toEqual([
      `<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:shd w:fill="FF0000" w:val="clear"/>${record}</w:tcPr>`,
      `<w:tcPr><w:tcW w:type="dxa" w:w="2400"/>${record}</w:tcPr>`,
      plain,
      plain,
    ]);
  });

  test('vertical alignment and borders record the previous w:tcPr the same way', () => {
    const part = load(TABLE(ROW(CELL('a'), CELL('b'))));
    const [first] = cellIds(part);
    const record = `<w:tcPrChange ${STAMP} w:id="0"><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr></w:tcPrChange>`;
    const aligned = apply(part, {
      op: 'setTableCellVerticalAlignment',
      tableId: tableId(part),
      cellIds: [first!],
      alignment: 'center',
      revision: BEA,
    });
    expect(cellProperties(aligned)[0]).toBe(
      `<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:vAlign w:val="center"/>${record}</w:tcPr>`
    );
    const bordered = apply(part, {
      op: 'setTableCellBorders',
      tableId: tableId(part),
      cellIds: [first!],
      scope: 'top',
      spec: { style: 'single', size: 8, color: { kind: 'hex', value: '000000' } },
      revision: BEA,
    });
    expect(cellProperties(bordered)[0]).toBe(
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:tcBorders>' +
        '<w:top w:color="000000" w:sz="8" w:val="single"/></w:tcBorders>' +
        `${record}</w:tcPr>`
    );
  });

  test('a write that changes no property records nothing', () => {
    const part = load(TABLE(ROW(CELL('a', `${TC_W}<w:shd w:val="clear" w:fill="FF0000"/>`))));
    const next = apply(part, fill(part, 'FF0000'));
    expect(serializeOoxmlPart(next)).toBe(serializeOoxmlPart(part));
  });

  test('a second change keeps the ORIGINAL as the record', () => {
    const part = load(TABLE(ROW(CELL('a'))));
    const red = apply(part, fill(part, 'FF0000'));
    const blue = apply(red, fill(red, '0000FF'));
    expect(cellProperties(blue)).toEqual([
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:shd w:fill="0000FF" w:val="clear"/>' +
        `<w:tcPrChange ${STAMP} w:id="1"><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>` +
        '</w:tcPrChange></w:tcPr>',
    ]);
  });

  test('a change back to the original drops the records of the whole row', () => {
    const part = load(TABLE(ROW(CELL('a', `${TC_W}<w:vAlign w:val="top"/>`), CELL('b'))));
    const align = (from: OoxmlPart, alignment: 'top' | 'center'): OoxmlPart =>
      apply(from, {
        op: 'setTableCellVerticalAlignment',
        tableId: tableId(from),
        cellIds: cellIds(from).slice(0, 1),
        alignment,
        revision: BEA,
      });
    const centered = align(part, 'center');
    expect(cellProperties(centered).filter((xml) => xml.includes('<w:tcPrChange'))).toHaveLength(2);
    expect(serializeOoxmlPart(align(centered, 'top'))).toBe(serializeOoxmlPart(part));
  });

  test("a fill beside another author's record in the row is refused", () => {
    const theirs =
      '<w:tcPrChange w:id="7" w:author="Ada"><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr></w:tcPrChange>';
    const part = load(
      TABLE(ROW(CELL('a'), CELL('b', `${TC_W}<w:vAlign w:val="bottom"/>${theirs}`)))
    );
    // Rejecting Bea's record and then Ada's would reset the first cell to `tcW auto`.
    expect(applyTreeOp(part, fill(part, 'FF0000'))).toEqual({
      ok: false,
      reason: 'unsupported-revision',
      detail: OTHER_AUTHORS_ROW_CHANGE,
    });
  });

  test('one op over several cells is one revision', () => {
    const part = load(TABLE(ROW(CELL('a'), CELL('b'))));
    const next = apply(part, fill(part, 'FF0000', cellIds(part)));
    const ids = collectByKind(next.root, 'generic')
      .filter((node) => node.localName === 'tcPrChange')
      .map((node) => node.attributes.find((attribute) => attribute.localName === 'id')?.value);
    expect(ids).toEqual(['0', '0']);
  });

  test("a cell's own revision marker stays on the cell, out of the record", () => {
    const marker = '<w:cellIns w:id="5" w:author="Ada"/>';
    const part = load(TABLE(ROW(CELL('a', `${TC_W}${marker}`))));
    const next = apply(part, fill(part, 'FF0000'));
    expect(cellProperties(next)[0]).toBe(
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:shd w:fill="FF0000" w:val="clear"/>' +
        '<w:cellIns w:author="Ada" w:id="5"/>' +
        `<w:tcPrChange ${STAMP} w:id="6"><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>` +
        '</w:tcPrChange></w:tcPr>'
    );
  });

  test("a cell in the author's own proposed row needs no record", () => {
    const row = `<w:tr><w:trPr><w:ins w:id="5" ${STAMP}/></w:trPr>${CELL('a')}</w:tr>`;
    const part = load(TABLE(row));
    const next = apply(part, fill(part, 'FF0000'));
    expect(cellProperties(next)).toEqual([
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:shd w:fill="FF0000" w:val="clear"/></w:tcPr>',
    ]);
  });

  test('Reject restores the old properties and Accept keeps the new ones', () => {
    const part = load(TABLE(ROW(CELL('a1'), CELL('a2')), ROW(CELL('b1'), CELL('b2'))));
    const next = apply(part, fill(part, 'FF0000'));
    const rejected = apply(next, { op: 'rejectAllRevisions' });
    expect(canonicalOoxmlFingerprint(rejected.root)).toBe(canonicalOoxmlFingerprint(part.root));
    const plain = '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>';
    expect(cellProperties(apply(next, { op: 'acceptAllRevisions' }))).toEqual([
      '<w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:shd w:fill="FF0000" w:val="clear"/></w:tcPr>',
      plain,
      plain,
      plain,
    ]);
  });

  test('an attribution without an author is refused', () => {
    const part = load(TABLE(ROW(CELL('a'))));
    expect(validateTreeOp(part, { ...fill(part, 'FF0000'), revision: { author: ' ' } })).toBe(
      'invalid-property-value'
    );
  });
});

describe('a tracked table deletion marks every row deleted', () => {
  const deleteTable = (part: OoxmlPart): TreeDocOp => ({
    op: 'deleteBlock',
    blockId: tableId(part),
    revision: BEA,
  });

  test('the table stays, with w:del on every row and w:cellDel in every cell', () => {
    const part = load(TABLE(ROW(CELL('a1'), CELL('a2')), ROW(CELL('b1'), CELL('b2'))));
    const xml = serializeOoxmlPart(apply(part, deleteTable(part)));
    const cell = (text: string): string =>
      `<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2400"/><w:cellDel ${STAMP} w:id="0"/></w:tcPr>` +
      `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const row = (...cells: string[]): string =>
      `<w:tr><w:trPr><w:del ${STAMP} w:id="0"/></w:trPr>${cells.join('')}</w:tr>`;
    expect(xml).toContain(
      '<w:tbl><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>' +
        `${row(cell('a1'), cell('a2'))}${row(cell('b1'), cell('b2'))}</w:tbl>`
    );
  });

  test('the last row is no obstacle, though a tracked row deletion keeps it', () => {
    const part = load(TABLE(ROW(CELL('only'))));
    const rowId = collectByKind(part.root, 'tableRow')[0]!.id;
    expect(
      validateTreeOp(part, {
        op: 'deleteTableRow',
        tableId: tableId(part),
        rowId,
        revision: BEA,
      })
    ).toBe('block-required');
    const xml = serializeOoxmlPart(apply(part, deleteTable(part)));
    expect(xml).toContain(`<w:tr><w:trPr><w:del ${STAMP} w:id="0"/></w:trPr>`);
  });

  test('a row already proposed for deletion keeps its own decision', () => {
    const pending = `<w:tr><w:trPr><w:del w:id="3" w:author="Ada"/></w:trPr>${CELL('a')}</w:tr>`;
    const part = load(TABLE(pending, ROW(CELL('b'))));
    const xml = serializeOoxmlPart(apply(part, deleteTable(part)));
    expect(xml).toContain('<w:tr><w:trPr><w:del w:author="Ada" w:id="3"/></w:trPr>');
    expect(xml).toContain(`<w:tr><w:trPr><w:del ${STAMP} w:id="4"/></w:trPr>`);
  });

  test("the author's own proposed table is withdrawn outright; any other table stays", () => {
    const row = `<w:tr><w:trPr><w:ins w:id="5" ${STAMP}/></w:trPr>${CELL('mine')}</w:tr>`;
    const own = load(TABLE(row));
    expect(collectByKind(apply(own, deleteTable(own)).root, 'table')).toHaveLength(0);
    const plain = load(TABLE(ROW(CELL('a'))));
    expect(collectByKind(apply(plain, deleteTable(plain)).root, 'table')).toHaveLength(1);
  });

  test("the author's own proposed row is withdrawn, and every other row proposed", () => {
    const original = TABLE(ROW(CELL('a')), ROW(CELL('b')));
    /** The table with a row Bea proposed between its two rows, holding `content`. */
    const withOwnRow = (content = ''): OoxmlPart =>
      load(
        original.replace(
          '</w:tr><w:tr>',
          `</w:tr><w:tr><w:trPr><w:ins w:id="5" ${STAMP}/></w:trPr>` +
            `<w:tc><w:p><w:r><w:t>mine</w:t></w:r>${content}</w:p></w:tc></w:tr><w:tr>`
        )
      );
    const part = withOwnRow();
    const next = apply(part, deleteTable(part));
    const xml = serializeOoxmlPart(next);
    expect(xml).not.toContain('mine');
    expect(xml.match(new RegExp(`<w:trPr><w:del ${STAMP} w:id="6"/>`, 'g'))).toHaveLength(2);
    // Reject leaves the table as it was before the withdrawn proposal; Accept removes it.
    const rejected = serializeOoxmlPart(apply(next, { op: 'rejectAllRevisions' }));
    expect(rejected.replaceAll('<w:trPr/>', '')).toBe(serializeOoxmlPart(load(original)));
    expect(collectByKind(apply(next, { op: 'acceptAllRevisions' }).root, 'table')).toHaveLength(0);
    // Withdrawing removes the row, so another author's pending change in it is refused.
    const held = withOwnRow('<w:ins w:id="7" w:author="Ada"><w:r><w:t>typed</w:t></w:r></w:ins>');
    expect(validateTreeOp(held, deleteTable(held))).toBe('unsupported-revision');
  });

  test("another author's pending row insertion is refused, as a tracked row deletion refuses it", () => {
    const row = `<w:tr><w:trPr><w:ins w:id="5" w:author="Ada"/></w:trPr>${CELL('theirs')}</w:tr>`;
    const part = load(TABLE(ROW(CELL('a')), row));
    expect(validateTreeOp(part, deleteTable(part))).toBe('unsupported-revision');
  });

  test('any other block refuses the attribution', () => {
    const part = load(TABLE(ROW(CELL('a'))));
    const paragraph = collectByKind(part.root, 'paragraph').at(-1)!;
    expect(validateTreeOp(part, { op: 'deleteBlock', blockId: paragraph.id, revision: BEA })).toBe(
      'unsupported-revision'
    );
  });

  test('Reject restores the table and Accept removes it', () => {
    const part = load(TABLE(ROW(CELL('a1'), CELL('a2')), ROW(CELL('b1'), CELL('b2'))));
    const next = apply(part, deleteTable(part));
    // Rejecting leaves the emptied `w:trPr` behind, as it does for a tracked row deletion.
    const rejected = apply(next, { op: 'rejectAllRevisions' });
    expect(serializeOoxmlPart(rejected)).not.toMatch(/<w:(del|cellDel)\b/);
    expect(serializeOoxmlPart(rejected).replaceAll('<w:trPr/>', '')).toBe(serializeOoxmlPart(part));
    const accepted = apply(next, { op: 'acceptAllRevisions' });
    expect(collectByKind(accepted.root, 'table')).toHaveLength(0);
  });
});
