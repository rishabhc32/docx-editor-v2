// Which author owns a table row two raced changes reached, as every replica decides it.
//
// The rule reads the records alone: the earlier `w:date`, then the lesser author, then the
// lesser revision id. It has to give the same answer on every replica, so it never depends on
// a locale, and a record with no readable date comes last.

import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../package/ooxml-tree.ts';
import {
  compareRowRecordOwners,
  recordsTableFormatting,
  ROW_PROPERTY_CONTAINERS,
  rowPropertyRecords,
  rowRecordOwnerOf,
  type RowRecordOwner,
} from '../store/table-row-owner.ts';

const owner = (author: string, date: string, id: string): RowRecordOwner => ({ author, date, id });

describe('compareRowRecordOwners', () => {
  test('the earlier date owns the row, whoever the author is', () => {
    const early = owner('Zoe', '2026-10-16T09:00:00Z', '9');
    const late = owner('Ann', '2026-10-16T09:00:01Z', '1');
    expect(compareRowRecordOwners(early, late)).toBeLessThan(0);
    expect(compareRowRecordOwners(late, early)).toBeGreaterThan(0);
  });

  test('on one date the lesser author owns it, by code unit and not by locale', () => {
    const at = '2026-10-16T09:00:00Z';
    expect(compareRowRecordOwners(owner('Bob', at, '1'), owner('Alice', at, '2'))).toBeGreaterThan(
      0
    );
    // Code-unit order puts every capital first; a locale comparison would not.
    expect(compareRowRecordOwners(owner('Zoe', at, '1'), owner('ann', at, '1'))).toBeLessThan(0);
  });

  test('one author on one date: the lesser revision id, as a number', () => {
    const at = '2026-10-16T09:00:00Z';
    expect(compareRowRecordOwners(owner('Ann', at, '9'), owner('Ann', at, '10'))).toBeLessThan(0);
    expect(compareRowRecordOwners(owner('Ann', at, '10'), owner('Ann', at, '10'))).toBe(0);
  });

  test('a record with no readable date comes after every dated one', () => {
    const dated = owner('Zoe', '2026-10-16T09:00:00Z', '9');
    expect(compareRowRecordOwners(dated, owner('Ann', '', '1'))).toBeLessThan(0);
    expect(compareRowRecordOwners(owner('Ann', 'soon', '1'), dated)).toBeGreaterThan(0);
  });
});

test('a record reads its WML attribution, and a missing attribute reads as empty', () => {
  const W = WML_NAMESPACE_URI;
  expect(
    rowRecordOwnerOf([
      { namespaceUri: W, localName: 'author', value: 'Ann' },
      { namespaceUri: 'urn:other', localName: 'date', value: '2026-10-16T09:00:00Z' },
      { namespaceUri: W, localName: 'id', value: '7' },
    ])
  ).toEqual({ author: 'Ann', date: '', id: '7' });
});

test('every container of a row ends with its record', () => {
  for (const { sequence, record } of Object.values(ROW_PROPERTY_CONTAINERS)) {
    expect(sequence.at(-1)).toBe(record);
  }
  expect(ROW_PROPERTY_CONTAINERS.tblPrEx!.sequence).toEqual([
    'tblW',
    'jc',
    'tblCellSpacing',
    'tblInd',
    'tblBorders',
    'shd',
    'tblLayout',
    'tblCellMar',
    'tblLook',
    'tblPrExChange',
  ]);
});

test("a row's records are its own and its cells', not those of a nested table", () => {
  const W = WML_NAMESPACE_URI;
  const record = (name: string, author: string) =>
    `<w:${name} w:id="1" w:author="${author}"><w:${name.replace('Change', '')}/></w:${name}>`;
  const cell = (inner: string) =>
    `<w:tc><w:tcPr><w:tcW w:w="1" w:type="dxa"/>${record('tcPrChange', 'Cell')}</w:tcPr>${inner}</w:tc>`;
  const nested = `<w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr>${cell('<w:p/>')}</w:tr></w:tbl><w:p/>`;
  const read = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>` +
      `<w:tr><w:tblPrEx><w:tblW w:w="0" w:type="auto"/>${record('tblPrExChange', 'Ex')}</w:tblPrEx>` +
      `<w:trPr><w:trHeight w:val="1"/>${record('trPrChange', 'Row')}</w:trPr>` +
      `${cell('<w:p/>')}${cell(nested)}</w:tr></w:tbl><w:p/></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!read.ok) throw new Error(read.reason);
  const rows: OoxmlElement[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'tableRow') rows.push(node);
    for (const child of node.children) visit(child);
  };
  visit(read.part.root);
  const authors = (row: OoxmlElement) =>
    rowPropertyRecords(row).map((found) => rowRecordOwnerOf(found.attributes).author);
  expect(authors(rows[0]!)).toEqual(['Ex', 'Row', 'Cell', 'Cell']);
  expect(authors(rows[1]!)).toEqual(['Cell']);
});

test('only a table formatting op with a revision proposes table formatting', () => {
  const revision = { author: 'Ann' };
  const fill = { op: 'setTableCellFill', tableId: 't', cellIds: ['c'], fill: 'FF0000' } as const;
  expect(recordsTableFormatting({ ...fill, revision } as never)).toBe(true);
  expect(recordsTableFormatting(fill as never)).toBe(false);
  expect(recordsTableFormatting({ ...fill, revision: undefined } as never)).toBe(false);
  expect(
    recordsTableFormatting({
      op: 'insertTableColumn',
      tableId: 't',
      gridColumnId: 'g',
      where: 'right',
      revision,
    } as never)
  ).toBe(false);
  expect(
    recordsTableFormatting({ op: 'setTableProperties', tableId: 't', revision } as never)
  ).toBe(true);
  expect(recordsTableFormatting({ op: 'insertText', revision } as never)).toBe(false);
});
