// A section cannot end inside a table cell: the engine refuses `setSectionMark` there
// (`tree-op-validate.ts`), as Word writes no section in a cell and a read ignores one. The room
// refuses the break by the engine's name, direct or suggested on its author's own new mark, and
// changes nothing; a break outside the table still goes on.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { apply, at, DIRECT, EDIT, markOf, refusal } from './change-apply-fixture.ts';
import { docx, p, t } from './word-fixtures.ts';

const FILE = docx(
  p(t('Before.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>' +
    p(t('In a cell.')) +
    '</w:tc></w:tr></w:tbl>' +
    p(t('After.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
);

const open = () => openOrderedWordFile(FILE, { file: 1 }).document;

describe('a section break in a table cell', () => {
  test("the organiser's direct break is refused by the engine's name", () => {
    const doc = open();
    const change = { kind: 'section-break', ...DIRECT, mark: markOf(doc, 1) } as const;
    expect(refusal(doc, 'olive:1', change, 'Olive')).toBe('engine-refused');
  });

  test("a party's break on their own new mark in the cell is refused by the engine's name", () => {
    const doc = open();
    apply(doc, 'ana:1', { kind: 'split', ...EDIT, place: at(doc, 1, 3) }, 'Ana');
    const change = { kind: 'section-break', ...EDIT, mark: markOf(doc, 1) } as const;
    expect(refusal(doc, 'ana:2', change, 'Ana')).toBe('engine-refused');
  });

  test('a break after the table goes on', () => {
    const doc = open();
    const change = { kind: 'section-break', ...DIRECT, mark: markOf(doc, 2) } as const;
    expect(refusal(doc, 'olive:1', change, 'Olive')).toBe('applied');
  });
});
