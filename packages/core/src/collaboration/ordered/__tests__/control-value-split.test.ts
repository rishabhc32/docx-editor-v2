// A value in the second half of a control that Enter split: the half's open item carries the
// Enter's suggestion, which is no insertion of the control, so the value is a suggestion of its
// own, as in any control the file holds. Before, it joined the Enter's suggestion with an
// untracked record, and a reject of the Enter took the value with it.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { isSplitEdge } from '../model/document.ts';
import { addressOf, apply, at, EDIT, opens } from './change-apply-fixture.ts';
import { docx, p, t } from './word-fixtures.ts';

const FILE = docx(
  p(
    t('Notes: ') +
      '<w:sdt><w:sdtPr><w:id w:val="9"/><w:text w:multiLine="1"/></w:sdtPr>' +
      `<w:sdtContent>${t('Acme Ltd')}</w:sdtContent></w:sdt>`
  )
);

describe('a value in a control that Enter split', () => {
  test('is a suggestion of its own, not part of the Enter', () => {
    const { document: doc } = openOrderedWordFile(FILE, { file: 1 });
    const enter = apply(doc, 'ana:1', { kind: 'split', ...EDIT, place: at(doc, 0, 11) });
    const half = opens(doc, 'sdt').find(isSplitEdge);
    expect(half).toBeDefined();
    const value = { kind: 'text', text: 'Beta Ltd' } as const;
    const control = addressOf(half!);
    const done = apply(doc, 'ana:2', { kind: 'control-value', ...EDIT, control, value });
    expect(done.suggestions).not.toContain(enter.suggestions[0]);
    expect(doc.suggestions.size).toBe(2);
  });
});
