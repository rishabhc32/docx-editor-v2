// A decision on a paragraph mark's own insertion or deletion, with real editors and one ordered
// room, against the editor alone: a pasted paragraph, a new table and a new row, and a join and
// an Enter in a file paragraph with empty properties. Each is accepted and rejected live, after a
// save and reopen, and after one and two folds.
//
// The decision leaves no properties that held nothing but the mark's revision markers, as the
// editor alone leaves none. The room once kept `<w:pPr><w:rPr/></w:pPr>` on an accepted pasted
// paragraph and on the cells of an accepted table or row, and the engine kept it after a reopen:
// both read the empty `w:rPr` as the mark's own properties.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import { decidedCheck, type Act } from './decided-harness.ts';
import { caretAt, closeAll, surface } from './editor-harness.ts';
import { copiedBetween } from './paste-harness.ts';
import { TABLE_FILE, caretIn, tableCommand } from './tables-harness.ts';

afterEach(closeAll);

const LINE = docx(p(t('Line: Revenue')) + p(t('Signed.')));

/** The body as a reader of the file sees it: no ids, dates, sections or run properties. */
const bodyOf = (xml: string): string =>
  (/<w:body>(.*)<\/w:body>/s.exec(xml)?.[1] ?? 'no body')
    .replace(/<w:sectPr.*?<\/w:sectPr>/s, '')
    .replace(/ w:(id|date|rsidR|rsidRPr|rsidDel)="[^"]*"| w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ xml:space="preserve"|<w:rPr>.*?<\/w:rPr>/g, '')
    .replace(/<\/w:r><w:r>/g, '')
    .replace(/<\/w:t><w:t>/g, '');

const check = decidedCheck(bodyOf);

/** Paste the paragraphs of `source` before "Revenue", as a suggestion. */
const paste =
  (source: Uint8Array, last: number, end?: number): Act =>
  (person) => {
    person.editor.setMode('suggesting');
    caretAt(person, 0, 6);
    const copy = copiedBetween(0, last, end, source);
    expect(surface(person).pasteRich(copy.text, copy.html)).toBe(true);
  };

const newTable: Act = (person) => {
  person.editor.setMode('suggesting');
  caretAt(person, 0, 6);
  expect(surface(person).insertTable(2, 2)).toBe(true);
};

const newRow: Act = (person) => {
  person.editor.setMode('suggesting');
  caretIn(person, '100');
  expect(tableCommand(person, { type: 'insertRow', where: 'below' }).ok).toBe(true);
};

/** Join the first two paragraphs, as a suggestion: the first one's mark is struck. */
const joinFirst: Act = (person) => {
  person.editor.setMode('suggesting');
  caretAt(person, 1, 0);
  surface(person).deleteBackward();
};

const enter: Act = (person) => {
  person.editor.setMode('suggesting');
  caretAt(person, 0, 6);
  surface(person).splitParagraph?.();
};

describe("a party's new paragraph marks, decided as the editor alone decides them", () => {
  for (const [label, file, step] of [
    ['a pasted paragraph', LINE, paste(docx(p(t('[x]'))), 0)],
    ['two pasted paragraphs', LINE, paste(docx(p(t('[x]')) + p(t('yz'))), 1, 2)],
    ['a new table', TABLE_FILE, newTable],
    ['a new row', TABLE_FILE, newRow],
  ] as const) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        await check(label, file, [step], decision, 'ben');
      });
    }
  }
});

describe('a join and an Enter in a file paragraph with empty properties, as alone', () => {
  for (const [label, properties] of [
    ['empty properties', '<w:pPr/>'],
    ['empty mark properties', '<w:pPr><w:rPr/></w:pPr>'],
    ['alignment and empty mark properties', '<w:pPr><w:jc w:val="center"/><w:rPr/></w:pPr>'],
  ] as const) {
    const file = docx(`<w:p>${properties}${t('Line: Revenue')}</w:p>` + p(t('Signed.')));
    for (const [gesture, step] of [
      ['joined', joinFirst],
      ['split', enter],
    ] as const) {
      for (const decision of ['accept', 'reject'] as const) {
        for (const actor of ['olive', 'ben'] as const) {
          test(`${label}, ${gesture} by ${actor}, ${decision}ed`, async () => {
            await check(`${label}, ${gesture}`, file, [step], decision, actor);
          });
        }
      }
    }
  }
});
