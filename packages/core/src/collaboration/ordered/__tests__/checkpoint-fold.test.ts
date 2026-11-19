// A fold changes nothing the room uses (design D14): random sessions replayed from the room's log,
// folded at the end of rounds, give back every shown unit and its labels, every pending suggestion
// whole, its card, who typed its units, the takeovers, comments, part revisions, what a page
// shows, the Word file, and the Word file after an accept and after a reject of each suggestion.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { apply, at, DATE, DIRECT, EDIT, textOf } from './change-apply-fixture.ts';
import { checkFold, folded, mediaStore } from './fold-check.ts';
import { SESSION_FILE } from './sessions/fixture.ts';
import { runSession } from './sessions/session.ts';
import { bodyOf, docx, MATRIX_ITEMS, p, t } from './word-fixtures.ts';

describe('a fold of a random session', () => {
  for (const seed of [7, 23]) {
    test(`seed ${seed}: the folded copy gives back what the room uses`, async () => {
      const { log, rounds } = runSession(seed, ['Olga', 'Pia', 'Quin'], 240, 3);
      const { document: doc, base } = openOrderedWordFile(SESSION_FILE, { file: 0 });
      const blobs = mediaStore(base);
      let applied = 0;
      for (const end of rounds) {
        for (; applied < end; applied += 1) {
          const change = log[applied]!;
          withUndoRecord(doc, () =>
            applyOrderedChange(doc, change, {
              person: change.by,
              date: change.at,
              ids: orderedIdMinter(change.id),
            })
          );
        }

        await checkFold(doc, base, blobs, `seed ${seed} after ${end} changes`);
      }
    }, 60_000);
  }
});

describe('a fold of a formatted field', () => {
  const pageField = MATRIX_ITEMS.find((item) => item.key === 'pagefield');

  // The room reads a field's face from the run a format reaches, as the engine does: a fold
  // writes a direct format there, not on the run that begins the field
  for (const [label, mode] of [
    ['a direct format', DIRECT],
    ['a suggested format', EDIT],
  ] as const) {
    test(`${label} on a field shows the same face after the fold`, async () => {
      const { document: doc, base } = openOrderedWordFile(docx(bodyOf(pageField!)), { file: 1 });

      // The field's paragraph, between the item's neighbours
      apply(doc, 'olga:1', {
        kind: 'format',
        ...mode,
        from: at(doc, 1, 0),
        to: at(doc, 1, textOf(doc, 1).length),
        properties: [{ localName: 'b' }],
      });
      await checkFold(doc, base, mediaStore(base), label);
    });
  }
});

describe('a fold of a hidden prompt', () => {
  // A suggested format on a prompt that direct typing then filled: the glossary keeps the prompt
  // with its record, and an emptying after the fold gives it back with that record. The prompt
  // starts at `start` in the item's paragraph
  for (const [key, start] of [
    ['placeholder', 6],
    ['blockprompt', 0],
  ] as const) {
    test(`${key}: the prompt comes back with its suggested format when its control empties`, async () => {
      const item = MATRIX_ITEMS.find((candidate) => candidate.key === key);
      const { document: doc, base } = openOrderedWordFile(docx(bodyOf(item!)), { file: 1 });
      apply(doc, 'ana:1', {
        kind: 'format',
        ...EDIT,
        from: at(doc, 1, start + 1),
        to: at(doc, 1, start + 5),
        properties: [{ localName: 'i' }],
      });
      apply(
        doc,
        'olive:2',
        { kind: 'insert', ...DIRECT, place: at(doc, 1, start), text: 'Q' },
        'Olive'
      );
      const blobs = mediaStore(base);
      const once = await folded(doc, base, doc.file + 1, blobs);
      const empty = (copy: OrderedDocument) => {
        apply(
          copy,
          'olive:3',
          { kind: 'delete', ...DIRECT, from: at(copy, 1, start), to: at(copy, 1, start + 1) },
          'Olive'
        );
      };

      // Folded once, and twice, before the control empties
      await checkFold(doc, base, blobs, key, empty);
      await checkFold(once.document, once.base, blobs, `${key} folded twice`, empty);
    });
  }
});

const W16DU = 'http://schemas.microsoft.com/office/word/2023/wordml/word16du';

describe('a fold of a long list', () => {
  test('one suggested deletion over 300 paragraphs', async () => {
    const paragraphs = Array.from({ length: 300 }, (_, k) => p(t(`Paragraph ${k}.`))).join('');
    const { document: doc, base } = openOrderedWordFile(docx(paragraphs), { file: 1 });
    apply(doc, 'olga:1', {
      kind: 'delete',
      ...EDIT,
      from: at(doc, 0, 0),
      to: at(doc, 299, textOf(doc, 299).length),
    });
    await checkFold(doc, base, mediaStore(base), 'a deletion over 300 paragraphs');
  });

  test("a Word file's 20,000 own tracked changes", async () => {
    const mark = (id: number) =>
      `<w:ins w:id="${id}" w:author="Reviewer" w:date="${DATE}" xmlns:w16du="${W16DU}" ` +
      `w16du:dateUtc="${DATE}">${t('x')}</w:ins>`;
    const body = Array.from({ length: 1000 }, (_, k) =>
      p(Array.from({ length: 20 }, (_, j) => mark(k * 20 + j)).join(''))
    ).join('');
    const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
    const again = await folded(doc, base, 2, mediaStore(base));
    expect(again.document.suggestions.size).toBe(20_000);
  });
});

describe("a fold of a Word file's own values", () => {
  const inserted = (attributes: string) =>
    docx(p(`<w:ins ${attributes}>${t('x')}</w:ins>${t('y')}`));
  const many = Array.from({ length: 10 }, (_, k) => `a${k}="1"`).join(' ');

  // Every value a load keeps from a revision mark comes back
  for (const [label, attributes] of [
    ['an empty date', 'w:id="1" w:author="A" w:date=""'],
    ['a long author', `w:id="1" w:author="${'a'.repeat(300)}"`],
    ['a tab in the author', 'w:id="1" w:author="A&#9;B"'],
    ['an attribute with no namespace', 'w:id="1" w:author="A" foo="1"'],
    ['a long attribute value', `w:id="1" w:author="A" w:rsidR="${'0'.repeat(300)}"`],
    ['more attributes than a load keeps', `w:id="1" w:author="A" ${many}`],
    ['a long w:id', `w:id="${'9'.repeat(70)}" w:author="A"`],
  ] as const) {
    test(`an insertion with ${label}`, async () => {
      const { document: doc, base } = openOrderedWordFile(inserted(attributes), { file: 1 });
      await checkFold(doc, base, mediaStore(base), label);
    });
  }

  test('one w:id of two authors: two suggestions, each with its own units', async () => {
    const mark = (author: string, text: string) =>
      `<w:ins w:id="1" w:author="${author}" w:date="${DATE}">${t(text)}</w:ins>`;
    const file = docx(p(mark('A', 'ab') + mark('B', 'cd')));
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const again = await folded(doc, base, 2, mediaStore(base));
    const unitsOf = (copy: OrderedDocument) =>
      [...copy.suggestions.values()].map(({ id, attribution }) => [
        attribution.author,
        [...(copy.labelled.get(id) ?? [])].map((item) => item.text).join(''),
      ]);
    expect(unitsOf(again.document)).toEqual([
      ['A', 'ab'],
      ['B', 'cd'],
    ]);
    expect(unitsOf(again.document)).toEqual(unitsOf(doc));
  });
});
