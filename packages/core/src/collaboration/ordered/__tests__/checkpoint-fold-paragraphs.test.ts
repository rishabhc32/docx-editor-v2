// Every unit is in a paragraph, live and after a fold (design D14). A decision that takes a
// paragraph mark away joins its words to the next paragraph; when a table starts there, that
// paragraph is in the table's first cell that shows, as compose draws the words, the editor
// alone's reject puts them and a fold loads them. Before, the room kept them before the table,
// in no paragraph: a later decision on that cell's row kept them live and took them after a fold.
// A take-back of the table by its author leaves the same words. The organiser's undo of her
// direct table, which would leave them in no paragraph, has no exact inverse: it is refused.

import { describe, expect, test } from 'bun:test';
import { orderedCardOf } from '../card-of.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  APPLY_FILE,
  apply,
  at,
  DIRECT,
  EDIT,
  markOf,
  opens,
  refusal,
  textOf,
} from './change-apply-fixture.ts';
import { checkEveryCard, checkFold, folded, mediaStore } from './fold-check.ts';

/** Reject the card that holds `suggestion`, as the organiser's decision does. */
function reject(doc: OrderedDocument, id: string, suggestion: string): void {
  const seen = orderedCardSeen(doc, suggestion) ?? [];
  apply(doc, id, { kind: 'decide', cards: [{ suggestions: seen }], status: 'rejected' }, 'Olive');
}

/**
 * Ben's new table right after the table (the engine adds a paragraph between) and Ana's words in
 * that paragraph; with `row`, Ana's new row in his table too.
 */
function wordsBeforeNewTable(row: boolean) {
  const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const doc = opened.document;
  const table = { kind: 'insert-table', ...EDIT, rows: 2, columns: 2, columnWidth: 2400 } as const;
  apply(doc, 'ben:1', { ...table, place: at(doc, 8, 0) }, 'Ben');
  apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: markOf(doc, 8), text: 'Ltd' });
  if (row) {
    const below = { kind: 'insert-row', ...EDIT, where: 'below' } as const;
    apply(doc, 'ana:3', { ...below, row: addressOf(opens(doc, 'tr')[2]!) });
  }

  expect(textOf(doc, 8)).toBe('Ltd');
  return { doc, base: opened.base, blobs: mediaStore(opened.base) };
}

describe('words whose paragraph mark a decision took away before a table', () => {
  test('go into the first cell that shows, and with its row', async () => {
    const { doc, base, blobs } = wordsBeforeNewTable(true);
    await checkEveryCard(doc, base, blobs, 'before the decision');
    reject(doc, 'olive:4', 'S:ben:1.1');

    // Ana's words in the first cell of her row, which joins them to its card
    expect(textOf(doc, 8)).toBe('Ltd');
    expect(orderedCardOf(doc, 'S:ana:3.1')?.suggestions).toEqual(['S:ana:2.1', 'S:ana:3.1']);
    await checkFold(doc, base, blobs, 'after the reject of the table');
    await checkEveryCard(doc, base, blobs, 'after the reject of the table');
    reject(doc, 'olive:5', 'S:ana:3.1');
    expect(textOf(doc, 8)).toBe('Due in 30 days.');
    await checkFold(doc, base, blobs, 'after the reject of the row');
  });

  test('go into the next paragraph when the whole table goes', async () => {
    const { doc, base, blobs } = wordsBeforeNewTable(false);
    reject(doc, 'olive:4', 'S:ben:1.1');
    expect(textOf(doc, 8)).toBe('LtdDue in 30 days.');
    await checkFold(doc, base, blobs, 'after the reject of the table');
    await checkEveryCard(doc, base, blobs, 'after the reject of the table');
  });
});

describe('words whose paragraph mark a take-back took away before a table', () => {
  test('go into the next paragraph when its author takes the new table back', async () => {
    const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const doc = opened.document;
    const table = {
      kind: 'insert-table',
      ...EDIT,
      rows: 2,
      columns: 2,
      columnWidth: 2400,
    } as const;
    apply(doc, 'ben:1', { ...table, place: at(doc, 8, 0) }, 'Ben');
    apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: markOf(doc, 8), text: 'Ltd' });
    apply(doc, 'ben:2', { kind: 'retract', suggestions: ['S:ben:1.1'] }, 'Ben');
    expect(textOf(doc, 8)).toBe('LtdDue in 30 days.');
    await checkFold(doc, opened.base, mediaStore(opened.base), 'the new table taken back');
  });
});

describe("the organiser's undo of her direct table that holds another person's words", () => {
  const undo = (status: 'accepted' | 'rejected') =>
    ({ kind: 'revert', suggestions: ['S:olive:1.1'], status }) as const;

  /** Olive's direct table after the table, and Ana's words in the paragraph the engine adds. */
  function tableWithWords(doc: OrderedDocument): void {
    const table = {
      kind: 'insert-table',
      ...DIRECT,
      rows: 2,
      columns: 2,
      columnWidth: 2400,
    } as const;
    apply(doc, 'olive:1', { ...table, place: at(doc, 8, 0) }, 'Olive');
    apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: markOf(doc, 8), text: 'Ltd' });
  }

  test('is refused by name, live and after a fold', async () => {
    const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const copy = await folded(opened.document, opened.base, 2, mediaStore(opened.base));
    for (const doc of [opened.document, copy.document]) {
      tableWithWords(doc);
      expect(refusal(doc, 'olive:2', undo('rejected'), 'Olive')).toBe('no-inverse');
      expect(textOf(doc, 8)).toBe('Ltd');
    }
  });

  test('works once a decision took the words away, and its redo is exact', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    tableWithWords(doc);
    reject(doc, 'olive:2', 'S:ana:2.1');
    const shown = () => [8, 9, 10].map((k) => textOf(doc, k));
    const before = shown();
    apply(doc, 'olive:3', undo('rejected'), 'Olive');
    expect(textOf(doc, 8)).toBe('Due in 30 days.');
    apply(doc, 'olive:4', undo('accepted'), 'Olive');
    expect(shown()).toEqual(before);
    await checkFold(doc, base, mediaStore(base), 'undone and made again');
  });
});
