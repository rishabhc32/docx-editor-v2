// A note shows only while its reference shows: a decision that hides the reference with its row
// leaves the note out of the notes part, as Word removes a note with its reference, so a saved
// file opened again gives the same document. The room shows the same: a note whose reference a
// decision hid goes with the cards in it, and an author's own new note goes with the own new
// cell or control value that takes its reference back.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { isHidden } from '../model/document.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  DIRECT,
  EDIT,
  FOOTNOTES,
  load,
  opens,
  textOf,
} from './change-apply-fixture.ts';
import { packageOf } from './word-fixtures.ts';

describe('notes and their references', () => {
  test('a note whose reference a rejected row hides is left out', () => {
    const doc = load();
    const [row] = opens(doc, 'tr');
    if (!row) throw new Error('fixture');
    apply(
      doc,
      'ben:1',
      { kind: 'insert-row', ...EDIT, row: addressOf(row), where: 'below' },
      'Ben'
    );
    // The new row's first cell is paragraph 6 now: "A1", "B1", then the new row's two cells
    expect([textOf(doc, 4), textOf(doc, 5), textOf(doc, 6)]).toEqual(['A1', 'B1', '']);
    apply(
      doc,
      'olga:1',
      {
        kind: 'insert-note',
        ...EDIT,
        place: at(doc, 6, 0),
        note: 'footnote',
        text: 'In the new row.',
      },
      'Olga'
    );
    const notes = () => serializeOoxmlPart(composeOrderedPart(doc, FOOTNOTES).part);
    expect(notes()).toContain('In the new row.');
    // Ben's row goes; Olga's note is still pending, but its reference shows nowhere
    withUndoRecord(doc, () => setStatus(doc, 'S:ben:1.1', 'rejected'));
    expect(notes()).not.toContain('In the new row.');
  });

  // The same rule after a decision: the room shows no note, and no card, that a reopen loses
  test("a note whose reference an accepted row deletion takes goes too, with Ana's card in it", () => {
    const doc = load();
    const notes = () => serializeOoxmlPart(composeOrderedPart(doc, FOOTNOTES).part);
    const authors = (one: OrderedDocument) =>
      orderedCards(one)
        .map((card) => card.author)
        .sort();
    // Olga's note in cell A1, and Ana's suggested words in it
    apply(
      doc,
      'olga:1',
      { kind: 'insert-note', ...DIRECT, place: at(doc, 4, 2), note: 'footnote', text: 'Row note.' },
      'Olga'
    );
    const index = [0, 1, 2, 3, 4, 5].find((i) => textOf(doc, i, FOOTNOTES).includes('Row note.'));
    if (index === undefined) throw new Error('no note paragraph');
    apply(doc, 'ana:1', {
      kind: 'insert',
      part: FOOTNOTES,
      mode: 'suggest',
      place: at(doc, index, 1, FOOTNOTES),
      text: 'See ',
    });
    const [row] = opens(doc, 'tr');
    if (!row) throw new Error('fixture');
    apply(doc, 'ben:1', { kind: 'delete-row', ...EDIT, row: addressOf(row) }, 'Ben');
    expect(authors(doc)).toEqual(['Ana', 'Ben']);
    const ben = doc.suggestions.get('S:ben:1.1');
    if (!ben) throw new Error('no row deletion');
    apply(
      doc,
      'olga:2',
      {
        kind: 'decide',
        status: 'accepted',
        cards: [{ suggestions: [{ suggestion: ben.id, version: ben.version }] }],
      },
      'Olga'
    );
    // The row goes, its reference with it, and the note with Ana's words, as the engine does
    expect(notes()).not.toContain('Row note.');
    expect(authors(doc)).toEqual([]);
    const reopened = loadOrderedPackage(packageOf(orderedPackageBytes(doc, packageOf(APPLY_FILE))));
    expect(authors(reopened)).toEqual(authors(doc));
  });

  // An author's own new cell, row or control value takes back their own new reference too
  test("a note whose own new reference its author's cell, row or control value takes back goes", () => {
    const authors = (one: OrderedDocument) =>
      orderedCards(one)
        .map((card) => card.author)
        .sort();
    const noteShown = (one: OrderedDocument) =>
      serializeOoxmlPart(composeOrderedPart(one, FOOTNOTES).part).includes('In it.') ||
      [...(one.stories.get(FOOTNOTES)?.tree.items() ?? [])].some(
        (item) => item.kind === 'text' && item.text.includes('In it.') && !isHidden(one, item)
      );
    const note = (one: OrderedDocument, id: string, index: number, n: number) =>
      apply(one, id, {
        kind: 'insert-note',
        ...EDIT,
        place: at(one, index, n),
        note: 'footnote',
        text: 'In it.',
      });
    const reopened = (one: OrderedDocument) =>
      loadOrderedPackage(packageOf(orderedPackageBytes(one, packageOf(APPLY_FILE))));
    // A column: Ana's new cell right of A1 (paragraph 5), her note in it, then the column deleted
    const columns = load();
    const [first] = opens(columns, 'tc');
    if (!first) throw new Error('fixture');
    apply(columns, 'ana:1', {
      kind: 'insert-column',
      ...EDIT,
      cell: addressOf(first),
      where: 'right',
    });
    note(columns, 'ana:2', 5, 0);
    const cell = opens(columns, 'tc').find((c) => c.by === 'S:ana:1.1');
    if (!cell) throw new Error('no new cell');
    apply(columns, 'ana:3', { kind: 'delete-column', ...EDIT, cell: addressOf(cell) });
    // The note's suggestion has no card left, and the saved file no note
    const suggestions = orderedCards(columns).flatMap((card) => card.suggestions);
    expect([noteShown(columns), suggestions.includes('S:ana:2.1')]).toEqual([false, false]);
    expect(noteShown(reopened(columns))).toBe(false);
    // A control value: Ana's note in the control's words, then her new value for it
    const controls = load();
    note(controls, 'ana:1', 2, 9);
    const [control] = opens(controls, 'sdt');
    if (!control) throw new Error('fixture');
    apply(controls, 'ana:2', {
      kind: 'control-value',
      ...EDIT,
      control: addressOf(control),
      value: { kind: 'text', text: 'Beta Ltd' },
    });
    expect([noteShown(controls), authors(controls)]).toEqual([false, ['Ana']]);
    expect(authors(reopened(controls))).toEqual(authors(controls));
    // A row: Ana's new row, her note in it as part of the same suggestion, then the row deleted
    const rows = load();
    const [row] = opens(rows, 'tr');
    if (!row) throw new Error('fixture');
    apply(rows, 'ana:1', { kind: 'insert-row', ...EDIT, row: addressOf(row), where: 'below' });
    apply(rows, 'ana:2', {
      kind: 'insert-note',
      ...EDIT,
      into: 'S:ana:1.1',
      place: at(rows, 6, 0),
      note: 'footnote',
      text: 'In it.',
    });
    expect(noteShown(rows)).toBe(true);
    const made = opens(rows, 'tr').find((r) => r.by === 'S:ana:1.1');
    if (!made) throw new Error('no new row');
    apply(rows, 'ana:3', { kind: 'delete-row', ...EDIT, row: addressOf(made) });
    expect([noteShown(rows), authors(rows)]).toEqual([false, []]);
    expect(authors(reopened(rows))).toEqual([]);
  });
});
