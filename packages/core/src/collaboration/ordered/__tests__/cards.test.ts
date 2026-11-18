// The card of one suggestion (`orderedCardOf`) is the card `orderedCards` gives, after random
// changes of every kind by three people; and finding it costs what the card and the units
// around it hold, not what the document holds.

import { describe, expect, test } from 'bun:test';
import { orderedCardOf } from '../card-of.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { leftOutNoteItems } from '../compose/note-references.ts';
import { orderedCards } from '../decide.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { addressOf, apply, at, BODY, EDIT, load, opens, textOf } from './change-apply-fixture.ts';
import { seededRandom } from './random.ts';
import { sessionDocument } from './sessions/fixture.ts';
import { randomChange, type Who } from './sessions/random-change.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const DATE = '2026-10-28T09:30:00Z';

function expectSameCards(doc: OrderedDocument, label: string): number {
  const byMember = new Map<string, readonly string[]>();
  for (const card of orderedCards(doc))
    for (const id of card.suggestions) byMember.set(id, card.suggestions);
  const leftOut = leftOutNoteItems(doc);
  let checked = 0;
  for (const suggestion of doc.suggestions.values()) {
    if (suggestion.status !== 'pending') continue;
    const local = orderedCardOf(doc, suggestion.id)?.suggestions;
    // A suggestion off the list is its own card, unless it is in a note no reference shows
    const items = [...(doc.labelled.get(suggestion.id) ?? [])];
    const alone = items.some((item) => leftOut.has(item)) ? undefined : [suggestion.id];
    const global = byMember.get(suggestion.id) ?? alone;
    expect({ label, suggestion: suggestion.id, card: local }).toEqual({
      label,
      suggestion: suggestion.id,
      card: global,
    });
    checked += 1;
  }
  return checked;
}

describe('the card of one suggestion', () => {
  for (const seed of [3, 17, 29, 41]) {
    test(`seed ${seed}: it is the card the whole list gives, after random changes`, () => {
      const doc = sessionDocument();
      const rand = seededRandom(seed);
      const people: readonly Who[] = [
        { author: 'Ana', organiser: false },
        { author: 'Ben', organiser: false },
        { author: 'Olga', organiser: true },
      ];
      let applied = 0;
      let checked = 0;
      for (let step = 1; step <= 260; step += 1) {
        const who = people[Math.floor(rand() * people.length)] ?? people[0];
        if (!who) continue;
        const body = randomChange(doc, who, rand);
        if (!body) continue;
        const id = `${who.author.toLowerCase()}:${step}`;
        const change = { v: 1, id, ...body } as OrderedChange;
        try {
          withUndoRecord(doc, () =>
            applyOrderedChange(doc, change, {
              person: testPerson(who.author),
              date: DATE,
              ids: orderedIdMinter(id),
            })
          );
          applied += 1;
        } catch (error) {
          if (!(error instanceof OrderedRefusal)) throw error;
        }
        if (step % 20 === 0) checked += expectSameCards(doc, `seed ${seed} step ${step}`);
      }
      expect(applied).toBeGreaterThan(100);
      expect(checked).toBeGreaterThan(100);
    });
  }

  test('a card costs what it holds, not what the document holds', () => {
    const steps = (paragraphs: number) => {
      const body = Array.from({ length: paragraphs }, (_, k) =>
        p(t(`Clause ${k} of the agreement.`))
      );
      const doc = loadOrderedPackage(packageOf(docx(body.join(''))));
      const middle = Math.floor(paragraphs / 2);
      apply(doc, 'ana:1', {
        kind: 'delete',
        ...EDIT,
        from: at(doc, middle, 0),
        to: at(doc, middle, 6),
      });
      apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, middle, 6), text: 'Article' });
      // Ben types inside Ana's new word: his text joins her card
      apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, middle, 9), text: 'X' }, 'Ben');
      expectSameCards(doc, `${paragraphs} paragraphs`);
      const tree = doc.stories.get(BODY)?.tree;
      if (!tree) throw new Error('no body');
      let count = 0;
      const prev = tree.prev.bind(tree);
      const next = tree.next.bind(tree);
      tree.prev = (item) => {
        count += 1;
        return prev(item);
      };
      tree.next = (item) => {
        count += 1;
        return next(item);
      };
      const card = orderedCardOf(doc, 'S:ana:2.1');
      return { count, card: card?.suggestions };
    };
    const small = steps(200);
    const large = steps(4000);
    expect(small.card).toEqual(['S:ana:1.1', 'S:ana:2.1', 'S:ben:1.1']);
    expect(large.card).toEqual(small.card);
    expect(large.count).toBe(small.count);
    // The deleted word, the new word with Ben's letter inside it, and their neighbours
    expect(small.count).toBeLessThan(80);
  });

  test("a change inside another person's new row or table joins their card, as the host's earlier check joins it", () => {
    // Word's new row marks the row and its cells' paragraph marks, not the cells: the cells
    // load with no inserter of their own, inside the row's insertion (`within`)
    const doc = load();
    const [first] = opens(doc, 'tr');
    if (!first) throw new Error('fixture');
    apply(
      doc,
      'ben:1',
      { kind: 'insert-row', ...EDIT, row: addressOf(first), where: 'below' },
      'Ben'
    );
    // Paragraphs 4 and 5 are A1 and B1; 6 and 7 are Ben's new row
    expect([4, 5, 6, 7, 8].map((k) => textOf(doc, k))).toEqual(['A1', 'B1', '', '', 'A2']);
    const cell = opens(doc, 'tc')[2];
    if (!cell || cell.by !== null) throw new Error("Ben's new cell, as Word marks it");
    apply(doc, 'olga:1', { kind: 'insert', ...EDIT, place: at(doc, 6, 0), text: 'x' }, 'Olga');
    const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } } as const;
    apply(
      doc,
      'olga:2',
      { kind: 'table-property', ...EDIT, on: [addressOf(cell)], op: fill },
      'Olga'
    );
    expect(orderedCardOf(doc, 'S:ben:1.1')?.suggestions).toEqual([
      'S:ben:1.1',
      'S:olga:1.1',
      'S:olga:2.1',
    ]);
    expectSameCards(doc, 'new row');
    // A new table is Word's form too
    const fresh = load();
    apply(
      fresh,
      'ben:1',
      {
        kind: 'insert-table',
        ...EDIT,
        place: at(fresh, 0, 0),
        rows: 2,
        columns: 2,
        columnWidth: 3000,
      },
      'Ben'
    );
    expect(textOf(fresh, 0)).toBe('');
    apply(fresh, 'olga:1', { kind: 'insert', ...EDIT, place: at(fresh, 0, 0), text: 'x' }, 'Olga');
    expect(orderedCardOf(fresh, 'S:olga:1.1')?.suggestions).toEqual(['S:ben:1.1', 'S:olga:1.1']);
    // A whole row between two of his rows is a change of its own: the host's earlier check reads a
    // row as its own paragraphs, not inside his
    const [top] = opens(fresh, 'tr');
    if (!top) throw new Error('fixture');
    apply(
      fresh,
      'olga:2',
      { kind: 'insert-row', ...EDIT, row: addressOf(top), where: 'below' },
      'Olga'
    );
    expect(orderedCardOf(fresh, 'S:olga:2.1')?.suggestions).toEqual(['S:olga:2.1']);
    expectSameCards(fresh, 'new table');
  });
});
