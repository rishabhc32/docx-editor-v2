// A decision the editor alone refuses is refused in the room too, by the engine's reason, so no
// copy decides what another copy cannot: the room asks the engine's own plan for the cards on
// the part they are in.
//
// Ben's column beside the organiser's direct deletion of the column it took width from, then
// Ana's table style with a header row and her deletion of the other column: the struck cells
// hold her style's cell records, and the engine plans no accept of her deletion and no reject of
// Ben's column (`incomplete-group`). The room once decided both, and a fold of its file gave
// other cell records than the live room (backlog row 103).
//
// A card of several suggestions is decided together, as the engine plans a batch of them: Ben's
// paragraph record in a cell of Ana's new column joins her card, and its reject takes both.
//
// The plan takes every item of a card, but only an item the editor decides with its dependencies
// can hold the card: Ana's deletion over a paragraph mark, with her new word in it, is decided as
// the editor alone decides each, though the batch would not drop her paragraph record there.
//
// A format record whose runs reach into another person's suggested deletion is a card of its own,
// and each of the two is decided alone, as Word decides them (r5-s10 of the pre-send check).

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  at,
  APPLY_FILE,
  DIRECT,
  EDIT,
  markOf,
  opens,
  refusal,
} from './change-apply-fixture.ts';

/** The table of `APPLY_FILE` after the four changes. */
function struck(): OrderedDocument {
  const { document: doc } = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const [a1, , , b2] = opens(doc, 'tc').map(addressOf);
  const table = [addressOf(opens(doc, 'tbl')[0]!)];
  const style = { op: 'setTableProperties', styleId: 'TableGrid', headerRowCount: 1 } as const;
  apply(doc, 'ben:1', { kind: 'insert-column', ...EDIT, cell: b2!, where: 'right' }, 'Ben');
  apply(doc, 'olive:1', { kind: 'delete-column', ...DIRECT, cell: b2! }, 'Olive');
  apply(doc, 'ana:1', { kind: 'table-property', ...EDIT, on: table, op: style }, 'Ana');
  apply(doc, 'ana:2', { kind: 'delete-column', ...EDIT, cell: a1! }, 'Ana');
  return doc;
}

const decide = (doc: OrderedDocument, suggestion: string, status: 'accepted' | 'rejected') =>
  ({
    kind: 'decide',
    cards: [{ suggestions: orderedCardSeen(doc, suggestion)! }],
    status,
  }) as const;

describe('a decision the engine plans no change for', () => {
  for (const [label, suggestion, status] of [
    ["an accept of Ana's column deletion", 'S:ana:2.1', 'accepted'],
    ["a reject of Ben's column", 'S:ben:1.1', 'rejected'],
  ] as const) {
    test(`${label} is refused as the editor alone refuses it`, () => {
      const doc = struck();
      expect(refusal(doc, 'olive:2', decide(doc, suggestion, status), 'Olive')).toBe(
        'engine-refused'
      );
    });
  }

  for (const [label, suggestion, status] of [
    ["a reject of Ana's column deletion", 'S:ana:2.1', 'rejected'],
    ["an accept of Ben's column", 'S:ben:1.1', 'accepted'],
  ] as const) {
    test(`${label} is decided, as the engine plans it`, () => {
      const doc = struck();
      apply(doc, 'olive:2', decide(doc, suggestion, status), 'Olive');
      expect(doc.suggestions.get(suggestion)?.status).toBe(status);
    });
  }
});

describe('a card of several suggestions', () => {
  // Seed sp5-run-1/quick-short-1 of the pre-send check (mismatch 5)
  for (const status of ['accepted', 'rejected'] as const) {
    test(`a paragraph record in a cell of another person's new column, ${status} with it`, () => {
      const { document: doc } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const right = addressOf(opens(doc, 'tc')[1]!);
      apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, cell: right, where: 'left' });
      const align = [{ localName: 'jc', attributes: { val: 'right' } }];
      const record = { kind: 'paragraph-property', ...EDIT, marks: [markOf(doc, 5)] } as const;
      apply(doc, 'bob:1', { ...record, properties: align }, 'Bob');
      expect(orderedCardSeen(doc, 'S:bob:1.1')?.map((s) => s.suggestion)).toEqual([
        'S:ana:1.1',
        'S:bob:1.1',
      ]);
      apply(doc, 'olive:1', decide(doc, 'S:bob:1.1', status), 'Olive');
      expect([...doc.suggestions.values()].map((s) => s.status)).toEqual([status, status]);
    });
  }

  for (const status of ['accepted', 'rejected'] as const) {
    test(`a deletion over a paragraph mark that holds the author's paragraph record, ${status}`, () => {
      const { document: doc } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const align = [{ localName: 'jc', attributes: { val: 'right' } }];
      const marks = [markOf(doc, 0), markOf(doc, 1)];
      apply(doc, 'ana:1', { kind: 'paragraph-property', ...EDIT, marks, properties: align });
      apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 0, 3), to: at(doc, 1, 2) });
      apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, 0, 3), text: 'New' });
      expect(orderedCardSeen(doc, 'S:ana:3.1')?.map((s) => s.suggestion)).toEqual([
        'S:ana:2.1',
        'S:ana:3.1',
      ]);
      apply(doc, 'olive:1', decide(doc, 'S:ana:3.1', status), 'Olive');
      expect(doc.suggestions.get('S:ana:2.1')?.status).toBe(status);
      expect(doc.suggestions.get('S:ana:1.1')?.status).toBe('pending');
    });
  }
});

describe("a format record that reaches into another person's deletion", () => {
  for (const status of ['accepted', 'rejected'] as const) {
    for (const suggestion of ['S:ana:1.1', 'S:ben:1.1']) {
      test(`${suggestion} is ${status} on its own`, () => {
        const { document: doc } = openOrderedWordFile(APPLY_FILE, { file: 1 });
        const bold = [{ localName: 'b' }];
        apply(doc, 'ana:1', {
          kind: 'format',
          ...EDIT,
          from: at(doc, 0, 0),
          to: at(doc, 0, 8),
          properties: bold,
        });
        apply(
          doc,
          'ben:1',
          { kind: 'delete', ...EDIT, from: at(doc, 0, 4), to: at(doc, 0, 12) },
          'Ben'
        );
        expect(orderedCardSeen(doc, suggestion)?.map((s) => s.suggestion)).toEqual([suggestion]);
        apply(doc, 'olive:1', decide(doc, suggestion, status), 'Olive');
        expect(doc.suggestions.get(suggestion)?.status).toBe(status);
      });
    }
  }
});
