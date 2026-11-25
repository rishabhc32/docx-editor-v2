// A decision the editor alone refuses is refused in the room too, by the engine's reason, so no
// copy decides what another copy cannot: the room asks the engine's own plan for the cards on
// the part they are in.
//
// Ben's column beside the organiser's direct deletion of the column it took width from, then
// Ana's table style with a header row and her deletion of the other column: the struck cells
// hold her style's cell records, and the engine plans no accept of her deletion and no reject of
// Ben's column (`incomplete-group`). The room once decided both, and a fold of its file gave
// other cell records than the live room (backlog row 103).

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  DIRECT,
  EDIT,
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
