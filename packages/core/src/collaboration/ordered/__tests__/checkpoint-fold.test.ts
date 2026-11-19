// A fold changes nothing the room uses (design D14): random sessions replayed from the room's log,
// folded at the end of rounds, give back every shown unit and its labels, every pending suggestion
// whole, its card, who typed its units, the takeovers, comments, part revisions, what a page
// shows, the Word file, and the Word file after an accept and after a reject of each suggestion.

import { describe, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { apply, at, DIRECT, EDIT, textOf } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { SESSION_FILE } from './sessions/fixture.ts';
import { runSession } from './sessions/session.ts';
import { bodyOf, docx, MATRIX_ITEMS } from './word-fixtures.ts';

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
