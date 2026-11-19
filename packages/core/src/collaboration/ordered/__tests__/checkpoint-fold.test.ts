// A fold changes nothing the room uses (design D14): random sessions replayed from the room's log,
// folded at the end of rounds, give back every shown unit and its labels, every pending suggestion
// whole, its card, who typed its units, the takeovers, comments, part revisions, what a page
// shows, the Word file, and the Word file after an accept and after a reject of each suggestion.

import { describe, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { SESSION_FILE } from './sessions/fixture.ts';
import { runSession } from './sessions/session.ts';

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
