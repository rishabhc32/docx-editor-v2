// A format pressed twice by a party, with real editors and one room (audit R1 row 24). The second
// press gives the words back the face they had, so the editor alone writes no change record and
// shows no card, as Word removes a run's change record once the run is back to its original. The
// room makes no card, no record and no off value (`w:u w:val="none"`) either.

import { afterEach, describe, expect, test } from 'bun:test';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { closeAll, join, localRoom, ORGANISER, settle } from './editor-harness.ts';
import {
  alone,
  exec,
  PLAIN_FILE,
  revisionAuthors,
  roomBody,
  select,
  type Step,
} from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** A toggle press on the word "Bold", twice. */
const twice =
  (mark: 'bold' | 'underline' | 'italic'): Step =>
  (editor) => {
    for (let k = 0; k < 2; k += 1) {
      select(editor, 6, 10);
      exec(editor, { type: 'toggleMark', mark });
    }
  };

async function people() {
  const room = localRoom(PLAIN_FILE, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: PLAIN_FILE, review });
  const olive = await join(room, 'Olive', {
    wordFile: PLAIN_FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

describe('a format pressed twice by a party (audit R1 row 24)', () => {
  for (const mark of ['underline', 'bold', 'italic'] as const)
    test(`${mark} twice leaves no card and no record, as in the editor alone`, async () => {
      const { room, ana, olive } = await people();
      twice(mark)(ana.editor);
      await settle();
      converged(room, ana, olive);
      expect(orderedCards(room.room.document)).toEqual([]);
      expect(revisionAuthors(olive)).toEqual([]);
      const body = await roomBody(olive);
      expect(body).not.toContain('rPrChange');
      expect(body).toBe(
        await alone(PLAIN_FILE, review, [{ author: 'Ana', mode: 'suggesting', step: twice(mark) }])
      );
    });

  test("undo and redo of the second press give the editor alone's result", async () => {
    const { room, ana, olive } = await people();
    const undone: Step = (editor) => {
      twice('underline')(editor);
      exec(editor, { type: 'undo' });
    };
    const redone: Step = (editor) => {
      undone(editor);
      exec(editor, { type: 'redo' });
    };
    undone(ana.editor);
    await settle();
    converged(room, ana, olive);
    expect(revisionAuthors(olive)).toEqual(['Ana']);
    expect(await roomBody(olive)).toBe(
      await alone(PLAIN_FILE, review, [{ author: 'Ana', mode: 'suggesting', step: undone }])
    );
    exec(ana.editor, { type: 'redo' });
    await settle();
    converged(room, ana, olive);
    expect(orderedCards(room.room.document)).toEqual([]);
    expect(await roomBody(olive)).toBe(
      await alone(PLAIN_FILE, review, [{ author: 'Ana', mode: 'suggesting', step: redone }])
    );
  });
});
