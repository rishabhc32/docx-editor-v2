// Undo of the deletion of one's own new column, with real editors and one ordered room (task
// 9.7).
//
// Deleting one's own pending column withdraws its cells, as the engine does, so the file holds
// nothing of it. In the editor alone, and in Word, an undo of that deletion gives the column
// back, and a second undo takes the column away. The room's undo took back the column's
// suggestion, which the deletion had already taken back, so it changed nothing. Each step here is
// compared with the editor alone, for a party, the organiser suggesting, and the organiser's
// direct column. The same holds for one's own new words deleted whole; a deletion of a part of
// them is in `editor-undo-own-words.test.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  FILE,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  caretInSurface,
  comparedBody,
  converged,
  commandOn,
  type Surface,
  TABLE_FILE,
} from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

type Step = 'insert' | 'delete' | 'undo' | 'redo';

/** Put the caret in the cell right of the one that holds `Rate`: the new column's. */
function caretInNewCell(s: Surface): void {
  const part = s.session.part();
  const ids = s.session.paragraphIds();
  const at = ids.findIndex((id) => paragraphTextOf(part, id) === 'Rate');
  const id = ids[at + 1];
  if (at < 0 || id === undefined) throw new Error('no new cell');
  s.setSelection({ anchor: { paragraphId: id, offset: 0 }, head: { paragraphId: id, offset: 0 } });
}

function step(s: Surface, what: Step): void {
  if (what === 'insert') {
    caretInSurface(s, 'Rate');
    expect(commandOn(s, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
  } else if (what === 'delete') {
    caretInNewCell(s);
    expect(commandOn(s, { type: 'deleteColumn' }).ok).toBe(true);
  } else if (what === 'undo') s.undo();
  else s.redo();
}

/** The body after each step on the editor alone. */
function aloneSteps(author: string, mode: 'suggesting' | 'edit', steps: readonly Step[]) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: TABLE_FILE,
    author,
    mode,
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    return steps.map((what) => {
      step(s, what);
      return comparedBody(s.session.part());
    });
  } finally {
    editor.destroy();
    container.remove();
  }
}

const STEPS: readonly Step[] = ['insert', 'delete', 'undo', 'undo', 'redo', 'redo', 'undo', 'undo'];

const WHO = [
  ['a party', 'Ana', 'suggesting'],
  ['the organiser suggesting', 'Olive', 'suggesting'],
  ["the organiser's direct column", 'Olive', 'edit'],
] as const;

describe('undo of the deletion of your own new column, as the editor alone', () => {
  for (const [who, author, mode] of WHO)
    test(`${who}: ${STEPS.join(', ')}`, async () => {
      const expected = aloneSteps(author, mode, STEPS);
      const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
      const olive = await join(room, 'Olive', {
        wordFile: TABLE_FILE,
        mode: 'edit',
        login: ORGANISER,
        review,
      });
      const typist = author === 'Ana' ? ana : olive;
      typist.editor.setMode(mode);
      const shown: string[] = [];
      for (const what of STEPS) {
        step(surface(typist), what);
        await settle();
        converged(room, ana, olive);
        expect({ what, refusals: typist.refusals }).toEqual({ what, refusals: [] });
        shown.push(comparedBody(surface(ana).session.part()));
      }
      expect(shown).toEqual(expected);
      // Once the column is taken back again, no card stays
      expect(orderedCards(room.room.document)).toEqual([]);
    });
});

describe('undo of the deletion of all your own new words, as the editor alone', () => {
  const STEPS_TEXT = ['type', 'delete', 'undo', 'undo', 'redo', 'redo', 'undo', 'undo'] as const;
  const run = (person: Person, what: (typeof STEPS_TEXT)[number]) => {
    const s = surface(person);
    if (what === 'type') {
      caretAt(person, 0, 0);
      s.insertPlainText('abc');
    } else if (what === 'delete') {
      caretAt(person, 0, 0, 3);
      s.deleteBackward();
    } else if (what === 'undo') s.undo();
    else s.redo();
  };
  test(`a party: ${STEPS_TEXT.join(', ')}`, async () => {
    const container = document.createElement('div');
    document.body.append(container);
    const editor = createDocxEditor({
      container,
      document: FILE,
      author: 'Ana',
      mode: 'suggesting',
      modules: [review],
    });
    const me = { name: 'alone', editor, container } as unknown as Person;
    const expected = STEPS_TEXT.map((what) => {
      run(me, what);
      return editorTexts(me)[0];
    });
    editor.destroy();
    container.remove();
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const shown: (string | undefined)[] = [];
    for (const what of STEPS_TEXT) {
      run(ana, what);
      await settle();
      converged(room, ana, olive);
      shown.push(editorTexts(olive)[0]);
    }
    expect(shown).toEqual(expected);
    expect(orderedCards(room.room.document)).toEqual([]);
  });
});
