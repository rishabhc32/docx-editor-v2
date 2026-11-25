// Columns a take-back or a reject removed, beside later records and columns, with real editors and
// one room (task 9.7). Each copy equals the editor alone, live and on the saved file reopened.

import { afterEach, expect, test } from 'bun:test';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  caretInSurface,
  commandOn,
  comparedBody,
  converged,
  TABLE_FILE,
  type Surface,
} from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** A step: who makes it, and the command or the organiser's reject of the first card. */
type Step = readonly [string, ((s: Surface) => void) | 'reject'];

const command =
  (text: string, given: Readonly<Record<string, unknown>>) =>
  (s: Surface): void => {
    caretInSurface(s, text);
    expect(commandOn(s, given).ok).toBe(true);
  };
const column = (text: string) => command(text, { type: 'insertColumn', where: 'right' });
const takeBack = command('', { type: 'deleteColumn' });
const align = command('Rate', { type: 'setTableCellVerticalAlignment', alignment: 'center' });

/** Delete the column of the cell before the one that holds `text`. */
const takeBackBefore =
  (text: string) =>
  (s: Surface): void => {
    const part = s.session.part();
    const ids = s.session.paragraphIds();
    const id = ids[ids.findIndex((candidate) => paragraphTextOf(part, candidate) === text) - 1];
    if (id === undefined) {
      throw new Error(`no cell before ${text}`);
    }

    s.setSelection({
      anchor: { paragraphId: id, offset: 0 },
      head: { paragraphId: id, offset: 0 },
    });
    expect(commandOn(s, { type: 'deleteColumn' }).ok).toBe(true);
  };

const rejectFirst = (editor: Person['editor']): void => {
  const [card] = editor.getReviewItems().filter((item) => item.kind === 'revision');
  expect(editor.rejectReviewItem(card?.key ?? '').ok).toBe(true);
};

/** The steps in one room, then on the saved file reopened: each body equals the editor alone's. */
async function checkRoom(steps: readonly Step[]): Promise<void> {
  const expected = alone({ author: 'Ana', mode: 'suggesting', review }, (s, editor) => {
    for (const [who, step] of steps) {
      if (step === 'reject') {
        rejectFirst(editor);
      } else {
        editor.setAuthor(who);
        step(s);
      }
    }
  });
  const seat = async (file: Uint8Array) => {
    const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
    const people = new Map<string, Person>();
    for (const name of new Set(steps.flatMap(([who, step]) => (step === 'reject' ? [] : [who])))) {
      people.set(name, await join(room, name, { wordFile: file, review }));
    }

    const olive = await join(room, 'Olive', {
      wordFile: file,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    return { room, people, olive };
  };
  const { room, people, olive } = await seat(TABLE_FILE);
  for (const [who, step] of steps) {
    if (step === 'reject') {
      rejectFirst(olive.editor);
    } else {
      step(surface(people.get(who)!));
    }

    await settle();
  }

  converged(room, olive, ...people.values());
  expect(comparedBody(surface(olive).session.part())).toBe(expected);
  const reopened = await seat(new Uint8Array(await olive.editor.save()));
  expect(comparedBody(surface(reopened.olive).session.part())).toBe(expected);
}

// A cell alignment sets no width: the cell a removed column took width from shows what that column
// left it, and the alignment's record holds that width as its original, made before the removal
// or after it
test.each([
  ['taken back', 'after'],
  ['taken back', 'before'],
  ['rejected', 'after'],
  ['rejected', 'before'],
])(
  "a party's column %s under their column made from its cell, their cell alignment %s",
  async (how, when) => {
    const removal: Step = how === 'taken back' ? ['Ana', takeBack] : ['Olive', 'reject'];
    const alignment: Step = ['Ana', align];
    const made: Step[] = [
      ['Ana', column('Fee')],
      ['Ana', column('')],
    ];
    await checkRoom([...made, ...(when === 'after' ? [removal, alignment] : [alignment, removal])]);
  }
);

// A party's column taken back beside another's made at the same cell, after a third party's column
// taken back under a later one of theirs: the cell before it takes its space, as the engine gives it
test("a party's column taken back beside another's at the same cell and a removed column", async () => {
  await checkRoom([
    ['Ben', column('Fee')],
    ['Cy', column('Fee')],
    ['Ana', column('Fee')],
    ['Ana', column('')],
    ['Ana', takeBack],
    ['Ben', takeBackBefore('Rate')],
  ]);
});
