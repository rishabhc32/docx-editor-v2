// An editor in a room keeps only the document trees it can still use.
//
// In a room every remote change drops the editor store's undo history, so no undo can bring an
// older tree back. Three caches still held them: the review queue memo kept the last eight
// trees with their whole queues, and the drawing slot and the image lookup each kept the tree
// the editor opened with, which collects the first version of every paragraph edited since.

import { afterEach, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const PARAGRAPHS = 30;
const FILE = docx(
  Array.from({ length: PARAGRAPHS }, (_unused, index) => p(t(`Clause ${index} text.`))).join('')
);

/** Ana suggests a word; Olive accepts it. Both editors take each other's changes. */
async function round(olive: Person, ana: Person, index: number): Promise<void> {
  caretAt(ana, index % PARAGRAPHS, 0);
  surface(ana).insertPlainText('A');
  await settle();
  const card = olive.editor.getReviewItems().find((item) => item.kind === 'revision');
  if (!card) throw new Error('no suggestion to accept');
  expect(olive.editor.acceptReviewItem(card.key).ok).toBe(true);
  await settle();
}

test('room editors let go of the trees they opened with and the ones they left', async () => {
  const review = reviewTestModule();
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  const opened = [olive, ana].map((person) => new WeakRef(surface(person).session.part().root));
  const session = surface(olive).session;
  const trees: WeakRef<object>[] = [];
  session.subscribe(() => {
    trees.push(new WeakRef(session.part().root));
  });
  for (let index = 0; index < 10; index += 1) await round(olive, ana, index);
  await settle();
  Bun.gc(true);

  expect(opened.map((ref) => ref.deref() === undefined)).toEqual([true, true]);
  expect(trees.length).toBeGreaterThanOrEqual(20);
  // The current tree, and at most one more that a reader still has in hand.
  expect(trees.filter((ref) => ref.deref() !== undefined).length).toBeLessThanOrEqual(2);
  for (const person of [olive, ana]) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
  }
});
