// A long session in a room keeps a bounded number of old document trees alive.
//
// Undo in a room belongs to the room's history, so the editor store's own undo history is never
// read there, but the decide commands still wrote entries to it. Each entry kept the package from
// before its commit, and the room's packages came in between, so each kept its own copy of every
// paragraph that changed since the last one, with every cache keyed by those nodes. The address
// hints of the remote carets kept the tree each caret last resolved in, too. A real editor grew
// from 0.5 GB to 1.5 GB in one session.

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

/**
 * One round: Ana suggests a word and moves her caret on, so Olive's editor resolves it; Olive
 * accepts the suggestion, a decide command of her editor.
 */
async function round(olive: Person, ana: Person, index: number): Promise<void> {
  caretAt(ana, index % PARAGRAPHS, 0);
  surface(ana).insertPlainText('A');
  await settle();
  caretAt(ana, (index + 15) % PARAGRAPHS, 1);
  await settle();
  const card = olive.editor.getReviewItems().find((item) => item.kind === 'revision');
  if (!card) throw new Error('no suggestion to accept');
  expect(olive.editor.acceptReviewItem(card.key).ok).toBe(true);
  await settle();
}

test('a long session in a room keeps a bounded number of old trees', async () => {
  const review = reviewTestModule();
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  const session = surface(olive).session;
  const trees: WeakRef<object>[] = [];
  session.subscribe(() => {
    trees.push(new WeakRef(session.part().root));
  });
  let rounds = 0;
  const liveAfter = async (count: number): Promise<number> => {
    for (let k = 0; k < count; k += 1) await round(olive, ana, rounds++);
    await settle();
    Bun.gc(true);
    return trees.filter((ref) => ref.deref() !== undefined).length;
  };
  const early = await liveAfter(4);
  const builtEarly = trees.length;
  const late = await liveAfter(8);
  // Each round makes new trees; the live ones stay as many as the current revision needs
  expect(trees.length - builtEarly).toBeGreaterThanOrEqual(16);
  expect(late - early).toBeLessThanOrEqual(2);
  for (const person of [olive, ana]) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
  }
});
