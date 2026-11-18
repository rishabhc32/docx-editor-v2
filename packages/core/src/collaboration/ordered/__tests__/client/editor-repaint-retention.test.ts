// In a room, a command that repaints every page several times keeps none of the old pages queued.
//
// A subtree `MutationObserver` on the pages layer queued one record for each page a paint replaced,
// and each record held the removed page until the task ended. Each room editor has its own pages
// layer, so the check runs on both editors while the other one's edits arrive.

import { afterEach, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

const PAGES = 4;
const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
const FILE = docx(
  Array.from({ length: PAGES }, (_, page) => p(t(`Page ${page + 1} text.`)) + p(t('More.'))).join(
    pageBreak
  )
);

const view = document.defaultView as unknown as { MutationObserver: typeof MutationObserver };
const original = view.MutationObserver;
const observers: MutationObserver[] = [];

afterEach(() => {
  view.MutationObserver = original;
  observers.length = 0;
  closeAll();
});

/** Nodes (each removed node and its subtree) that the observers' queued records still hold. */
function queuedRemovedNodes(): number {
  let count = 0;
  const visit = (node: Node) => {
    count += 1;
    for (const child of node.childNodes) visit(child);
  };
  for (const observer of observers)
    for (const record of observer.takeRecords())
      for (const node of record.removedNodes) visit(node);
  return count;
}

/** Repaints every page of the person's editor `times` times, in the caller's task. */
function repaintEveryPage(person: Person, times: number): void {
  for (let k = 0; k < times; k += 1) surface(person).setShowParagraphMarks(k % 2 === 0);
}

test('room editors repaint every page several times in one task and queue no old page', async () => {
  view.MutationObserver = class extends original {
    constructor(callback: MutationCallback) {
      super(callback);
      observers.push(this);
    }
  };
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  for (const person of [olive, ana])
    expect(person.container.querySelectorAll('.docx-page')).toHaveLength(PAGES);
  await settle();

  // One task: Ana suggests a word, then both editors repaint every page several times.
  caretAt(ana, 0, 0);
  surface(ana).type('New ');
  repaintEveryPage(ana, 4);
  repaintEveryPage(olive, 4);
  expect(queuedRemovedNodes()).toBe(0);

  // The suggestion reaches Olive and her editor repaints; the room still converges.
  await settle();
  repaintEveryPage(olive, 2);
  expect(queuedRemovedNodes()).toBe(0);
  await settle();
  expect(editorTexts(olive)[0]).toBe('New Page 1 text.');
  expect(editorTexts(ana)[0]).toBe('New Page 1 text.');
  expect(sameAsRoom(olive, room)).toBe(true);
  expect(sameAsRoom(ana, room)).toBe(true);
});
