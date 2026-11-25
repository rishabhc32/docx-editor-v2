// Words a decision left with no paragraph, settled into the next paragraph after it.

import { isBlockToken } from '../compose/blocks.ts';
import { isHidden, storyOfItem } from '../model/document.ts';
import { moveItems, relabel } from '../model/mutations.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';

/**
 * Move each run of items that `decided` left before a block with no paragraph mark of its own
 * into the next paragraph that shows, with that paragraph's containers. A decision that takes a
 * paragraph mark away joins its words to the next paragraph; when a table or a block control
 * starts there, the next paragraph is inside it (its first cell), as compose draws them, the
 * editor alone's reject puts them, and a fold loads them. So a later decision on that table or
 * control takes the words with it, as in the editor alone, live as after a fold.
 */
export function settleLooseWords(doc: OrderedDocument, decided: readonly string[]): void {
  const walked = new Set<OrderedItem>();
  for (const id of decided) {
    for (const mark of doc.labelled.get(id) ?? []) {
      if (mark.kind !== 'mark' || walked.has(mark) || !isHidden(doc, mark)) {
        continue;
      }

      const story = storyOfItem(doc, mark);
      const run = story ? looseRun(doc, story, mark, walked) : null;
      if (story && run) {
        intoNextParagraph(doc, story, run);
      }
    }
  }
}

/** Whether an item ends a paragraph or a block: a mark that shows, or a block's edge. */
const endsRun = (doc: OrderedDocument, item: OrderedItem): boolean =>
  isBlockToken(item) || (item.kind === 'mark' && !isHidden(doc, item));

/**
 * The run of items around a mark that does not show, from the last paragraph's end up to the
 * block that starts or ends right after it, when something in the run shows. Null when a mark
 * that shows ends the run, as a paragraph, or when an earlier walk read the run (`walked`), so
 * a decision of many paragraph marks reads each item once.
 */
function looseRun(
  doc: OrderedDocument,
  story: OrderedStory,
  mark: OrderedItem,
  walked: Set<OrderedItem>
) {
  walked.add(mark);
  let end = story.tree.next(mark);
  while (end && !endsRun(doc, end)) {
    if (walked.has(end)) {
      return null;
    }

    walked.add(end);
    end = story.tree.next(end);
  }

  if (!end || !isBlockToken(end)) {
    return null;
  }

  const run: OrderedItem[] = [];
  for (let item = story.tree.prev(end); item && !endsRun(doc, item); item = story.tree.prev(item)) {
    run.push(item);
  }

  return run.some((item) => !isHidden(doc, item)) ? run.reverse() : null;
}

/**
 * Put a run at the start of the next paragraph that shows, right after the last block edge
 * before its mark, with that mark's containers. A run with no paragraph after it stays, which
 * compose refuses as before.
 */
function intoNextParagraph(doc: OrderedDocument, story: OrderedStory, run: OrderedItem[]): void {
  // The run ends at a block edge
  let edge = story.tree.next(run[run.length - 1]!)!;
  let target = story.tree.next(edge);
  while (target && (target.kind !== 'mark' || isHidden(doc, target))) {
    if (isBlockToken(target)) {
      edge = target;
    }

    target = story.tree.next(target);
  }

  if (!target) {
    return;
  }

  moveItems(doc, story, run, edge);
  const { within } = target;
  for (const item of run) {
    relabel(doc, story, item, { within });
  }
}
