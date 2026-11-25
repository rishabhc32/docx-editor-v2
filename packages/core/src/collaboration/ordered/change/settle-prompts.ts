// A block control's prompt of two or more paragraphs, settled after each change (task 9.6).

import { isBlockToken } from '../compose/blocks.ts';
import { ComposeContext } from '../compose/context.ts';
import { promptRemovedForGood, showsBlockPrompt } from '../compose/paragraph.ts';
import { insertionTakenBack, isHidden, storyOfItem } from '../model/document.ts';
import { relabel } from '../model/mutations.ts';
import { recordPromptMarkChange, promptControlsOf, promptMarksOf } from '../model/prompt-marks.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { itemsInside } from '../placement.ts';
import { replacedPromptUnit } from './settle-named-prompts.ts';

/**
 * A prompt mark's control: items never move, so the answer holds while the mark and the control's
 * open item are in a story. A settle can put another open item in its place (`settle-emptied.ts`).
 */
const controlOfMark = new WeakMap<OrderedItem, OrderedItem | null>();

/**
 * Show or hide the marks of every block prompt of two or more paragraphs, by its control's
 * state. The writes go in the change's undo record.
 */
export function settlePromptMarks(doc: OrderedDocument): void {
  const marks = promptMarksOf(doc);
  const done = new Set<OrderedItem>();
  for (const mark of marks) {
    const story = storyOfItem(doc, mark);
    if (!story) continue;
    const open = controlOf(story, mark);
    if (!open || done.has(open)) continue;
    done.add(open);
    settleControl(doc, story, open);
  }

  // A one-paragraph prompt has no prompt mark: only its taken-back Enter needs a settle
  for (const open of promptControlsOf(doc)) {
    if (done.has(open)) continue;
    const story = storyOfItem(doc, open);
    if (!story || isHidden(doc, open) || !promptRemovedForGood(story, open)) continue;
    if (takenBack(doc, story, open)) restorePrompt(doc, story, open);
  }
}

/** The block control a prompt mark sits in right inside: the first open block item back. */
function controlOf(story: OrderedStory, mark: OrderedItem): OrderedItem | null {
  const known = controlOfMark.get(mark);
  if (known === null || (known !== undefined && known.leaf?.tree === story.tree)) {
    return known;
  }

  let found: OrderedItem | null = null;
  let depth = 0;
  for (let item = story.tree.prev(mark); item; item = story.tree.prev(item)) {
    if (!isBlockToken(item)) continue;
    if (item.kind === 'close') depth += 1;
    else if (depth > 0) depth -= 1;
    else {
      found = item.token === 'sdtBlock' && item.prompt === true ? item : null;
      break;
    }
  }

  controlOfMark.set(mark, found);
  return found;
}

/**
 * Set a block prompt's marks again. Typing into a prompt replaces it whole: the engine empties
 * the control into the caret's paragraph and drops the prompt's other paragraphs
 * (`clearPlaceholder`), as Word does, and the whole prompt shows again, from the glossary, once
 * the control is empty. The room keeps the prompt's units and marks and hides them while content
 * fills the control. Units hide where they are drawn (`hideBlockPrompt`); a mark must hide in the
 * model, because the paragraphs a view counts are the marks the model shows. So while the prompt
 * is hidden, a mark whose paragraph shows nothing but the prompt goes (`gone`), and compose
 * drops its paragraph whole, with no join (`composeBlocks`, `paragraphAt`): a join at the
 * control's last mark would reach past the control's end. Otherwise it shows. The paragraph a
 * person typed in keeps its mark, as the engine keeps the caret's paragraph.
 */
function settleControl(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): void {
  // A control that went (removed, or its insertion rejected) keeps its marks as they are
  if (isHidden(doc, open)) return;

  // So does one whose prompt Enter removed for good (`promptBreakSite` set them), unless the
  // Enter and all that followed it in the control were taken back: then it never happened, and
  // the prompt comes back whole, as the editor's undo brings it back
  if (promptRemovedForGood(story, open)) {
    if (takenBack(doc, story, open)) restorePrompt(doc, story, open);
    return;
  }

  const ctx = new ComposeContext(doc, story, {});
  const hidden = !showsBlockPrompt(ctx, open);

  // Each prompt mark ends a paragraph of the prompt, with what was typed or put in it
  const paragraphs: { readonly mark: OrderedItem; readonly holds: boolean }[] = [];
  let holds = false;
  let depth = 0;
  for (const item of itemsInside(story, open, doc)) {
    if (isBlockToken(item)) {
      depth += item.kind === 'open' ? 1 : -1;
      if (!isHidden(doc, item)) holds = true;
      continue;
    }

    if (depth === 0 && item.kind === 'mark' && item.prompt === true) {
      paragraphs.push({ mark: item, holds });
      holds = false;
      continue;
    }

    // The prompt's own units hide with the prompt, and so does a control in it that goes with it
    // (`hideBlockPrompt`); anything else that shows is content
    if ((item.prompt === true && item.kind !== 'mark') || ctx.hiddenPrompts.has(item)) {
      continue;
    }

    if (!isHidden(doc, item)) holds = true;
  }

  // With no paragraph to keep (an Enter that removed the prompt, then taken back), all stay
  const replaced = hidden && paragraphs.some((paragraph) => paragraph.holds);
  for (const { mark, holds: kept } of paragraphs) {
    const gone = replaced && !kept;
    if (mark.gone === gone) continue;
    relabel(doc, story, mark, { gone });
    recordPromptMarkChange(doc, mark);
  }
}

/** Whether everything in the control but its prompt was taken back: retracted, or reverted. */
function takenBack(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): boolean {
  for (const item of itemsInside(story, open, doc)) {
    // The prompt, and the file's own marks of its paragraphs (a one-paragraph prompt's mark is
    // not flagged); and what went for good, which is no content (a decision's words)
    if (item.prompt === true || item.gone || (item.kind === 'mark' && item.by === null)) continue;
    if (!insertionTakenBack(doc, item)) {
      return false;
    }
  }

  return true;
}

/** The control's prompt shown again: its units and marks back. */
function restorePrompt(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): void {
  for (const item of itemsInside(story, open, doc)) {
    // A unit the glossary's prompt replaced stays gone (`settle-named-prompts.ts`)
    if (item.prompt !== true || !item.gone || replacedPromptUnit(doc, item)) continue;
    relabel(doc, story, item, { gone: false });
    recordPromptMarkChange(doc, item);
  }
}
