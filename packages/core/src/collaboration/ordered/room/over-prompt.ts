// A party's edit that puts content in a control in place of its shown prompt, while another
// person's pending suggestion is on that prompt.
//
// Content in a control replaces its prompt: Word drops the prompt's runs, and the engine empties
// them (`clearPlaceholder`), with any revision on them. The save keeps the prompt's words only,
// so another person's pending suggestion on the prompt would not survive a save, and the room's
// rule finds it gone. The suggestions-only policy refuses such an edit as `nested`.

import { ComposeContext } from '../compose/context.ts';
import { ownItems, promptControlOf, showsPrompt } from '../compose/paragraph.ts';
import { isHidden, pendingByOther } from '../model/document.ts';
import type { OrderedMaker } from '../model/person.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import type { OrderedCaret, OrderedTouched } from '../change/targets.ts';
import type { OrderedEditChange } from '../change/types.ts';
import { itemsInside } from '../placement.ts';

/** Where a caret is against a shown prompt. */
type PromptPlace = 'edge' | 'inside' | 'block-edge' | 'block';

/** Kinds that fill the control they name, in place of its prompt. */
const FILLS: ReadonlySet<OrderedEditChange['kind']> = new Set(['control-value', 'building-block']);

/** Kinds whose content goes into a control at its prompt's edge: typed text (`promptSite`). */
const TYPES_AT_EDGE: ReadonlySet<OrderedEditChange['kind']> = new Set([
  'insert',
  'move',
  'insert-note',
]);

/**
 * The unit next to a caret in its paragraph, as a view reads it: past container edges, markers
 * with no width and hidden items; null at a paragraph mark.
 */
function unitBeside(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedItem | null,
  step: 'prev' | 'next'
): OrderedItem | null {
  for (let item = from; item !== null; item = story.tree[step](item)) {
    if (item.kind === 'mark') return null;
    if (isUnit(item) && !isHidden(doc, item)) return item;
  }

  return null;
}

/** The control whose shown prompt a caret is at, and where. */
function promptAt(
  ctx: ComposeContext,
  story: OrderedStory,
  caret: OrderedCaret
): { readonly control: OrderedItem; readonly place: PromptPlace } | null {
  // Inside an item, both are that item (`before` is `at` then)
  const before = unitBeside(ctx.doc, story, caret.before, 'prev');
  const after = unitBeside(ctx.doc, story, caret.at, 'next');
  const next = after ? promptControlOf(story, after) : null;
  const previous = before ? promptControlOf(story, before) : null;

  // Typed text goes into the control at the unit after the caret first (`promptSite`)
  const control = next ?? previous;
  if (!control || !showsPrompt(ctx, control)) {
    return null;
  }

  const both = next === previous;
  if (control.token === 'sdtBlock') return { control, place: both ? 'block' : 'block-edge' };
  return { control, place: both ? 'inside' : 'edge' };
}

/**
 * Whether the change puts content in the control at a caret there, as the edits place it (the
 * room policy test checks each kind at each place against the save):
 *  - a control value or a building block fills the control;
 *  - at a caret strictly inside an inline prompt, or in a block prompt's paragraph, every new
 *    item goes in the control; not a table (it goes beside the paragraph), and not Enter, which
 *    the engine refuses inside an inline prompt. Enter removes a block prompt anywhere in its
 *    paragraph, its edges too, as Word's Return replaces the selected prompt (policy 4);
 *  - at an inline prompt's edge only typed text goes in (placement rule 7, `promptSite`): words,
 *    a move's words, a note's reference, a paste of one paragraph. An object, a tab, a break or
 *    a table goes beside the control.
 */
function fillsAt(change: OrderedEditChange, place: PromptPlace): boolean {
  if (change.kind === 'insert-table') return false;
  if (change.kind === 'split') return place === 'block' || place === 'block-edge';
  if (place !== 'edge') return true;
  if (change.kind === 'insert-fragment')
    return change.blocks.length === 1 && change.blocks[0]?.e === 'w:p' && change.whole !== true;
  return TYPES_AT_EDGE.has(change.kind);
}

/** Whether a pending suggestion of someone other than `maker` is on a control's prompt. */
function othersOnPrompt(
  doc: OrderedDocument,
  maker: OrderedMaker,
  story: OrderedStory,
  control: OrderedItem
): boolean {
  const others = (id: string | null) => pendingByOther(doc, id, maker);
  for (const item of ownItems(itemsInside(story, control), control)) {
    // A unit gone for good shows nowhere, and the save drops it with its records
    if (item.prompt !== true || item.gone) continue;
    if (others(item.by) || item.dels.some(others)) return true;
    if (item.fmts.some((f) => others(f.suggestion))) return true;
    if (item.props.some((p) => others(p.suggestion))) return true;
  }

  return false;
}

/**
 * Whether a party's edit puts content in a control in place of its shown prompt while another
 * person's pending suggestion is on that prompt: the save would lose that suggestion.
 */
export function fillsOthersPrompt(
  doc: OrderedDocument,
  maker: OrderedMaker,
  change: OrderedEditChange,
  touched: OrderedTouched
): boolean {
  const story = touched.story;
  if (story === null) {
    return false;
  }

  const ctx = new ComposeContext(doc, story, {});
  if (FILLS.has(change.kind)) {
    for (const item of touched.items) {
      const showsOwnPrompt = item.kind === 'open' && item.prompt === true && showsPrompt(ctx, item);
      if (showsOwnPrompt && othersOnPrompt(doc, maker, story, item)) {
        return true;
      }
    }
  }

  for (const caret of touched.carets) {
    const found = promptAt(ctx, story, caret);
    if (found && fillsAt(change, found.place) && othersOnPrompt(doc, maker, story, found.control)) {
      return true;
    }
  }

  return false;
}
