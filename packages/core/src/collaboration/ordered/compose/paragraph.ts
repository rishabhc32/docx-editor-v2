// One paragraph from its items: its view (the units it shows), and its node. An untouched
// paragraph whose items are all still there is the loaded node itself.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { isLabelled, movedAwayBy } from '../model/document.ts';
import { isFilledPrompt } from '../model/prompt-marks.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import { childWhere, isParagraphProperties, withChildren, withMarkFace } from '../ooxml-nodes.ts';
import { ComposeContext, type OrderedViewChar, type OrderedViewParagraph } from './context.ts';
import { composeInline } from './inline.ts';
import { keptLinkEdges } from './kept-links.ts';
import { isNotesStory } from './note-references.ts';
import { faceAfter, paragraphPropertiesAfter, withNumberingRecords } from './properties.ts';
import { withMarkSection } from './sections.ts';
import { textBoxesUntouched } from './text-boxes.ts';

/** A drawn paragraph. */
export interface DrawnParagraph {
  readonly node: OoxmlNode;
  readonly view: OrderedViewParagraph;
  /** Something in it is labelled, or comes from another paragraph. */
  readonly touched: boolean;
  /** A unit sits in the paragraph, shown or not. */
  readonly hadUnits: boolean;
  /** A unit shows in it. */
  readonly drawn: boolean;
  /** A decision on a suggestion took a unit in it away (`ComposeContext.decidedAway`). */
  readonly decided: boolean;
  /** A decision changed whether a unit in it reads as text (`ComposeContext.decided`). */
  readonly reached: boolean;
  /** A unit in it shows struck by a pending suggestion. */
  readonly struck: boolean;
  /** A unit in it shows, neither struck nor a suggestion's insertion. */
  readonly own: boolean;
  /** A unit in it shows as a suggestion's insertion an accept decided. */
  readonly accepted: boolean;
  /** The words of insertions a reject took away in it. */
  readonly rejected: string;
}

/** Draw one paragraph: the items since the last shown mark, and its mark. */
export function composeParagraph(
  ctx: ComposeContext,
  buffer: readonly OrderedItem[],
  mark: OrderedItem,
  blocks: readonly OrderedItem[]
): DrawnParagraph {
  const { view, shell, touched, ...seen } = viewOfParagraph(ctx, buffer, mark, blocks);
  const home = mark.home;
  const original = home === undefined ? undefined : ctx.doc.originals.get(home);
  if (
    !touched &&
    !ctx.options.rebuild &&
    original &&
    isWhole(buffer, home, mark) &&
    textBoxesUntouched(ctx, buffer)
  )
    return { node: original, view, touched: false, ...seen };
  ctx.home = shell.id;
  const pPr = markProperties(ctx, mark, shell);
  ctx.paragraphKey = mark.space;
  const node = withChildren(shell, [...(pPr ? [pPr] : []), ...composeInline(ctx, buffer)]);
  return { node, view, touched: true, ...seen };
}

/** A paragraph's view, its mark's paragraph node, and what its drawing tells the blocks. */
interface ParagraphView extends Omit<DrawnParagraph, 'node'> {
  readonly shell: OoxmlElement;
}

/**
 * One paragraph's view (the units it shows, in order) and its mark's paragraph node, and whether
 * it is touched, held units, drew units, or lost a unit to a decision. It hides the prompts of
 * controls that hold content (in `ctx`), as the drawing does.
 */
export function viewOfParagraph(
  ctx: ComposeContext,
  buffer: readonly OrderedItem[],
  mark: OrderedItem,
  blocks: readonly OrderedItem[]
): ParagraphView {
  const shell = mark.node;
  if (!shell || shell.kind === 'textValue')
    throw new RangeError('A paragraph mark without its paragraph.');
  const home = mark.home;
  hidePrompts(ctx, buffer);
  const chars: OrderedViewChar[] = [];
  const containers: string[] = [];
  let touched = isLabelled(mark);
  let hadUnits = false;
  let drawn = false;
  let decided = false;
  let reached = false;
  let struck = false;
  let own = false;
  let accepted = false;
  let rejected = '';
  const kept = keptLinkEdges(ctx, buffer);
  for (const item of buffer) {
    if (isLabelled(item) || item.home !== home) touched = true;
    if (item.kind === 'open') {
      if (!ctx.gone(item) || kept.has(item))
        containers.push(`${item.token === 'hyperlink' ? 'h' : 'c'}:${item.space}`);
      continue;
    }

    if (item.kind === 'close') {
      if (!ctx.gone(item) || kept.has(item)) containers.pop();
      continue;
    }

    if (!isUnit(item)) continue;
    hadUnits = true;
    if (ctx.decidedAway(item)) decided = true;
    if (ctx.decided(item)) {
      reached = true;
    }

    rejected += ctx.rejectedWords(item);
    if (ctx.gone(item)) continue;
    drawn = true;
    const deleted = ctx.struckBy(item) !== null;
    struck ||= deleted;
    own ||= ctx.ownWord(item);
    accepted ||= ctx.acceptedInsertion(item);
    const within = containers.join('>');
    for (let k = 0; k < item.length; k += 1) {
      chars.push({
        address: { space: item.space, offset: item.offset + k },
        ch: item.kind === 'text' ? item.text.charAt(k) : item.text,
        containers: within,
        face: item.face,
        deleted,
        item,
      });
    }
  }

  const view: OrderedViewParagraph = { id: shell.id, mark, chars, items: buffer, blocks };
  return {
    view,
    shell,
    touched,
    hadUnits,
    drawn,
    decided,
    reached,
    struck,
    own,
    accepted,
    rejected,
  };
}

/** Every unit of the loaded paragraph is here, in one piece. */
function isWhole(
  buffer: readonly OrderedItem[],
  home: string | undefined,
  mark: OrderedItem
): boolean {
  let units = 0;
  for (const item of buffer) {
    if (item.home !== home) return false;
    units += item.length;
  }

  return mark.unitCount === undefined || units === mark.unitCount;
}

/** The mark's `w:pPr`: its properties under their records, and its own revision markers. */
function markProperties(
  ctx: ComposeContext,
  mark: OrderedItem,
  shell: OoxmlElement
): OoxmlElement | null {
  const base = childWhere(shell, isParagraphProperties) ?? null;
  const after = withNumberingRecords(ctx, mark, paragraphPropertiesAfter(ctx, mark, base));
  const pPr = ctx.fresh(withMarkSection(ctx, mark, after));

  // A split's new mark takes the face of the mark it split: that one keeps the ids
  const face = ctx.fresh(faceAfter(ctx, mark.fmts, mark.face), ctx.homed(mark));
  const by = ctx.insertedBy(mark);
  const del = ctx.struckBy(mark);
  const moved = mark.move !== null;

  // A notes part's marks take ids apart from the body's (a struck note's body, D20)
  const notes = isNotesStory(ctx.story);

  // A mark in a cell of a new table takes its rows' id, as the engine writes one (Word's form)
  let insertRole: string = REVISION_ROLE.markInsert;
  if (notes) {
    insertRole = REVISION_ROLE.noteMarkInsert;
  } else if (by !== null && mark.within.by.includes(by)) {
    insertRole = REVISION_ROLE.rowInsert;
  }

  const markers: OoxmlNode[] = [];
  if (by)
    markers.push(
      build(ctx.mint(), 'generic', moved ? 'moveTo' : 'ins', ctx.revisionMark(by, insertRole), [])
    );
  if (del)
    markers.push(
      build(
        ctx.mint(),
        'generic',
        movedAwayBy(ctx.doc, mark, del) ? 'moveFrom' : 'del',
        ctx.revisionMark(del, notes ? REVISION_ROLE.noteMarkDelete : REVISION_ROLE.markDelete),
        []
      )
    );
  let rPr = face;
  if (markers.length > 0) {
    const runProperties = face ?? ctx.drawn(build(ctx.mint(), 'runProperties', 'rPr', [], []));
    rPr = withChildren(runProperties, [...markers, ...runProperties.children]);
  }

  if (!pPr && !rPr) return null;
  const written = withMarkFace(pPr ?? build(ctx.mint(), 'paragraphProperties', 'pPr', [], []), rPr);
  return base === null ? ctx.drawn(written) : written;
}

/**
 * Whether a suggestion put this unit in a control, whatever its status now: the engine kept the
 * control's prompt as a glossary block then, and the control names it for good. Not a break: an
 * Enter at a prompt's edge leaves the prompt as it is, and one inside it removes it (policy 4).
 * Not a direct edit taken back (`revert`): the engine's undo gives the control back as it was.
 * Not a unit a decision took away before the control showed a prompt (`change/settle-emptied.ts`
 * marks it gone): it filled no prompt.
 */
const filledBy = (ctx: ComposeContext, item: OrderedItem): boolean =>
  item.by !== null && !item.retracted && !item.gone && !ctx.revertedDirect(item.by) && isUnit(item);

/**
 * The items after a control's open item up to its close item, each with the open item of the
 * control inside it that holds it, or null: a control there that shows a prompt of its own, or
 * that a prompt holds, keeps its own items.
 */
function* promptWalk(
  items: Iterable<OrderedItem>,
  open: OrderedItem
): Generator<readonly [OrderedItem, OrderedItem | null]> {
  let nested: OrderedItem | null = null;
  for (const item of items) {
    if (item.kind === 'close' && item.pair === open.pair) return;
    if (
      nested === null &&
      item.kind === 'open' &&
      item.prompt === true &&
      item.pair !== undefined
    ) {
      nested = item;
    }

    yield [item, nested];
    if (nested !== null && item.kind === 'close' && item.pair === nested.pair) {
      nested = null;
    }
  }
}

/**
 * The items after a control's open item up to its close item, without those of a control inside
 * it that shows a prompt of its own: that prompt shows or hides by that control's own rule.
 */
export function* ownItems(items: Iterable<OrderedItem>, open: OrderedItem): Generator<OrderedItem> {
  for (const [item, nested] of promptWalk(items, open)) {
    if (nested === null) {
      yield item;
    }
  }
}

/**
 * A control's own items (`ownItems`), and the items of the controls inside it that go with its
 * prompt: content that replaces a prompt replaces what the prompt holds, as the engine's
 * `clearPlaceholder` empties the control. Only a control there that shows its own prompt and
 * that no edit reached goes: one with a suggestion in it, or one someone put there, stays, shown
 * or not by its own rule, so no suggestion goes with the prompt and no later decision takes the
 * control away; one with words of its own stays too.
 */
function promptParts(
  items: Iterable<OrderedItem>,
  open: OrderedItem
): { readonly own: OrderedItem[]; readonly going: OrderedItem[] } {
  const own: OrderedItem[] = [];
  const nested = new Map<OrderedItem, OrderedItem[]>();
  for (const [item, holder] of promptWalk(items, open)) {
    if (holder === null) {
      own.push(item);
    } else if (nested.has(holder)) {
      nested.get(holder)?.push(item);
    } else {
      nested.set(holder, [item]);
    }
  }

  const goes = (held: readonly OrderedItem[]) =>
    held.every((item) => !isLabelled(item) && (item.prompt === true || !isUnit(item)));
  return { own, going: [...nested.values()].filter(goes).flat() };
}

/**
 * A control's prompt shows only while the control holds no other content (struck content
 * counts): typed text replaces it, and it comes back when that text goes.
 */
function hidePrompts(ctx: ComposeContext, buffer: readonly OrderedItem[]): void {
  for (let i = 0; i < buffer.length; i += 1) {
    const open = buffer[i];
    if (open?.kind === 'open' && open.prompt) hideInlinePrompt(ctx, open, buffer.slice(i + 1));
  }
}

/** One inline control's prompt, by the rule of `hidePrompts`; `after` runs past its close. */
function hideInlinePrompt(ctx: ComposeContext, open: OrderedItem, after: Iterable<OrderedItem>) {
  const { own, going } = promptParts(after, open);
  const units = own.filter((item) => !item.prompt && isUnit(item));
  if (units.some((item) => filledBy(ctx, item)) || isFilledPrompt(ctx.doc, open.space))
    ctx.filledPrompts.add(open.space);
  const present = units.some((item) => !ctx.unshown(item));
  if (!present) return;
  ctx.unprompted.add(open.space);
  for (const item of own) {
    if (item.prompt) {
      ctx.hiddenPrompts.add(item);
    }
  }

  for (const item of going) {
    ctx.hiddenPrompts.add(item);
  }
}

/** Whether every unit of a control's prompt is gone, and it has one: Enter removed it for good. */
function removedForGood(prompts: readonly OrderedItem[]): boolean {
  const units = prompts.filter(isUnit);
  return units.length > 0 && units.every((item) => item.gone);
}

/**
 * A block control's prompt hides when the control holds other content, in any paragraph, or
 * when Enter in it removed it (policy 4).
 */
export function hideBlockPrompt(ctx: ComposeContext, open: OrderedItem): void {
  const prompts: OrderedItem[] = [];
  let present = false;
  const { own, going } = promptParts(ctx.story.tree.items(ctx.story.tree.next(open)), open);
  for (const item of own) {
    if (item.prompt) {
      // A prompt's marks show or go in the model (`change/settle-prompts.ts`)
      if (item.kind !== 'mark') prompts.push(item);
      continue;
    }

    if (filledBy(ctx, item)) ctx.filledPrompts.add(open.space);
    if (isUnit(item) && !ctx.gone(item)) present = true;
  }

  if (isFilledPrompt(ctx.doc, open.space)) ctx.filledPrompts.add(open.space);

  // Enter in the prompt removed it for good (policy 4): the control shows no prompt
  const removed = removedForGood(prompts);
  if (removed) ctx.promptRemoved.add(open.space);
  if (!present && !removed) return;
  ctx.unprompted.add(open.space);
  for (const item of [...prompts, ...going]) {
    ctx.hiddenPrompts.add(item);
  }
}

/** Whether Enter removed a block control's prompt for good (policy 4): its every unit is gone. */
export function promptRemovedForGood(story: OrderedStory, open: OrderedItem): boolean {
  const inside = [...ownItems(story.tree.items(story.tree.next(open)), open)];
  return removedForGood(inside.filter((item) => item.prompt === true));
}

/** Whether a block control shows its prompt: `hideBlockPrompt` finds no reason to hide it. */
export function showsBlockPrompt(ctx: ComposeContext, open: OrderedItem): boolean {
  if (open.token !== 'sdtBlock' || open.prompt !== true) return false;
  hideBlockPrompt(ctx, open);
  return !ctx.unprompted.has(open.space);
}

/** Whether a control, inline or block, shows its prompt: its drawing finds no reason to hide it. */
export function showsPrompt(ctx: ComposeContext, open: OrderedItem): boolean {
  if (open.kind !== 'open' || open.prompt !== true) return false;
  if (open.token === 'sdtBlock') return showsBlockPrompt(ctx, open);
  hideInlinePrompt(ctx, open, ctx.story.tree.items(ctx.story.tree.next(open)));
  return !ctx.unprompted.has(open.space);
}

/** Whether a control shows its prompt in a drawing of its story alone (`showsPrompt`). */
export const promptShows = (
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem
): boolean => showsPrompt(new ComposeContext(doc, story, {}), open);

/** The control whose prompt an item is: the nearest open item around it that holds a prompt. */
export function promptControlOf(story: OrderedStory, item: OrderedItem): OrderedItem | null {
  if (item.prompt !== true) return null;
  let depth = 0;
  for (let at = story.tree.prev(item); at !== null; at = story.tree.prev(at)) {
    if (at.kind === 'close') depth += 1;
    else if (at.kind === 'open' && depth > 0) depth -= 1;
    else if (at.kind === 'open' && at.prompt === true) return at;
  }

  return null;
}

/**
 * The control whose prompt holds a control, in its paragraph or around it (past the marks of a
 * prompt of several paragraphs), or null. The walk ends at its paragraph's start.
 */
function promptHolderOf(story: OrderedStory, open: OrderedItem): OrderedItem | null {
  let depth = 0;
  for (let at = story.tree.prev(open); at !== null; at = story.tree.prev(at)) {
    if (at.kind === 'mark' && at.prompt !== true) {
      return null;
    }

    if (at.kind === 'close') {
      depth += 1;
    } else if (at.kind === 'open' && depth > 0) {
      depth -= 1;
    } else if (at.kind === 'open' && at.prompt === true) {
      return at;
    }
  }

  return null;
}

/**
 * The controls whose prompts hold a prompt's item, innermost first: its control, then each
 * control whose prompt holds the one before.
 */
export function* promptControlsAround(
  story: OrderedStory,
  item: OrderedItem
): Generator<OrderedItem> {
  let control = promptControlOf(story, item);
  for (; control !== null; control = promptHolderOf(story, control)) {
    yield control;
  }
}

/**
 * The control whose prompt content replaces at a prompt's item: of its control and the controls
 * whose prompts hold that one, the outermost that shows its prompt, as the engine replaces the
 * outermost (`placeholderControlForInsertion`); null when none shows one.
 */
export function replacedPromptOf(ctx: ComposeContext, item: OrderedItem): OrderedItem | null {
  let replaced: OrderedItem | null = null;
  for (const control of promptControlsAround(ctx.story, item)) {
    if (showsPrompt(ctx, control)) {
      replaced = control;
    }
  }

  return replaced;
}

/**
 * Whether a prompt's unit shows nowhere: content took the place of its control's prompt, or of
 * the prompt that holds its control, which goes with it (`promptParts`).
 */
export function promptUnitHidden(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): boolean {
  const ctx = new ComposeContext(doc, story, {});
  for (const control of promptControlsAround(story, item)) {
    showsPrompt(ctx, control);
    if (ctx.hiddenPrompts.has(item)) {
      return true;
    }
  }

  return false;
}
