// A control that names a glossary prompt shows that prompt once its content went, settled after
// each change (task 9.6). The writes go in the change's undo record.

import { replaceNode } from '../../../store/package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { withRestoredPrompt } from '../../../store/store/content-control-prompt-restore.ts';
import { findContentControl } from '../../../store/store/tree-op-nodes.ts';
import { isBlockToken } from '../compose/blocks.ts';
import { ComposeContext } from '../compose/context.ts';
import { ownItems } from '../compose/paragraph.ts';
import { placeholderNameOf } from '../compose/prompts.ts';
import { insertionTakenBack, isHidden, isSplitEdge, storyOfItem } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import {
  recordFilledPrompt,
  recordPromptMarkChange,
  promptedControlsOf,
} from '../model/prompt-marks.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { isControlContent, isWml, withChildren } from '../ooxml-nodes.ts';
import { closeOf } from '../placement.ts';
import type { OrderedApplyContext } from '../room/room.ts';
import {
  controlScratch,
  pending,
  promptUnits,
  settleContext,
  shownMarkOf,
} from './settle-emptied.ts';

/** A control's prompt as it is now: its units that show, the units that went, and its mark. */
interface Scan {
  /** The prompt's units that are not gone. */
  readonly prompt: readonly OrderedItem[];
  /** Units that filled the prompt and went by a decision or a deletion, not taken back. */
  readonly went: readonly OrderedItem[];
  /** Whether content shows in the control, so its prompt does not. */
  readonly present: boolean;
  /** The paragraph mark that shows the control; units go before it in a block control. */
  readonly mark: OrderedItem;
}

/** A replacement a deletion that is not a decision made: given back if that deletion is. */
interface Swap {
  readonly replaced: readonly OrderedItem[];
  readonly units: readonly OrderedItem[];
  readonly went: readonly OrderedItem[];
}

interface State {
  /** Replacements that can be given back, by the control's open item. */
  readonly swaps: Map<OrderedItem, Swap>;
  /** Units whose going was settled: only a unit that newly went asks for the glossary. */
  readonly settled: WeakSet<OrderedItem>;
  /** Prompt units a replacement took away: the prompt's own rules never bring them back. */
  readonly replaced: WeakSet<OrderedItem>;
}

const states = new WeakMap<OrderedDocument, State>();

function stateOf(doc: OrderedDocument): State {
  let found = states.get(doc);
  if (!found) {
    found = { swaps: new Map(), settled: new WeakSet(), replaced: new WeakSet() };
    states.set(doc, found);
  }

  return found;
}

/**
 * Whether a replacement took this prompt unit away for good, or until it is given back:
 * `change/settle-prompts.ts` must not show it again when an Enter is taken back.
 */
export function replacedPromptUnit(doc: OrderedDocument, item: OrderedItem): boolean {
  return states.get(doc)?.replaced.has(item) === true;
}

/**
 * Show the glossary's prompt in each control that names one, once its content went.
 *
 * Content that takes the place of a shown prompt drops the prompt's runs, with any revision on
 * them, as Word does (Word 16.114, `~$p4-enter-tracked.docx`: the prompt's text goes with no
 * `w:del`). Once the control is empty again, Word and the engine show the block its
 * `w:placeholder` names (ECMA-376 §17.5.2.25), not the runs the content replaced: after a reject
 * of the typing, an accepted deletion of it, or a direct deletion of it. Only an undo gives the
 * replaced runs back. The room keeps a loaded prompt's units in the story and hides them while
 * content fills the control, so when units that filled it went (by a decision or a deletion, not
 * taken back), the prompt's units go and the glossary's go in after the words that went. Only
 * that moment counts: a later edit of the prompt shown stays. A control that names no block of
 * the glossary keeps its own runs: the engine keeps them in the glossary (`promptToRetain`).
 */
export function settleNamedPrompts(doc: OrderedDocument, apply: OrderedApplyContext): void {
  const controls = promptedControlsOf(doc);
  if (controls.size === 0) return;
  for (const open of [...controls]) {
    const story = storyOfItem(doc, open);
    if (!story) {
      controls.delete(open);
      continue;
    }

    // Left as before: a control Enter split, or with a pending insertion or deletion on it
    if (open.prompt !== true || isHidden(doc, open) || isSplitEdge(open) || pending(doc, open)) {
      continue;
    }

    settleNamedControl(doc, story, open, apply);
  }
}

/**
 * Show the glossary's prompt in one control once units that filled it went, or give the
 * replaced prompt back when a deletion that went is taken back.
 */
function settleNamedControl(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  apply: OrderedApplyContext
): void {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') {
    return;
  }

  const name = placeholderNameOf(shell);
  const blocks = name === undefined ? undefined : doc.placeholders.get(name);
  if (blocks && giveBack(doc, story, open)) {
    return;
  }

  const scan = scanControl(doc, story, open);
  if (!scan || scan.present) {
    return;
  }

  const fresh = markSettled(doc, scan.went);
  if (fresh.length === 0) {
    return;
  }

  const ctx = new ComposeContext(doc, story, {});
  const decided = fresh.filter((item) => ctx.decidedAway(item));
  if (blocks && !(open.token === 'sdtBlock' && blocks.length > 1)) {
    const control = glossaryControl(doc, story, open, shell, scan.mark);
    if (control) {
      showGlossaryPrompt(doc, story, open, scan, control, apply, decided.length > 0);
    }
  }

  // A decision is final: the words it took away go for good, so a later Enter or edit in the
  // control is no decision's (compose shows a decision's prompt only after one). The control
  // keeps naming the glossary block that its prompt went to when those words filled it.
  if (decided.length > 0) {
    for (const item of decided) {
      relabel(doc, story, item, { gone: true });
    }

    recordFilledPrompt(doc, open.space);
  }
}

/** Mark settled the units that went and were not yet, in the undo record, and return them. */
function markSettled(doc: OrderedDocument, went: readonly OrderedItem[]): OrderedItem[] {
  const state = stateOf(doc);
  const fresh = went.filter((item) => !state.settled.has(item));
  if (fresh.length === 0) {
    return fresh;
  }

  for (const item of fresh) {
    state.settled.add(item);
  }

  recordUndo(doc, () => {
    for (const item of fresh) {
      state.settled.delete(item);
    }
  });
  return fresh;
}

/**
 * The control's prompt and content as it is now, or null when it is not one this settle
 * changes: Enter split it, its prompt has several paragraphs, or it holds a block of its own (a
 * glossary block of several paragraphs in a block control is left too, by the caller).
 */
function scanControl(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): Scan | null {
  const prompt: OrderedItem[] = [];
  const went: OrderedItem[] = [];
  let present = false;
  let mark: OrderedItem | null = null;
  for (const item of ownItems(story.tree.items(story.tree.next(open)), open)) {
    if (isBlockToken(item) || (item.kind === 'zero' && item.block === true)) {
      if (!isHidden(doc, item)) return null;
      continue;
    }

    if (item.kind === 'mark') {
      // A prompt of several paragraphs, or an inline control across paragraphs
      if (item.prompt === true || open.token === 'sdt') return null;
      if (isHidden(doc, item)) continue;
      if (mark) return null;
      mark = item;
      continue;
    }

    if (item.kind !== 'text' && item.kind !== 'atom') continue;
    if (item.prompt === true) {
      if (!item.gone) prompt.push(item);
      continue;
    }

    if (!isHidden(doc, item)) present = true;
    else if (!item.gone && !insertionTakenBack(doc, item)) went.push(item);
  }

  if (open.token === 'sdt') {
    mark = shownMarkOf(doc, story, open);
  }

  return mark && prompt.length > 0 ? { prompt, went, present, mark } : null;
}

/**
 * The control as the engine shows it once emptied, with the glossary's prompt, drawn from its
 * paragraph as compose draws it; null when that prompt is the one it shows now.
 */
function glossaryControl(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  mark: OrderedItem
): OoxmlElement | null {
  const drawn = controlScratch(doc, story, open, shell, mark);
  if (!drawn) {
    return null;
  }

  const now = findContentControl(drawn.scratch, shell.id);
  if (!now) {
    return null;
  }

  const written = replaceNode(drawn.scratch, shell.id, emptied(now));
  if (!written.ok) {
    return null;
  }

  const cleared = written.part;
  const prompted = withRestoredPrompt(cleared, shell.id, drawn.options);
  const control = prompted === cleared ? undefined : findContentControl(prompted, shell.id);
  const shown = now.children.find(isControlContent);
  const prompt = control?.children.find(isControlContent);
  if (!control || !shown || !prompt || ooxmlTreesEqual(shown, prompt)) return null;
  return control;
}

const isRun = (node: OoxmlNode): boolean => node.kind === 'run' || isWml(node, 'r');
const isParagraph = (node: OoxmlNode): boolean => node.kind === 'paragraph' || isWml(node, 'p');

/**
 * The control as an edit that took all its content leaves it: no prompt flag, and no runs in its
 * content (at block level, its paragraph keeps its properties and markers).
 */
function emptied(control: OoxmlElement): OoxmlElement {
  const withoutRuns = (nodes: readonly OoxmlNode[]): OoxmlNode[] =>
    nodes.flatMap((node) => {
      if (node.kind === 'textValue' || isRun(node)) return [];
      return isParagraph(node)
        ? [
            withChildren(
              node,
              node.children.filter((c) => !isRun(c))
            ),
          ]
        : [node];
    });
  return withChildren(
    control,
    control.children.map((child) => {
      if (child.kind === 'textValue') return child;
      if (isWml(child, 'sdtPr'))
        return withChildren(
          child,
          child.children.filter((c) => !isWml(c, 'showingPlcHdr'))
        );
      return isControlContent(child) ? withChildren(child, withoutRuns(child.children)) : child;
    })
  );
}

/** Put the glossary's prompt in place of the control's prompt units. */
function showGlossaryPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  scan: Scan,
  control: OoxmlElement,
  apply: OrderedApplyContext,
  decided: boolean
): void {
  const units = promptUnits(settleContext(doc, story, apply), control, open);
  const after = story.tree.prev(open.token === 'sdt' ? closeOf(story, open) : scan.mark);
  if (units.length === 0 || !after) return;
  insertItems(doc, story, after, units);
  const state = stateOf(doc);
  for (const unit of scan.prompt) {
    relabel(doc, story, unit, { gone: true });
    state.replaced.add(unit);
  }

  recordUndo(doc, () => {
    for (const unit of scan.prompt) state.replaced.delete(unit);
  });
  if (!decided) {
    const swap: Swap = { replaced: scan.prompt, units, went: scan.went };
    state.swaps.set(open, swap);
    recordUndo(doc, () => state.swaps.delete(open));
  }

  recordPromptMarkChange(doc, scan.mark);
}

/**
 * Give the replaced prompt back when a unit whose going made the replacement shows again or is
 * taken back: the glossary's units go, and the replaced units show again. True when it did.
 */
function giveBack(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): boolean {
  const state = states.get(doc);
  const swap = state?.swaps.get(open);
  if (!state || !swap) return false;
  const back = swap.went.some(
    (item) => !item.gone && (!isHidden(doc, item) || insertionTakenBack(doc, item))
  );
  if (!back) return false;
  for (const unit of swap.units) {
    relabel(doc, story, unit, { gone: true });
    state.replaced.add(unit);
  }

  for (const unit of swap.replaced) {
    relabel(doc, story, unit, { gone: false });
    state.replaced.delete(unit);
  }

  state.swaps.delete(open);
  for (const item of swap.went) state.settled.delete(item);
  recordUndo(doc, () => {
    for (const unit of swap.units) state.replaced.delete(unit);
    for (const unit of swap.replaced) state.replaced.add(unit);
    for (const item of swap.went) state.settled.add(item);
    state.swaps.set(open, swap);
  });
  const mark = shownMarkOf(doc, story, open);
  if (mark) recordPromptMarkChange(doc, mark);
  return true;
}
