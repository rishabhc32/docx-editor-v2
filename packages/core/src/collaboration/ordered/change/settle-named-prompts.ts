// A control that names a glossary prompt shows that prompt once its content went, settled after
// each change (task 9.6). The writes go in the change's undo record.

import { isHyperlinkNode } from '../../../store/package/hyperlink.ts';
import { createNodeIdAllocator, replaceNode } from '../../../store/package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { withRestoredPrompt } from '../../../store/store/content-control-prompt-restore.ts';
import {
  contentControlPropertiesOf,
  findContentControl,
} from '../../../store/store/tree-op-nodes.ts';
import { isBlockToken } from '../compose/blocks.ts';
import { ComposeContext } from '../compose/context.ts';
import { composeInline } from '../compose/inline.ts';
import { ownItems } from '../compose/paragraph.ts';
import { placeholderNameOf, promptRecords, propertiesReadBack } from '../compose/prompts.ts';
import {
  insertionTakenBack,
  isHidden,
  isLabelled,
  isSplitEdge,
  storyOfItem,
} from '../model/document.ts';
import { insertItems, relabel, replaceItems } from '../model/mutations.ts';
import {
  recordFilledPrompt,
  recordPromptMarkChange,
  promptedControlsOf,
} from '../model/prompt-marks.ts';
import { suggestionOfRevision } from '../model/revision-ids.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { isControlContent, isParagraphProperties, isWml, withChildren } from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import type { OrderedApplyContext } from '../room/room.ts';
import {
  controlScratch,
  pending,
  promptPlace,
  promptUnits,
  readsRecordBack,
  setOriginal,
  settleContext,
  shownMarkOf,
} from './settle-emptied.ts';

/** A control's prompt as it is now: its units that show, the units that went, and its mark. */
interface Scan {
  /** The prompt's units and the markers made on it (a comment's range) that are not gone. */
  readonly prompt: readonly OrderedItem[];
  /** Units that filled the prompt and went by a decision or a deletion, not taken back. */
  readonly went: readonly OrderedItem[];
  /** Whether content shows in the control, so its prompt does not. */
  readonly present: boolean;
  /** Whether units that filled the prompt were taken back: that fill never was. */
  readonly takenBack: boolean;
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
  /** Controls whose prompt a first fill kept, by their open item's space. */
  readonly kept: Set<string>;
}

const states = new WeakMap<OrderedDocument, State>();

function stateOf(doc: OrderedDocument): State {
  let found = states.get(doc);
  if (!found) {
    found = {
      swaps: new Map(),
      settled: new WeakSet(),
      replaced: new WeakSet(),
      kept: new Set(),
    };
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
 * the glossary keeps its own runs: the engine keeps them in the glossary (`promptToRetain`). A
 * prompt goes where the engine puts it (`promptPlace`), so its own runs go there too when a
 * marker made on the content stands between them and that place.
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
  if (giveBack(doc, story, open)) {
    return;
  }

  const scan = scanControl(doc, story, open);
  if (!scan) {
    return;
  }

  // A fill taken back never was: the next fill keeps the prompt as it shows then
  const { kept } = stateOf(doc);
  if (!scan.present && scan.takenBack && kept.delete(open.space)) {
    recordUndo(doc, () => kept.add(open.space));
  }

  if (scan.present) {
    if (!blocks) {
      keepPrompt(doc, story, open, shell, scan);
    }

    return;
  }

  const fresh = markSettled(doc, scan.went);
  if (fresh.length === 0) {
    return;
  }

  const ctx = new ComposeContext(doc, story, {});
  const decided = fresh.filter((item) => ctx.decidedAway(item));
  const glossary =
    blocks && !(open.token === 'sdtBlock' && blocks.length > 1)
      ? glossaryControl(doc, story, open, shell, scan.mark)
      : null;
  const control =
    glossary ??
    (blocks ? null : keptPrompt(doc, story, open, shell, scan)) ??
    misplacedPrompt(doc, story, open, shell, scan);
  if (control) {
    swapPrompt(doc, story, open, scan, control, apply, decided.length > 0);
  }

  // A decision is final: the words it took away go for good, so a later Enter or edit in the
  // control is no decision's (compose shows a decision's prompt only after one). The control
  // keeps naming the glossary block that its prompt went to when those words filled it.
  if (decided.length > 0) {
    readValueBack(doc, story, open, shell, decided);
    for (const item of decided) {
      relabel(doc, story, item, { gone: true });
    }

    recordFilledPrompt(doc, open.space);
  }
}

/**
 * The control's value record read back from the prompt it shows again, as the engine reads it
 * after a decision (`readsRecordBack`): the words the decision took away go for good, so no
 * later drawing can tell they stated it.
 */
function readValueBack(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  went: readonly OrderedItem[]
): void {
  const properties = contentControlPropertiesOf(shell);
  const edited = readsRecordBack(doc, shell, went)
    ? propertiesReadBack(shell, false, createNodeIdAllocator(story.part))
    : undefined;
  if (!properties || !edited) {
    return;
  }

  const node = withChildren(
    shell,
    shell.children.map((child) => (child === properties ? edited : child))
  );
  replaceItems(doc, story, [open], [{ ...open, node, leaf: undefined }]);
}

/**
 * Keep a control's prompt as content first fills it, as the engine keeps it in the glossary
 * (`promptToRetain`). Compose keeps the prompt of the control's original: once records are on the
 * prompt (a suggested format), the original shows it as drawn with them. A later fill keeps the
 * first one, as the engine's control names the block it kept. A format on the prompt taken back
 * after that is taken off the kept prompt too: it never was.
 */
function keepPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  scan: Scan
): void {
  const { kept } = stateOf(doc);
  const first = !kept.has(open.space);
  if (first) {
    kept.add(open.space);
    recordUndo(doc, () => kept.delete(open.space));
  }

  const redraw = first
    ? promptItems(story, open).some((item) => !item.gone && isLabelled(item))
    : formatTakenOff(doc, shell, scan);
  const control = redraw ? drawnPrompt(doc, story, open, shell, scan) : null;
  if (control) {
    setOriginal(doc, shell.id, control);
  }
}

/**
 * Whether the kept prompt holds a format record that none of the prompt's units holds now: its
 * suggestion took it back.
 */
function formatTakenOff(doc: OrderedDocument, shell: OoxmlElement, scan: Scan): boolean {
  const original = doc.originals.get(shell.id);
  if (!original) {
    return false;
  }

  const held = new Set(scan.prompt.flatMap((unit) => unit.fmts.map((f) => f.suggestion)));
  return promptRecords(original).some((record) => {
    const suggestion = isWml(record, 'rPrChange') ? suggestionOfRevision(doc, record) : undefined;
    return suggestion !== undefined && !held.has(suggestion);
  });
}

/**
 * A control's prompt items, in the containers that hold its words (the organiser's link on
 * them), as the engine keeps the control's content whole.
 */
function promptItems(story: OrderedStory, open: OrderedItem): OrderedItem[] {
  const own = [...ownItems(story.tree.items(story.tree.next(open)), open)];
  const around: string[] = [];
  const holding = new Set<string>();
  for (const item of own) {
    if (item.kind === 'open' && item.pair !== undefined) {
      around.push(item.pair);
    } else if (item.kind === 'close') {
      around.pop();
    } else if (item.prompt === true && isUnit(item)) {
      for (const pair of around) {
        holding.add(pair);
      }
    }
  }

  return own.filter(
    (item) =>
      item.kind !== 'mark' &&
      (item.prompt === true || (item.pair !== undefined && holding.has(item.pair)))
  );
}

/** The control's original, its content drawn from the prompt's items as compose draws them. */
function drawnPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  scan: Scan
): OoxmlElement | null {
  const original = doc.originals.get(shell.id);
  if (!original || original.kind === 'textValue') {
    return null;
  }

  const content = original.children.find(isControlContent);
  if (!content || content.kind === 'textValue') {
    return null;
  }

  const prompt = promptItems(story, open);
  const ctx = new ComposeContext(doc, story, {});
  ctx.paragraphKey = scan.mark.space;
  const drawn =
    open.token === 'sdt'
      ? composeInline(ctx, prompt)
      : content.children.map((paragraph) => {
          if (!isParagraph(paragraph) || paragraph.kind === 'textValue') {
            return paragraph;
          }

          const units = prompt.filter((item) => item.home === paragraph.id);
          const pPr = paragraph.children.filter(isParagraphProperties);
          return withChildren(paragraph, [...pPr, ...composeInline(ctx, units)]);
        });
  return withChildren(
    original,
    original.children.map((child) => (child === content ? withChildren(content, drawn) : child))
  );
}

/**
 * The control with the prompt it kept when content first filled it (its original, `keepPrompt`),
 * when no glossary block the room holds names it yet and the prompt's units show otherwise now (a
 * format made after that fill): the engine shows the block it kept, which a fold writes. Null
 * when the room kept no prompt, the units show it, or the prompt holds a control (`holdsControl`).
 */
function keptPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  scan: Scan
): OoxmlElement | null {
  const original = doc.originals.get(shell.id);
  if (
    !stateOf(doc).kept.has(open.space) ||
    !original ||
    original.kind === 'textValue' ||
    holdsControl(story, open)
  ) {
    return null;
  }

  const kept = original.children.find(isControlContent);
  const shown = drawnPrompt(doc, story, open, shell, scan)?.children.find(isControlContent);
  return kept && shown && !ooxmlTreesEqual(kept, shown) ? original : null;
}

/**
 * The control with its own prompt, when the prompt that shows again is not where the engine
 * shows it (`promptPlace`): a comment made on the content that filled it keeps its range around
 * the prompt.
 */
function misplacedPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  scan: Scan
): OoxmlElement | null {
  // In place when the prompt's own items come last, but for the markers that close the content,
  // with no other item between them (a comment made on part of the prompt). A prompt that holds a
  // control stays too: a swap gives back the prompt's units only.
  const after = promptPlace(doc, story, promptEnd(story, open, scan));
  if (!after || holdsControl(story, open)) {
    return null;
  }

  if (after.prompt === true && standsTogether(doc, story, promptItems(story, open))) {
    return null;
  }

  return drawnPrompt(doc, story, open, shell, scan);
}

/** Whether items stand together: no item that shows lies between them but their own. */
function standsTogether(
  doc: OrderedDocument,
  story: OrderedStory,
  items: readonly OrderedItem[]
): boolean {
  const own = new Set(items);
  const last = items.at(-1);
  for (let at: OrderedItem | null = items[0] ?? null; at && at !== last; at = story.tree.next(at)) {
    if (!own.has(at) && !isHidden(doc, at)) {
      return false;
    }
  }

  return true;
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
  let takenBack = false;
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

    // A marker a fold read as the prompt's (a comment's range) goes and comes back with its units
    if (item.prompt === true && (isUnit(item) || item.kind === 'zero')) {
      if (!item.gone) {
        prompt.push(item);
      }

      continue;
    }

    if (!isUnit(item)) {
      continue;
    }

    if (!isHidden(doc, item)) {
      present = true;
    } else if (insertionTakenBack(doc, item)) {
      takenBack = true;
    } else if (!item.gone) {
      went.push(item);
    }
  }

  if (open.token === 'sdt') {
    mark = shownMarkOf(doc, story, open);
  }

  return mark && prompt.some(isUnit) ? { prompt, went, present, takenBack, mark } : null;
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

/** Whether a control's prompt holds a control: it stays, as a swap gives back units only. */
const holdsControl = (story: OrderedStory, open: OrderedItem): boolean =>
  itemsInside(story, open).some((item) => item.kind === 'open' && item.prompt === true);
const isRun = (node: OoxmlNode): boolean => node.kind === 'run' || isWml(node, 'r');
const isParagraph = (node: OoxmlNode): boolean => node.kind === 'paragraph' || isWml(node, 'p');

/** A run, or a link with its runs: the engine drops a link whose words all went. */
const isWords = (node: OoxmlNode): boolean => isRun(node) || isHyperlinkNode(node);

/**
 * The control as an edit that took all its content leaves it: no prompt flag, and no runs or
 * links in its content (at block level, its paragraph keeps its properties and markers).
 */
function emptied(control: OoxmlElement): OoxmlElement {
  const withoutWords = (nodes: readonly OoxmlNode[]): OoxmlNode[] =>
    nodes.flatMap((node) => {
      if (node.kind === 'textValue' || isWords(node)) {
        return [];
      }

      return isParagraph(node)
        ? [
            withChildren(
              node,
              node.children.filter((c) => !isWords(c))
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
      return isControlContent(child) ? withChildren(child, withoutWords(child.children)) : child;
    })
  );
}

/** Where a control's prompt ends: its close item inline, its paragraph mark at block level. */
const promptEnd = (story: OrderedStory, open: OrderedItem, scan: Scan): OrderedItem =>
  open.token === 'sdt' ? closeOf(story, open) : scan.mark;

/**
 * Put the control's prompt (the glossary's, or its own at the engine's place) in place of its
 * prompt units.
 */
function swapPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  scan: Scan,
  control: OoxmlElement,
  apply: OrderedApplyContext,
  decided: boolean
): void {
  const units = promptUnits(settleContext(doc, story, apply), control, open);
  if (units.length === 0) {
    return;
  }

  // The replaced prompt goes first, so the new one goes where the engine puts it among what is
  // left: inside a comment's range made on the prompt. The control's open item shows, so the place
  // is at the latest right after it
  const state = stateOf(doc);
  for (const unit of scan.prompt) {
    relabel(doc, story, unit, { gone: true });
    state.replaced.add(unit);
  }

  insertItems(doc, story, promptPlace(doc, story, promptEnd(story, open, scan)) ?? open, units);

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
