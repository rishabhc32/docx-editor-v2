// Side lists of a document's prompt marks and prompt controls, filled as items enter a story
// (`indexItem`), so a change settles only these and a document with none pays nothing.

import type { OrderedDocument, OrderedItem } from './types.ts';
import { recordUndo } from './undo-record.ts';

/** The paragraph marks of block control prompts (task 9.6), which `settle-prompts.ts` settles. */
const marksOf = new WeakMap<OrderedDocument, Set<OrderedItem>>();

/** Block controls with a prompt, which `settle-prompts.ts` settles. */
const controlsOf = new WeakMap<OrderedDocument, Set<OrderedItem>>();

/**
 * Prompt marks a settle showed or hid, until the client's next read of what a change wrote
 * (`client/dirty.ts`): the change itself does not name them. At most every prompt mark waits.
 */
const changedOf = new WeakMap<OrderedDocument, Set<OrderedItem>>();

/** Controls with no prompt: a decision can empty one, and `settle-emptied.ts` gives its prompt. */
const promptlessOf = new WeakMap<OrderedDocument, Set<OrderedItem>>();

/**
 * Every control with a prompt, inline and block: one that names a glossary prompt shows it once
 * its content went (`settle-named-prompts.ts`).
 */
const promptedOf = new WeakMap<OrderedDocument, Set<OrderedItem>>();

/** Add an item to the document's set in `lists`. */
function addFor(
  lists: WeakMap<OrderedDocument, Set<OrderedItem>>,
  doc: OrderedDocument,
  item: OrderedItem
): void {
  const known = lists.get(doc);
  if (known) {
    known.add(item);
  } else {
    lists.set(doc, new Set([item]));
  }
}

const isControl = (item: OrderedItem): boolean =>
  item.kind === 'open' && (item.token === 'sdt' || item.token === 'sdtBlock');

/**
 * List an item that enters a story: a control without a prompt, a control with a prompt, a
 * block control with a prompt (Enter can remove its prompt for good, and taking the Enter back
 * brings it), and a prompt's paragraph mark.
 */
export function indexPromptItem(doc: OrderedDocument, item: OrderedItem): void {
  if (item.prompt !== true) {
    if (isControl(item)) {
      addFor(promptlessOf, doc, item);
    }

    return;
  }

  if (isControl(item)) {
    addFor(promptedOf, doc, item);
  }

  if (item.kind === 'open' && item.token === 'sdtBlock') {
    addFor(controlsOf, doc, item);
    return;
  }

  if (item.kind === 'mark') {
    addFor(marksOf, doc, item);
  }
}

/**
 * The prompt marks the document's stories took, in no order. A mark taken out of its story
 * stays listed; the settle passes over a mark that is in no story.
 */
export function promptMarksOf(doc: OrderedDocument): ReadonlySet<OrderedItem> {
  return marksOf.get(doc) ?? new Set();
}

/**
 * The controls with no prompt the document's stories took, in no order. A control taken out of
 * its story stays listed until a settle passes over it; it is listed again if it comes back.
 */
export function promptlessControlsOf(doc: OrderedDocument): Set<OrderedItem> {
  return promptlessOf.get(doc) ?? new Set();
}

/**
 * The controls with a prompt, inline and block, the document's stories took, in no order. A
 * control taken out of its story stays listed until a settle passes over it.
 */
export function promptedControlsOf(doc: OrderedDocument): Set<OrderedItem> {
  return promptedOf.get(doc) ?? new Set();
}

/** The block controls with a prompt the document's stories took, in no order. */
export function promptControlsOf(doc: OrderedDocument): ReadonlySet<OrderedItem> {
  return controlsOf.get(doc) ?? new Set();
}

const droppedOf = new WeakMap<OrderedDocument, WeakSet<OrderedItem>>();

/**
 * Record (or forget) a block control's later paragraph mark that its emptying dropped, as the
 * engine keeps only the first paragraph of an emptied control (`change/settle-emptied.ts`).
 */
export function recordDroppedMark(doc: OrderedDocument, mark: OrderedItem, drop: boolean): void {
  let dropped = droppedOf.get(doc);
  if (!dropped) droppedOf.set(doc, (dropped = new WeakSet()));
  if (drop) dropped.add(mark);
  else dropped.delete(mark);
  recordUndo(doc, () => {
    if (drop) dropped.delete(mark);
    else dropped.add(mark);
  });
}

/** Whether an emptying dropped the mark's paragraph: compose draws nothing of it, no join. */
export function isDroppedMark(doc: OrderedDocument, mark: OrderedItem): boolean {
  return droppedOf.get(doc)?.has(mark) === true;
}

const filledOf = new WeakMap<OrderedDocument, Set<string>>();

/**
 * Record a prompt control (by its open item's space) whose prompt a suggestion filled, when the
 * words it filled the prompt with went for good: it still names the glossary block that keeps
 * its prompt, as the engine's control does (`change/settle-named-prompts.ts`).
 */
export function recordFilledPrompt(doc: OrderedDocument, space: string): void {
  let filled = filledOf.get(doc);
  if (!filled) filledOf.set(doc, (filled = new Set()));
  if (filled.has(space)) return;
  filled.add(space);
  recordUndo(doc, () => filled.delete(space));
}

/** Whether a suggestion once filled the prompt of the control whose open item is `space`. */
export function isFilledPrompt(doc: OrderedDocument, space: string): boolean {
  return filledOf.get(doc)?.has(space) === true;
}

/** Record a prompt mark (or a prompt unit) that a settle showed or hid. */
export function recordPromptMarkChange(doc: OrderedDocument, mark: OrderedItem): void {
  addFor(changedOf, doc, mark);
}

/** The prompt marks settles showed or hid since the last call, which forgets them. */
export function takePromptMarkChanges(doc: OrderedDocument): OrderedItem[] {
  const changed = changedOf.get(doc);
  if (!changed || changed.size === 0) return [];
  const out = [...changed];
  changed.clear();
  return out;
}
