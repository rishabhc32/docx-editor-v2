// A control an edit or a decision emptied gets its prompt as units, settled after each change
// (task 9.6). The writes go in the change's undo record.

import type { EditOptions } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  withDecisionPrompt,
  withRestoredPrompt,
} from '../../../store/store/content-control-prompt-restore.ts';
import { recordStatesText } from '../../../store/store/tracked-content-control-value.ts';
import { findContentControl } from '../../../store/store/tree-op-nodes.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { isBlockToken } from '../compose/blocks.ts';
import { composeOrderedParagraph } from '../compose/compose.ts';
import { ComposeContext } from '../compose/context.ts';
import { controlContent } from '../compose/inline.ts';
import { ownItems } from '../compose/paragraph.ts';
import { withStoryBlocks } from '../compose/paragraph-ops.ts';
import { withControlRecords } from '../compose/prompts.ts';
import { containerPropertiesAfter } from '../compose/properties.ts';
import { contextOf, withNewIds, type Ctx } from '../edits/context.ts';
import { templateItems } from '../edits/templates.ts';
import { RevisionReader } from '../load/state.ts';
import { isHidden, isSplitEdge, itemAt, statusOf, storyOfItem } from '../model/document.ts';
import { insertItems, relabel, replaceItems } from '../model/mutations.ts';
import { personAttribution } from '../model/person.ts';
import {
  recordDroppedMark,
  recordPromptMarkChange,
  promptlessControlsOf,
} from '../model/prompt-marks.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedItemAddress,
  type OrderedStory,
} from '../model/types.ts';
import { recordUndo } from '../model/undo-record.ts';
import { isControlContent, isWml, withChildren, withId } from '../ooxml-nodes.ts';
import { closeOf } from '../placement.ts';
import type { OrderedApplyContext } from '../room/room.ts';

/**
 * An emptied control: where its prompt's units go, the paragraph mark that shows them, the units
 * that went, and whether a decision took one of them away.
 */
interface Emptied {
  readonly after: OrderedItem;
  readonly mark: OrderedItem;
  readonly went: readonly OrderedItem[];
  readonly decided: boolean;
  /** A block control's later paragraph marks: the engine keeps its first paragraph only. */
  readonly dropped: readonly OrderedItem[];
  /** The edges Enter put in to split the control, which end this part of it. */
  readonly edges: readonly OrderedItem[];
}

/** The units a control lost so far, and whether a decision took one of them away. */
interface Went {
  readonly units: OrderedItem[];
  decided: boolean;
}

/** A control a direct edit emptied, as it was before its prompt: given back if its words are. */
interface PromptedByEdit {
  readonly node: OoxmlNode | null;
  readonly original: OoxmlNode | undefined;
  /** The block control's later paragraph marks the prompt dropped. */
  readonly dropped: readonly OrderedItem[];
  /** The units the edit took away: when one shows again, the edit was taken back. */
  readonly went: readonly OrderedItem[];
  /** Another person's pending insertion the edit took off the control. */
  readonly untracked: Untracked | null;
}

/** Per document, the open items of controls a direct edit emptied, by what they were before. */
const promptedByEdit = new WeakMap<OrderedDocument, Map<OrderedItem, PromptedByEdit>>();

function promptedOf(doc: OrderedDocument): Map<OrderedItem, PromptedByEdit> {
  let found = promptedByEdit.get(doc);
  if (!found) {
    found = new Map();
    promptedByEdit.set(doc, found);
  }

  return found;
}

/** The context a settle writes units in: the change's person, date and id minter. */
export function settleContext(
  doc: OrderedDocument,
  story: OrderedStory,
  apply: OrderedApplyContext
): Ctx {
  const attribution = personAttribution(apply.person, apply.date);
  return contextOf({ doc, story: story.name, attribution, ids: apply.ids }, []);
}

/**
 * Give each control an edit or a decision emptied its prompt as units, and take the prompt of a
 * control a direct edit emptied away again when its words come back. Left as they were: a
 * control with another pending insertion or deletion on it, one that is another person's pending
 * insertion and a decision emptied, and a block control that holds a block of its own.
 */
export function settleEmptiedControls(doc: OrderedDocument, apply: OrderedApplyContext): void {
  unprompt(doc);
  unsplit(doc);
  const controls = promptlessControlsOf(doc);
  const contexts = new Map<OrderedStory, ComposeContext>();
  for (const open of [...controls]) {
    const story = storyOfItem(doc, open);
    if (!story) {
      controls.delete(open);
      continue;
    }

    if (open.prompt === true || isHidden(doc, open)) {
      continue;
    }

    const inserted = insertedPending(doc, open);
    if (!inserted && pending(doc, open)) {
      continue;
    }

    let ctx = contexts.get(story);
    if (!ctx) {
      ctx = new ComposeContext(doc, story, {});
      contexts.set(story, ctx);
    }

    const emptied =
      open.token === 'sdt' ? emptiedInline(ctx, story, open) : emptiedBlock(ctx, story, open);

    // A decision leaves another person's pending control as it was
    if (!emptied || (inserted && emptied.decided)) {
      continue;
    }

    if (givePrompt(doc, story, open, emptied, apply)) {
      controls.delete(open);
    }
  }
}

/**
 * Whether a pending suggestion inserted or struck the control, or a container around it. The
 * edge Enter put in to split a control carries the Enter's suggestion, which is no insertion of
 * the control.
 */
export function pending(doc: OrderedDocument, open: OrderedItem): boolean {
  const by = isSplitEdge(open) ? null : open.by;
  const ids = [by, ...open.dels, ...open.within.by, ...open.within.dels];
  return ids.some((id) => id !== null && statusOf(doc, id) === 'pending');
}

/**
 * Whether the control itself is another person's pending insertion, and nothing else pending is
 * on it: a direct edit that empties it takes that insertion off the control, as the engine drops
 * the person's `w:ins` around a control whose words the organiser deleted (`untrack`).
 */
function insertedPending(doc: OrderedDocument, open: OrderedItem): boolean {
  if (isSplitEdge(open) || open.by === null || statusOf(doc, open.by) !== 'pending') return false;
  const others = [...open.dels, ...open.within.by, ...open.within.dels];
  return !others.some((id) => id !== null && statusOf(doc, id) === 'pending');
}

/** A pending insertion a direct edit took off a control: the suggestion, its units that went. */
interface Untracked {
  readonly by: string;
  readonly units: readonly OrderedItem[];
}

/**
 * The control as the file's own, or as `by` inserted it again: its edges and the given units.
 * Taken off, the words that went filled no prompt.
 */
function relabelInserted(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  units: readonly OrderedItem[],
  by: string | null
): void {
  for (const item of [open, closeOf(story, open), ...units]) relabel(doc, story, item, { by });
}

/**
 * Take another person's pending insertion off a control a direct edit emptied: it is no longer a
 * suggestion, as the engine drops that person's `w:ins` around it. Taking the edit back gives the
 * insertion back (`unprompt`).
 */
function untrack(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  went: readonly OrderedItem[]
): Untracked | null {
  const by = open.by;
  if (by === null || !insertedPending(doc, open)) return null;
  const units = went.filter((item) => item.by === by);
  relabelInserted(doc, story, open, units, null);
  return { by, units };
}

/**
 * Take one item of a control into `went` when it is a unit that went. False when it is a unit
 * that shows: the control is not empty. Units of a prompt that went for good are not the
 * control's words.
 */
function takeWent(ctx: ComposeContext, item: OrderedItem, went: Went): boolean {
  if ((item.kind !== 'text' && item.kind !== 'atom') || item.prompt === true) {
    return true;
  }

  if (!ctx.gone(item)) {
    return false;
  }

  went.units.push(item);
  if (ctx.decidedAway(item)) {
    went.decided = true;
  }

  return true;
}

/**
 * An inline control whose units all went, as compose's `closeBox` finds it: its prompt's units go
 * right before its close item.
 */
function emptiedInline(
  ctx: ComposeContext,
  story: OrderedStory,
  open: OrderedItem
): Emptied | null {
  const went: Went = { units: [], decided: false };
  for (let item = story.tree.next(open); item; item = story.tree.next(item)) {
    if (item.kind === 'close' && item.pair === open.pair) {
      // A part of a control Enter split ends at its edge; an Enter taken back leaves no edge
      if (isSplitEdge(item) && isHidden(ctx.doc, item)) {
        continue;
      }

      if (went.units.length === 0) {
        return null;
      }

      const after = story.tree.prev(item);
      const mark = shownMarkFrom(ctx.doc, story, item);
      const edges = [open, item].filter(isSplitEdge);
      if (!after || !mark) {
        return null;
      }

      return { after, mark, went: went.units, decided: went.decided, dropped: [], edges };
    }

    // A control across paragraphs (Enter split it) is not one compose empties
    if (item.kind === 'mark' || !takeWent(ctx, item, went)) {
      return null;
    }
  }

  return null;
}

/**
 * A block control whose units all went, as compose's `blockElement` finds it: its prompt's units
 * go right before its first paragraph's mark. The engine keeps one paragraph of an emptied
 * control, its first (`valueContent`), so the later ones go (`dropParagraphs`).
 */
function emptiedBlock(ctx: ComposeContext, story: OrderedStory, open: OrderedItem): Emptied | null {
  const went: Went = { units: [], decided: false };
  const marks: OrderedItem[] = [];
  for (let item = story.tree.next(open); item; item = story.tree.next(item)) {
    if (item.kind === 'close' && item.pair === open.pair) {
      const [mark, ...dropped] = marks;
      if (isSplitEdge(item) || went.units.length === 0 || !mark) {
        return null;
      }

      const after = story.tree.prev(mark);
      if (!after) {
        return null;
      }

      return { after, mark, went: went.units, decided: went.decided, dropped, edges: [] };
    }

    if (isBlockToken(item) || (item.kind === 'zero' && item.block === true)) {
      if (!isHidden(ctx.doc, item)) {
        return null;
      }

      continue;
    }

    if (item.kind === 'mark') {
      if (!isHidden(ctx.doc, item)) {
        marks.push(item);
      }

      continue;
    }

    if (!takeWent(ctx, item, went)) {
      return null;
    }
  }

  return null;
}

/** The first paragraph mark at or after `item` that shows: the paragraph `item` is drawn in. */
export function shownMarkFrom(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): OrderedItem | null {
  for (let at: OrderedItem | null = item; at; at = story.tree.next(at))
    if (at.kind === 'mark' && !isHidden(doc, at)) return at;
  return null;
}

/**
 * The paragraph mark that shows a control: for an inline control the one after it, for a block
 * control its first mark that shows.
 */
export function shownMarkOf(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem
): OrderedItem | null {
  if (open.token === 'sdt') {
    return shownMarkFrom(doc, story, closeOf(story, open));
  }

  for (const item of ownItems(story.tree.items(story.tree.next(open)), open)) {
    if (item.kind === 'mark' && !isHidden(doc, item)) {
      return item;
    }
  }

  return null;
}

/** A scratch part with a control's paragraph as compose draws it, and the glossary to read. */
interface ControlScratch {
  readonly scratch: OoxmlPart;
  readonly options: EditOptions;
}

/**
 * A scratch part that holds the paragraph of `mark` as compose draws it, inside the block
 * control when `open` is one, for the engine's prompt rules to run on; null when that paragraph
 * is not drawn.
 */
export function controlScratch(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  shell: OoxmlElement,
  mark: OrderedItem
): ControlScratch | null {
  const index = story.tree.before(mark).marks;
  const paragraph = composeOrderedParagraph(doc, story.name, index);
  if (!paragraph || paragraph.view.mark !== mark) {
    return null;
  }

  let block: OoxmlNode = paragraph.node;
  if (open.token === 'sdtBlock') {
    const ctx = new ComposeContext(doc, story, {});
    const original = doc.originals.get(shell.id);
    const props = containerPropertiesAfter(ctx, open, shell.children);
    block = withChildren(shell, [...props, controlContent(ctx, original, [paragraph.node])]);
  }

  return {
    scratch: withStoryBlocks(story.part, [block]),
    options: { placeholderBlocks: (name: string) => doc.placeholders.get(name) },
  };
}

/**
 * The control as the engine's emptying edit or decision leaves it, drawn from its paragraph as
 * compose draws it, or null when it shows no prompt then. A decision reads its value record back
 * from its prompt, as the engine does: a list or a date showing its prompt holds no value. A
 * reject of typing over the prompt that did not state the record (the file's, beside its prompt)
 * and a direct edit leave the record as it was.
 */
function promptedControl(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  emptied: Emptied
): OoxmlElement | null {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue') {
    return null;
  }

  const drawn = controlScratch(doc, story, open, shell, emptied.mark);
  if (!drawn) {
    return null;
  }

  const restore = emptied.decided ? withDecisionPrompt : withRestoredPrompt;
  const prompted = restore(drawn.scratch, shell.id, drawn.options);
  if (prompted === drawn.scratch) {
    return null;
  }

  // A reject of typing over the prompt: no word that went was struck, and they did not state the
  // record. A field's code is no text of the control, as the engine reads it
  const words = emptied.went.filter((unit) => unit.kind === 'text' && unit.instr !== true);
  const struck = emptied.went.some((unit) =>
    unit.dels.some((id) => statusOf(doc, id) !== 'rejected')
  );
  const typedOver = !struck && !recordStatesText(shell, words.map((unit) => unit.text).join(''));
  const read = emptied.decided && !typedOver ? withControlRecords(prompted, [shell.id]) : prompted;
  return findContentControl(read, shell.id) ?? null;
}

/** The runs of a block control's paragraphs, without their paragraph properties. */
function paragraphRuns(blocks: readonly OoxmlNode[]): OoxmlNode[] {
  return blocks.flatMap((child) =>
    child.kind !== 'textValue' && (child.kind === 'paragraph' || isWml(child, 'p'))
      ? child.children.filter((run) => !isWml(run, 'pPr'))
      : []
  );
}

/**
 * The units of the prompt in a prompted control's content, loaded as a paragraph would be. A
 * record a room checkpoint listed (a prompt the glossary kept with its records) is its
 * suggestion's again.
 */
export function promptUnits(ctx: Ctx, control: OoxmlElement, open: OrderedItem): OrderedItem[] {
  const content = control.children.find(isControlContent);
  if (!content || content.kind === 'textValue') return [];
  const runs = open.token === 'sdt' ? content.children : paragraphRuns(content.children);
  if (runs.length === 0) return [];
  const paragraph = build(ctx.ids.node(), 'paragraph', 'p', [], runs);
  return templateItems(ctx, paragraph, RevisionReader.forPrompt(ctx.doc, open.by ?? ''))
    .filter(isUnit)
    .map(
      (item): OrderedItem => ({
        ...item,
        by: null,
        dels: [],
        within: open.within,
        prompt: true,
      })
    );
}

/**
 * Give one emptied control the prompt the engine writes: a direct edit its type's prompt or the
 * prompt the control names (`withRestoredPrompt`), a decision Word's blank prompt or the named
 * one (`withDecisionPrompt`). Compose alone could draw it, but then the list holds no units for
 * it and typing over the prompt is refused; so, as for a new control's prompt
 * (`edits/control-shell.ts`), the open item shows a prompt, the original is the prompted control,
 * and the prompt's units go in after the words that went. False when it shows none.
 */
function givePrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  emptied: Emptied,
  apply: OrderedApplyContext
): boolean {
  const shell = open.node;
  const control = shell ? promptedControl(doc, story, open, emptied) : null;
  if (!shell || !control) {
    return false;
  }

  const ctx = settleContext(doc, story, apply);
  const units = promptUnits(ctx, control, open);
  if (units.length === 0) {
    return false;
  }

  // The control keeps its id: the part and the list name it by that id
  const original = withId(withNewIds(ctx, control), shell.id);
  const node = withChildren(
    original,
    original.children.filter((child) => !isControlContent(child))
  );
  insertItems(doc, story, emptied.after, units);
  dropParagraphs(doc, story, emptied.dropped, true);
  if (emptied.edges.length > 0) {
    noteSplitPart(doc, story, emptied.edges, units);
  }

  if (emptied.decided) {
    goneForGood(doc, story, emptied.went);
  }

  const prompted: OrderedItem = { ...open, prompt: true, node, leaf: undefined };
  replaceItems(doc, story, [open], [prompted]);
  const untracked = untrack(doc, story, prompted, emptied.went);
  const before = doc.originals.get(shell.id);
  setOriginal(doc, shell.id, original);
  if (!emptied.decided) {
    rememberPromptedByEdit(doc, prompted, {
      node: shell,
      original: before,
      dropped: emptied.dropped,
      went: emptied.went,
      untracked,
    });
  }

  // The paragraph is drawn again with its prompt
  recordPromptMarkChange(doc, emptied.mark);
  return true;
}

/**
 * Mark gone the words a decision took away: a decision is final, and they filled no prompt. A
 * later suggestion over the prompt is what keeps it in the glossary (`filledBy` in
 * `compose/paragraph.ts`).
 */
function goneForGood(doc: OrderedDocument, story: OrderedStory, went: readonly OrderedItem[]) {
  for (const item of went) {
    if (!item.gone) {
      relabel(doc, story, item, { gone: true });
    }
  }
}

/**
 * Note a control a direct edit emptied, as it was before its prompt: a direct edit can be taken
 * back (`revert`), and the control is then given back as it was.
 */
function rememberPromptedByEdit(
  doc: OrderedDocument,
  prompted: OrderedItem,
  before: PromptedByEdit
): void {
  const byEdit = promptedOf(doc);
  byEdit.set(prompted, before);
  recordUndo(doc, () => byEdit.delete(prompted));
}

/**
 * Drop (or give back) a block control's later paragraphs: the engine keeps the first paragraph
 * of an emptied control (`valueContent`), and taking the edit back gives them back. Their marks
 * go, and compose draws nothing of them and joins nothing, as a prompt's own paragraph that
 * content replaced (`isDroppedMark`).
 */
function dropParagraphs(
  doc: OrderedDocument,
  story: OrderedStory,
  marks: readonly OrderedItem[],
  drop: boolean
): void {
  for (const mark of marks) {
    relabel(doc, story, mark, { gone: drop });
    recordDroppedMark(doc, mark, drop);
    recordPromptMarkChange(doc, mark);
  }
}

/** Per document, the prompt units of each part of a split control, by its split edges. */
const splitParts = new WeakMap<
  OrderedDocument,
  Map<OrderedItem, { readonly story: OrderedStory; readonly edges: readonly OrderedItemAddress[] }>
>();

/**
 * Note the units given to a part of a control Enter split. Each part is a control of its own:
 * its prompt's units go before the edge that ends it, and they go when the Enter is taken back,
 * since the parts are one control again (`unsplit`).
 */
function noteSplitPart(
  doc: OrderedDocument,
  story: OrderedStory,
  edges: readonly OrderedItem[],
  units: readonly OrderedItem[]
): void {
  const [first] = units;
  if (!first) {
    return;
  }

  let parts = splitParts.get(doc);
  if (!parts) {
    parts = new Map();
    splitParts.set(doc, parts);
  }

  const record = { story, edges: edges.map(({ space, offset }) => ({ space, offset })) };
  parts.set(first, record);
  recordUndo(doc, () => parts.delete(first));
}

/**
 * Take the prompt of a part of a split control away when the Enter that split it is taken back:
 * the parts are one control again, and it holds the other part's words.
 */
function unsplit(doc: OrderedDocument): void {
  const parts = splitParts.get(doc);
  if (!parts) return;
  for (const [first, record] of [...parts]) {
    const { story } = record;
    const merged = record.edges.some((address) => {
      const edge = itemAt(story, address);
      return edge === null || isHidden(doc, edge);
    });
    if (first.gone || !merged) {
      if (first.gone) parts.delete(first);
      continue;
    }

    for (let unit: OrderedItem | null = first; unit?.prompt === true; unit = story.tree.next(unit))
      if (!unit.gone && isUnit(unit)) relabel(doc, story, unit, { gone: true });
    parts.delete(first);
    recordUndo(doc, () => parts.set(first, record));
    const mark = shownMarkFrom(doc, story, first);
    if (mark) recordPromptMarkChange(doc, mark);
  }
}

/** Set a control's original, in the undo record. */
export function setOriginal(doc: OrderedDocument, id: string, node: OoxmlNode | undefined): void {
  const before = doc.originals.get(id);
  if (node) doc.originals.set(id, node);
  else doc.originals.delete(id);
  recordUndo(doc, () => {
    if (before) doc.originals.set(id, before);
    else doc.originals.delete(id);
  });
}

/**
 * Take the prompt of each control a direct edit emptied away again once a word the edit took
 * away shows (the edit taken back), as the engine's undo leaves it: the prompt goes for good and
 * the control is as it was before.
 */
function unprompt(doc: OrderedDocument): void {
  const byEdit = promptedByEdit.get(doc);
  if (!byEdit) return;
  for (const [open, before] of [...byEdit]) {
    const story = storyOfItem(doc, open);
    if (!story) {
      byEdit.delete(open);
      continue;
    }

    if (isHidden(doc, open) || !before.went.some((item) => !isHidden(doc, item))) {
      continue;
    }

    restoreControl(doc, story, open, before);
    byEdit.delete(open);
    recordUndo(doc, () => byEdit.set(open, before));
  }
}

/**
 * Give a prompted control back as it was before its prompt: its prompt's units go for good, its
 * dropped paragraphs come back, and its open item, insertion and original are what they were.
 */
function restoreControl(
  doc: OrderedDocument,
  story: OrderedStory,
  open: OrderedItem,
  before: PromptedByEdit
): void {
  const units: OrderedItem[] = [];
  let mark: OrderedItem | null = null;
  for (const item of ownItems(story.tree.items(story.tree.next(open)), open)) {
    if (item.kind === 'mark' && !isHidden(doc, item)) {
      mark ??= item;
    }

    if (item.prompt === true && !item.gone) {
      units.push(item);
    }
  }

  mark ??= shownMarkFrom(doc, story, closeOf(story, open));
  for (const unit of units) {
    relabel(doc, story, unit, { gone: true });
  }

  dropParagraphs(doc, story, before.dropped, false);
  const { prompt: _prompt, ...plain } = open;
  const restored: OrderedItem = { ...plain, node: before.node, leaf: undefined };
  replaceItems(doc, story, [open], [restored]);
  const { untracked } = before;
  if (untracked) {
    relabelInserted(doc, story, restored, untracked.units, untracked.by);
  }

  if (open.node) {
    setOriginal(doc, open.node.id, before.original);
  }

  if (mark) {
    recordPromptMarkChange(doc, mark);
  }
}
