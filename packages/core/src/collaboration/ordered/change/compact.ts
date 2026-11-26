// The cleanup at a close (plan D8): a live room removes nothing (design D1), so the items no state
// of the room can show again go here, with every reference to them, and the saved list does not
// grow for the life of the room.

import { isNotesStory, noteKeyOf, referenceKeyOf } from '../compose/note-references.ts';
import {
  indexItem,
  insertionTakenBack,
  isHidden,
  isSplitEdge,
  labelIndex,
  labelUnindex,
  statusOf,
  unindexItem,
} from '../model/document.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  BLOCK_TOKENS,
  type OrderedContainerToken,
  type OrderedDocument,
  type OrderedFormatRecord,
  type OrderedItem,
  type OrderedItemLabels,
  type OrderedPropertyRecord,
  type OrderedStory,
} from '../model/types.ts';
import { forgetEmptiedControls } from './settle-emptied.ts';

/** What a cleanup removed. @public */
export interface OrderedCompaction {
  /** The items it took out of the stories. */
  readonly items: number;
}

/**
 * Containers that read what they held, not only what shows: a link whose every unit went goes
 * too, and a control's value record and prompt follow the units a decision took away
 * (`hadUnits`, `ComposeContext.decided`, `rejectedWords`). The units in them stay.
 */
const HOLDING_TOKENS: ReadonlySet<OrderedContainerToken> = new Set([
  'hyperlink',
  'sdt',
  'smartTag',
  'customXml',
  'dir',
  'bdo',
  'fldSimple',
  'sdtBlock',
  'customXmlBlock',
]);

/**
 * Remove what no state of the room can show again, at a close, before the save. Run it only when
 * no client is connected: a close sends every client `reload`, which forgets its undo history, and
 * the room on the next file refuses an undo or a revert of a change made before it (`other-file`).
 * So the cleanup ends the undo of own edits (`taken`, a record's `change`, `ownEdited`, `added`),
 * and a direct edit is as final as a decision: what it hid goes, and the controls it emptied keep
 * their prompts (`forgetEmptiedControls`).
 *
 * It removes an item hidden for good (`hiddenForGood`) that nothing reads (`deadItems`): no
 * comment range, field slot or kept note names it, nothing it holds or its container holds stays,
 * no container that reads what it held holds it, no row keeps it as a removed cell, and it is not
 * the last item of a space that keeps one (a typed item grows into the next offsets of its space,
 * so an address is never used again). A removed address then names no item, and every reader
 * refuses it (`unknown-address`). The paragraph or block that drew a removed item is still drawn
 * from the items (`rebuilt`). It keeps every suggestion, with its version: a change record in a
 * part the list does not hold (a prompt in the glossary) names its suggestion by its `w:id`.
 * Like every write, it records how to take it back when an undo record is open.
 *
 * @public
 */
export function compactOrdered(doc: OrderedDocument): OrderedCompaction {
  endOwnEditUndo(doc);
  forgetEmptiedControls(doc);
  let items = 0;
  for (const [story, dead] of deadItems(doc)) {
    for (const item of drawnAgain(doc, story, new Set(dead))) {
      markRebuilt(doc, item);
    }

    for (const item of dead) {
      removeItem(doc, story, item);
      items += 1;
    }
  }

  return { items };
}

/**
 * The items whose paragraph or block a removed item made compose draw again from the items: a
 * labelled item touches its paragraph, and a removed mark or block edge touches the block around
 * it. Without it, a loaded paragraph or block would look untouched and compose would write the
 * loaded node, with what the removed item held (`compose/paragraph.ts`, `compose/blocks.ts`).
 */
function drawnAgain(
  doc: OrderedDocument,
  story: OrderedStory,
  dead: ReadonlySet<OrderedItem>
): Set<OrderedItem> {
  const out = new Set<OrderedItem>();
  const blocks: OrderedItem[] = [];
  const touchBlock = () => {
    for (let k = blocks.length - 1; k >= 0; k -= 1) {
      const open = blocks[k];
      if (open !== undefined && !dead.has(open)) {
        out.add(open);
        return;
      }
    }
  };
  let touched = false;
  for (const item of story.tree.items()) {
    const block = item.token !== undefined && BLOCK_TOKENS.has(item.token);
    if (block && item.kind === 'close') {
      blocks.pop();
    }

    if (dead.has(item)) {
      if (block || item.kind === 'mark') {
        touchBlock();
      }

      touched ||= !block;
    } else if (touched && item.kind === 'mark' && !isHidden(doc, item)) {
      out.add(item);
      touched = false;
    }

    if (block && item.kind === 'open') {
      blocks.push(item);
    }
  }

  return out;
}

/**
 * Mark an item drawn again from the items whatever its labels (`rebuilt`, as a load marks what
 * the file labelled): the only write of the flag after a load.
 */
function markRebuilt(doc: OrderedDocument, item: OrderedItem): void {
  if (item.rebuilt === true) {
    return;
  }

  const writable: { rebuilt?: boolean } = item;
  writable.rebuilt = true;
  recordUndo(doc, () => {
    delete writable.rebuilt;
  });
}

/** Take an item out of its story and every index. */
function removeItem(doc: OrderedDocument, story: OrderedStory, item: OrderedItem): void {
  const before = story.tree.prev(item);
  story.tree.remove(item);
  unindexItem(doc, story, item);
  recordUndo(doc, () => {
    story.tree.insertAfter(before, [item]);
    indexItem(doc, story, item);
  });
}

/**
 * Whether no change after the close can show the item again (`isHidden`). Every status is final
 * then: a decision is, and the room refuses the revert of a direct edit made before its file
 * (`other-file`). Only a prompt's units and marks show again, when the content over them goes.
 */
function hiddenForGood(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.prompt === true) {
    return false;
  }

  if (item.gone || item.retracted) {
    return true;
  }

  const rejected = (id: string | null) => statusOf(doc, id) === 'rejected';
  const accepted = (id: string | null) => statusOf(doc, id) === 'accepted';
  if (item.within.by.some(rejected) || item.within.dels.some(accepted)) {
    return true;
  }

  // A split's edge goes with the container it split, or with its Enter taken back
  if (isSplitEdge(item)) {
    return insertionTakenBack(doc, item) || splitContainerGone(doc, item);
  }

  return rejected(item.by) || item.dels.some(accepted);
}

/** Whether both ends of the container a split's edge belongs to are hidden for good. */
function splitContainerGone(doc: OrderedDocument, item: OrderedItem): boolean {
  if (item.pair === undefined) {
    return false;
  }

  for (const story of doc.stories.values()) {
    const ends = story.spaces.get(item.pair);
    if (ends !== undefined && ends.length > 0) {
      return ends.every((end) => hiddenForGood(doc, end));
    }
  }

  return false;
}

/** A unit's address as a comment range or a field slot names it. */
const addressKey = (space: string, offset: number): string => `${space}|${offset}`;

/**
 * The items the cleanup removes, by story, in story order. The notes stories go first: a note
 * whose every reference is hidden is left out by those references (`notesWithHiddenReferences`),
 * so a hidden reference stays while an item of its note stays.
 */
function deadItems(doc: OrderedDocument): Map<OrderedStory, OrderedItem[]> {
  const out = new Map<OrderedStory, OrderedItem[]>();
  const stories = [...doc.stories.values()];
  const keptNotes = new Set<string>();
  for (const story of stories.filter(isNotesStory)) {
    const dead = storyDeadItems(doc, story, keptNotes);
    out.set(story, dead);
    const gone = new Set(dead);
    for (const item of story.tree.items()) {
      const key = noteKeyOf(item);
      if (key !== undefined && item.kind === 'open' && !gone.has(item)) {
        keptNotes.add(key);
      }
    }
  }

  for (const story of stories) {
    if (!isNotesStory(story)) {
      out.set(story, storyDeadItems(doc, story, keptNotes));
    }
  }

  return out;
}

/** The items of a story that the cleanup removes, in story order. */
function storyDeadItems(
  doc: OrderedDocument,
  story: OrderedStory,
  keptNotes: ReadonlySet<string>
): OrderedItem[] {
  const all = story.tree.toArray();
  const named = namedAddresses(doc, story, all);
  const dead = new Set(
    all.filter(
      (item) =>
        hiddenForGood(doc, item) &&
        !named.has(addressKey(item.space, item.offset)) &&
        !keptNotes.has(referenceKeyOf(item) ?? '')
    )
  );
  const containers = new Map<string, OrderedItem[]>();
  for (const item of all) {
    if (item.pair === undefined) {
      continue;
    }

    const edges = containers.get(item.pair);
    if (edges === undefined) {
      containers.set(item.pair, [item]);
    } else {
      edges.push(item);
    }
  }

  // Each rule keeps items, which another rule may read: run them all until none keeps more
  for (let kept = true; kept; ) {
    kept = [
      keepContainers(containers, dead),
      keepHeld(all, dead),
      keepAround(all, dead),
      keepRemovedCells(all, dead),
      keepSpaceEnds(story, dead),
    ].includes(true);
  }

  return all.filter((item) => dead.has(item));
}

/** The addresses a comment range or a field slot in the story names. */
function namedAddresses(
  doc: OrderedDocument,
  story: OrderedStory,
  items: readonly OrderedItem[]
): Set<string> {
  const named = new Set<string>();
  for (const range of doc.commentRanges.values()) {
    if (range.story !== story.name) {
      continue;
    }

    for (const address of [range.start, range.end, range.reference]) {
      if (address !== null) {
        named.add(addressKey(address.space, address.offset));
      }
    }
  }

  for (const item of items) {
    if (item.fieldSlot !== undefined) {
      named.add(item.fieldSlot.field);
    }
  }

  return named;
}

/**
 * Keep every edge of a container (its open and close items, and the edges Enter split it with,
 * which share its `pair`) while one of them stays. True when it kept one.
 */
function keepContainers(
  containers: ReadonlyMap<string, readonly OrderedItem[]>,
  dead: Set<OrderedItem>
): boolean {
  let changed = false;
  for (const edges of containers.values()) {
    if (edges.some((edge) => dead.has(edge)) && edges.some((edge) => !dead.has(edge))) {
      for (const edge of edges) {
        dead.delete(edge);
      }

      changed = true;
    }
  }

  return changed;
}

/** Keep what a container that reads what it held holds, while its open item stays. */
function keepHeld(all: readonly OrderedItem[], dead: Set<OrderedItem>): boolean {
  let holding = 0;
  const open: boolean[] = [];
  let changed = false;
  for (const item of all) {
    if (item.kind === 'close' && item.token !== undefined && HOLDING_TOKENS.has(item.token)) {
      if (open.pop() === true) {
        holding -= 1;
      }
    }

    if (holding > 0 && dead.has(item)) {
      dead.delete(item);
      changed = true;
    }

    if (item.kind === 'open' && item.token !== undefined && HOLDING_TOKENS.has(item.token)) {
      const stays = !dead.has(item);
      open.push(stays);
      if (stays) {
        holding += 1;
      }
    }
  }

  return changed;
}

/**
 * Keep the edges of every container around an item that stays: without them, the item would draw
 * outside the container, where a note or a row compose leaves out kept it from showing.
 */
function keepAround(all: readonly OrderedItem[], dead: Set<OrderedItem>): boolean {
  const opens: OrderedItem[] = [];
  let changed = false;
  for (const item of all) {
    if (item.kind === 'close') {
      opens.pop();
    }

    if (!dead.has(item)) {
      // A container that stays already kept every container around it
      for (let k = opens.length - 1; k >= 0; k -= 1) {
        const open = opens[k];
        if (open === undefined || !dead.has(open)) {
          break;
        }

        dead.delete(open);
        changed = true;
      }
    }

    if (item.kind === 'open') {
      opens.push(item);
    }
  }

  return changed;
}

/**
 * Keep the edges of a cell that went while its row stays: compose draws it as a removed cell, and
 * its table fits the grid to the cells as the engine does (`removedColumnCell`). What it held goes.
 */
function keepRemovedCells(all: readonly OrderedItem[], dead: Set<OrderedItem>): boolean {
  const blocks: OrderedItem[] = [];
  let changed = false;
  for (const item of all) {
    if (item.token === undefined || !BLOCK_TOKENS.has(item.token)) {
      continue;
    }

    if (item.kind === 'close') {
      blocks.pop();
      continue;
    }

    const row = blocks.at(-1);
    if (item.token === 'tc' && dead.has(item) && row?.token === 'tr' && !dead.has(row)) {
      dead.delete(item);
      changed = true;
    }

    blocks.push(item);
  }

  return changed;
}

/** Keep the last item of each space that keeps an item: its offsets are never minted again. */
function keepSpaceEnds(story: OrderedStory, dead: Set<OrderedItem>): boolean {
  let changed = false;
  const spaces = new Set<string>();
  for (const item of dead) {
    spaces.add(item.space);
  }

  for (const space of spaces) {
    const list = story.spaces.get(space) ?? [];
    const last = list.at(-1);
    if (last !== undefined && dead.has(last) && list.some((item) => !dead.has(item))) {
      dead.delete(last);
      changed = true;
    }
  }

  return changed;
}

// ---- the undo of own edits ----

/** A record without the change that made it: only that change's undo read it. */
function withoutChange<T extends OrderedFormatRecord | OrderedPropertyRecord>(record: T): T {
  if (!('change' in record) || record.change === undefined) {
    return record;
  }

  const { change: _change, ...rest } = record;

  // Still a T: `change` is optional on every record type that has it
  return rest as T;
}

/**
 * End the undo of every own edit: no client holds one after a close. A cancelled format record
 * stays while its suggestion is pending, as a take-back of that suggestion gives it back.
 */
function endOwnEditUndo(doc: OrderedDocument): void {
  for (const story of doc.stories.values()) {
    for (const item of story.tree.items()) {
      const labels: Partial<OrderedItemLabels> = {};
      if (item.taken) {
        labels.taken = null;
      }

      if (item.fmts.some((record) => withoutChange(record) !== record)) {
        labels.fmts = item.fmts.map(withoutChange);
      }

      if (item.props.some((record) => withoutChange(record) !== record)) {
        labels.props = item.props.map(withoutChange);
      }

      const cancelled = item.cancelledFmts.filter(
        (record) => statusOf(doc, record.suggestion) === 'pending'
      );
      if (cancelled.length !== item.cancelledFmts.length) {
        labels.cancelledFmts = cancelled;
      }

      if (Object.keys(labels).length > 0) {
        writeLabels(doc, item, labels);
      }
    }
  }

  const ownEdited = new Map(doc.ownEdited);
  const added = new Map(doc.added);
  doc.ownEdited.clear();
  doc.added.clear();
  recordUndo(doc, () => {
    for (const [change, items] of ownEdited) {
      doc.ownEdited.set(change, items);
    }

    for (const [change, ranges] of added) {
      doc.added.set(change, ranges);
    }
  });
}

/**
 * Write labels that change nothing that shows, so the tree's sums stay and no suggestion's
 * version moves (`relabel` moves them).
 */
function writeLabels(
  doc: OrderedDocument,
  item: OrderedItem,
  labels: Partial<OrderedItemLabels>
): void {
  const old: Partial<OrderedItemLabels> = {};
  for (const key of Object.keys(labels) as (keyof OrderedItemLabels)[]) {
    Object.assign(old, { [key]: item[key] });
  }

  const write = (next: Partial<OrderedItemLabels>) => {
    labelUnindex(doc, item);
    Object.assign(item, next);
    labelIndex(doc, item);
  };
  write(labels);
  recordUndo(doc, () => write(old));
}
