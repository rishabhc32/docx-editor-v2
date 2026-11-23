// Whether a party's change goes inside another person's pending insertion, where Word nests it
// in theirs: at a caret before the change applies, and where its items landed once it applied.

import { isBlockToken } from '../compose/blocks.ts';
import type { OrderedChangeOutcome } from '../change/apply.ts';
import type { OrderedCaret, OrderedTouched } from '../change/targets.ts';
import { isHidden, isPending, madeBy, pendingByOther } from '../model/document.ts';
import type { OrderedMaker } from '../model/person.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';

/**
 * Other people's pending suggestions that inserted an item or a container around it. With no
 * maker, everyone's: a direct edit is no one's suggestion, so every pending one is another's.
 */
export function insertedByOthers(
  doc: OrderedDocument,
  maker: OrderedMaker | null,
  item: OrderedItem
): string[] {
  const other = (id: string) =>
    maker === null ? isPending(doc, id) : pendingByOther(doc, id, maker);
  const out: string[] = [];
  if (item.by !== null && other(item.by)) {
    out.push(item.by);
  }

  for (const id of item.within.by) {
    if (other(id)) {
      out.push(id);
    }
  }

  return out;
}

/** The shown mark that ends the paragraph an item is in: the one Enter there copies. */
export function paragraphMarkOf(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedItem
): OrderedItem | null {
  for (let at: OrderedItem | null = from; at !== null; at = story.tree.next(at))
    if (at.kind === 'mark' && !isHidden(doc, at)) return at;
  return null;
}

/**
 * Whether a row, a cell, a table, a note, a text box or a block control that `id` inserted holds
 * an item (`within`).
 */
const heldByNew = (item: OrderedItem, id: string): boolean =>
  item.within.by.includes(id) || (item.quiet && item.by === id);

/**
 * The open items of the controls and links that hold what comes right after an item, in its
 * paragraph: from a container's open item on, up to its close item.
 */
function containersAfter(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedItem
): OrderedItem[] {
  const out: OrderedItem[] = [];
  let depth = 0;
  for (let item: OrderedItem | null = from; item !== null; item = story.tree.prev(item)) {
    if ((item.kind === 'mark' && !isHidden(doc, item)) || isBlockToken(item)) {
      break;
    }

    if ((item.kind !== 'open' && item.kind !== 'close') || isHidden(doc, item)) {
      continue;
    }

    if (item.kind === 'close') {
      depth += 1;
    } else if (depth > 0) {
      depth -= 1;
    } else {
      out.push(item);
    }
  }

  return out;
}

/**
 * Whether a caret's new items go inside another person's pending insertion, where Word nests
 * them in it: between two of its units (one run of theirs), in a row, a cell, a table, a note, a
 * text box or a block control it made around the caret's paragraph, or in a control or a link it
 * made, from its start on, whatever the words there are. A caret is where the change's new items
 * go: past struck text, as the engine types there (`OrderedCaret`). Beside their paragraph mark,
 * past a link's or a control's end, or between their deletion and their insertion, the new run
 * sits beside their suggestions, as the host's check reads the Word file.
 * With no maker, inside anyone's (`insertedByOthers`): a direct edit there is typed into that
 * insertion as its author's (`directInsert`), as Word extends a tracked insertion with untracked
 * typing, so it makes no suggestion of its own.
 */
export function caretInside(
  doc: OrderedDocument,
  maker: OrderedMaker | null,
  story: OrderedStory | null,
  caret: OrderedCaret
): boolean {
  const { before, at } = caret;
  if (before !== null && isUnit(before) && isUnit(at)) {
    const around = new Set(insertedByOthers(doc, maker, at));
    if (insertedByOthers(doc, maker, before).some((id) => around.has(id))) {
      return true;
    }
  }

  if (story === null) {
    return false;
  }

  const mark = paragraphMarkOf(doc, story, at);
  if (mark !== null && insertedByOthers(doc, maker, mark).some((id) => heldByNew(mark, id))) {
    return true;
  }

  return (
    before !== null &&
    containersAfter(doc, story, before).some(
      (open) => insertedByOthers(doc, maker, open).length > 0
    )
  );
}

/** The caret a run of a suggestion's new items sits at, from its first item; null inside it. */
function landingOf(
  doc: OrderedDocument,
  story: OrderedStory,
  first: OrderedItem,
  id: string
): OrderedCaret | null {
  const shown = (item: OrderedItem) => item.kind !== 'zero' && !isHidden(doc, item);
  let before = story.tree.prev(first);
  while (before !== null && !shown(before)) {
    before = story.tree.prev(before);
  }

  if (before?.by === id) {
    return null;
  }

  let at = story.tree.next(first);
  while (at !== null && (!shown(at) || at.by === id)) {
    at = story.tree.next(at);
  }

  return at === null ? null : { before, at };
}

/**
 * Whether a party's edit reached inside another person's pending insertion, once applied: new
 * items landed in it (`caretInside` at each run), or it struck a unit of their prompt. At a prompt
 * the edits place and strike by the prompt's rules, not by the caret or the range, so the checks
 * before the apply leave it; elsewhere this agrees with them.
 */
export function appliedNested(
  doc: OrderedDocument,
  maker: OrderedMaker,
  touched: OrderedTouched,
  outcome: OrderedChangeOutcome
): boolean {
  const story = touched.story;
  if (story === null) {
    return false;
  }

  for (const id of outcome.suggestions) {
    if (!madeBy(doc, id, maker)) {
      continue;
    }

    for (const item of doc.labelled.get(id) ?? []) {
      if (!story.tree.has(item)) {
        continue;
      }

      const struck = item.prompt === true && item.dels.includes(id);
      if (struck && insertedByOthers(doc, maker, item).length > 0) {
        return true;
      }

      if (item.by !== id) {
        continue;
      }

      // A row, a cell or a table lands where its labels say (`within`); a run of words at a caret
      if (insertedByOthers(doc, maker, item).length > 0) {
        return true;
      }

      const landing = isBlockToken(item) ? null : landingOf(doc, story, item, id);
      if (landing !== null && caretInside(doc, maker, story, landing)) {
        return true;
      }
    }
  }

  return false;
}
