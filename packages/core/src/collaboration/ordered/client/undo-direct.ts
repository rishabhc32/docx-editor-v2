// The organiser's undo and redo of direct edits (tasks 6.7, 9.12).

import { isHidden } from '../model/document.ts';
import { BLOCK_TOKENS, type OrderedDocument, type OrderedItem } from '../model/types.ts';
import {
  isOrderedEditChange,
  ORDERED_CHANGE_LIMITS,
  type OrderedChange,
  type OrderedChangePlace,
} from '../change/types.ts';
import { nextShownMark, paragraphOfMark, placeAt } from './paragraphs.ts';
import { OrderedClientRefusal } from './refusal.ts';

/** A change of the inverse, before it takes an id. */
export type OrderedInverseBody =
  | { readonly kind: 'retract'; readonly suggestions: readonly string[] }
  | {
      readonly kind: 'revert';
      readonly suggestions: readonly string[];
      readonly status: 'accepted' | 'rejected';
    }
  | {
      readonly kind: 'delete';
      readonly mode: 'edit';
      readonly part: string;
      readonly from: OrderedChangePlace;
      readonly to: OrderedChangePlace;
    };

const noInverse = (detail: string): OrderedClientRefusal =>
  new OrderedClientRefusal('ordered-refused', 'no-inverse', detail);

/**
 * The suggestions a change made (`S:<change>.<k>`, from 1 while they exist): its direct ones,
 * or with `direct` false its pending ones.
 */
export function madeBy(doc: OrderedDocument, changeId: string, direct = true): string[] {
  const out: string[] = [];
  for (let k = 1; k <= ORDERED_CHANGE_LIMITS.targets; k += 1) {
    const found = doc.suggestions.get(`S:${changeId}.${k}`);
    if (!found) break;
    if (direct ? found.direct === true : found.status === 'pending') out.push(found.id);
  }

  return out;
}

/** The shown units a change typed into other suggestions: units of its spaces it did not label. */
function typedInside(
  doc: OrderedDocument,
  change: OrderedChange,
  own: ReadonlySet<string>
): OrderedItem[] {
  if (!isOrderedEditChange(change)) return [];
  const story = doc.stories.get(change.part);
  if (!story) return [];
  const out: OrderedItem[] = [];
  for (let k = 1; k <= ORDERED_CHANGE_LIMITS.targets; k += 1) {
    const items = story.spaces.get(`i:${change.id}.${k}`);
    if (!items) break;
    for (const item of items)
      if (item.by !== null && !own.has(item.by) && !isHidden(doc, item)) out.push(item);
  }

  return out;
}

/** Direct deletions of exactly these units, one for each run of them in a paragraph's view. */
function deletionsOf(
  doc: OrderedDocument,
  part: string,
  units: readonly OrderedItem[]
): OrderedInverseBody[] {
  const story = doc.stories.get(part);
  if (!story || units.length === 0) return [];
  const wanted = new Set(units);
  const marks = new Set<OrderedItem>();

  // A unit's paragraph ends at its shown mark, never past a block edge
  const blockEdge = (at: OrderedItem): boolean =>
    (at.kind === 'open' || at.kind === 'close') &&
    at.token !== undefined &&
    BLOCK_TOKENS.has(at.token);
  for (const unit of units) {
    const mark = nextShownMark(doc, story, unit, blockEdge);
    if (mark) {
      marks.add(mark);
    }
  }

  const out: OrderedInverseBody[] = [];
  for (const mark of marks) {
    const paragraph = paragraphOfMark(doc, story, mark);
    if (!paragraph) continue;
    const chars = paragraph.view.chars;
    for (let start = 0; start < chars.length; start += 1) {
      if (!wanted.has(chars[start]!.item)) continue;
      let end = start + 1;
      while (end < chars.length && wanted.has(chars[end]!.item)) end += 1;
      const from = placeAt(paragraph, start);
      const to = placeAt(paragraph, end);
      out.push({ kind: 'delete', mode: 'edit', part, from, to });
      start = end;
    }
  }

  return out;
}

/** `ids` in chunks of at most the format's cap of targets for one change. */
function capped(ids: readonly string[]): string[][] {
  const chunks: string[][] = [];
  for (let k = 0; k < ids.length; k += ORDERED_CHANGE_LIMITS.targets) {
    chunks.push(ids.slice(k, k + ORDERED_CHANGE_LIMITS.targets));
  }

  return chunks;
}

/** `ids` as revert changes under the cap. */
function reverts(ids: readonly string[], status: 'accepted' | 'rejected'): OrderedInverseBody[] {
  return capped(ids).map((suggestions) => ({ kind: 'revert', suggestions, status }));
}

/**
 * The changes that take back a direct step (its changes in order), read from the list now; a
 * refusal by name when a suggestion it turns changed since (it cannot, for one's own direct
 * edits). Redo is the inverse of the undo's changes. A direct edit is a suggestion accepted at
 * once; addresses never change, so its exact inverse is a status turn of the suggestions it made
 * (`revert`): typed text goes, deleted text and the pending insertions it struck show again,
 * Enter joins back, a join splits again, a format or a paragraph property goes off, inside
 * another person's pending insertion too (a direct edit lays its own settled record there).
 * Text the organiser typed inside another person's pending insertion belongs to that insertion,
 * so a direct deletion of exactly those units takes it back. The inverse is read from the list
 * as it is now, so another person's suggestion inside the text stays as it is.
 */
export function inverseOfDirect(
  doc: OrderedDocument,
  changes: readonly OrderedChange[]
): OrderedInverseBody[] {
  const reject: string[] = [];
  const accept: string[] = [];
  const retract: string[] = [];
  const deletions: OrderedInverseBody[] = [];
  for (const change of [...changes].reverse()) {
    // A tracked change of a step that also holds direct ones: its own suggestions go back
    if (isOrderedEditChange(change) && change.mode === 'suggest') {
      retract.push(...madeBy(doc, change.id, false));
      continue;
    }

    if (change.kind === 'revert') {
      for (const id of change.suggestions) {
        const now = doc.suggestions.get(id)?.status;
        if (now !== change.status) throw noInverse(`${id} changed since`);
        (change.status === 'rejected' ? accept : reject).push(id);
      }

      continue;
    }

    if (!isOrderedEditChange(change) || change.mode !== 'edit') throw noInverse(change.kind);
    const made = madeBy(doc, change.id);
    for (const id of made) {
      if (doc.suggestions.get(id)?.status !== 'accepted') throw noInverse(`${id} changed since`);
      reject.push(id);
    }

    deletions.push(...deletionsOf(doc, change.part, typedInside(doc, change, new Set(made))));
  }

  // Empty when nothing of the step shows any more (another person took back the insertion
  // the organiser typed in): the step is spent
  const bodies: OrderedInverseBody[] = [
    ...deletions,
    ...reverts(reject, 'rejected'),
    ...reverts(accept, 'accepted'),
  ];
  for (const suggestions of capped(retract)) {
    bodies.push({ kind: 'retract', suggestions });
  }

  return bodies;
}
