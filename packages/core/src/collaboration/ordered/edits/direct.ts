// Direct edits (the organiser with tracking off): suggestions settled at once and marked direct,
// so `revert` takes them back exactly (tasks 6.7, 9.12). On pending content a direct edit lays
// its settled record on the units and leaves the pending suggestion as it was.

import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import {
  innermostNewContainer,
  insertedPending,
  isPending,
  isSplitEdge,
  itemAt,
  statusOf,
  storyOfItem,
  suggestionOrNull,
} from '../model/document.ts';
import { recordAdded, relabel, settleDirect } from '../model/mutations.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedFormatWrite,
  type OrderedItem,
  type OrderedParagraphOp,
  type OrderedStory,
  type OrderedSuggestion,
} from '../model/types.ts';
import { coveredMarkers } from '../placement.ts';
import { asAuthor, type Ctx } from './context.ts';
import { formatMark, joinParagraphs, paragraphOp, splitParagraph } from './marks.ts';
import { deleteRow } from './tables.ts';
import { deleteText, formatText, insertText, unitsOf, type InsertOptions } from './text.ts';

/**
 * Accept a direct edit's own suggestion, marked as a direct edit's. Its words and objects in
 * another person's pending insertion are that insertion's (`into` when given, else
 * `pendingAround`), as Word writes them there and a load reads them back; noted as units the
 * change added to it, so its undo deletes them (`inverseOfDirect`). A move's words keep their
 * move.
 */
export function settle(ctx: Ctx, suggestion: string | null, into: string | null = null): void {
  const doc = ctx.doc;
  if (!isPending(doc, suggestion)) {
    return;
  }

  settleDirect(doc, suggestion);
  for (const item of [...(doc.labelled.get(suggestion) ?? [])]) {
    const story = storyOfItem(doc, item);
    if (!story || !isUnit(item) || item.by !== suggestion || item.move !== null) {
      continue;
    }

    const holder = into === null ? pendingAround(doc, story, item) : { by: into, quiet: false };
    if (holder !== null) {
      relabel(doc, story, item, holder);
      recordAdded(doc, ctx.ids.changeId, unitsOf(story, [item]));
    }
  }
}

/**
 * The pending insertion of another person that a direct edit's unit sits in, as Word writes it
 * inside their `w:ins`: the new inline control around it, which one `w:ins` holds whole; else
 * their new row, cell, note or text box, which covers it with no marker of its own (`quiet`).
 */
function pendingAround(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): { readonly by: string; readonly quiet: boolean } | null {
  const control = pendingControlAround(doc, story, item);
  if (control !== null) {
    return { by: control, quiet: false };
  }

  const container = pendingContainerAround(doc, story, item);
  return container === null ? null : { by: container, quiet: true };
}

/**
 * The pending insertion of a new inline control around the item, in its paragraph, if any: one
 * `w:ins` holds the control whole, with all that is in it.
 */
export function pendingControlAround(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): string | null {
  const closed = new Set<string>();
  for (
    let prev = story.tree.prev(item);
    prev && prev.kind !== 'mark';
    prev = story.tree.prev(prev)
  ) {
    if (prev.pair === undefined) {
      continue;
    }

    if (prev.kind === 'close') {
      closed.add(prev.pair);
    } else if (prev.token === 'sdt' && !closed.has(prev.pair) && !isSplitEdge(prev)) {
      const by = prev.by;
      if (by !== null && isPending(doc, by)) {
        return by;
      }
    }
  }

  return null;
}

/**
 * The pending new row, cell, note or text box a unit sits in, if any: the innermost new row or
 * cell around it, else the note or text box whose paragraph mark carries its insertion unmarked
 * (`quiet`).
 */
function pendingContainerAround(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): string | null {
  const container = innermostNewContainer(doc, item);
  if (container !== null) {
    return container;
  }

  let mark = story.tree.next(item);
  while (mark && mark.kind !== 'mark') {
    mark = story.tree.next(mark);
  }

  return mark?.quiet === true && isPending(doc, mark.by) ? mark.by : null;
}

/** The pending insertion that shows a character with its own marker, if any. */
function pendingInsertion(ctx: Ctx, char: OrderedViewChar | undefined): OrderedSuggestion | null {
  if (!char || char.item.by === null || char.item.quiet) {
    return null;
  }

  if (!insertedPending(ctx.doc, char.item)) {
    return null;
  }

  return ctx.doc.suggestions.get(char.item.by) ?? null;
}

/** Type text (or one object, `pieces`) directly at a view offset. */
export function directInsert(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  text: string,
  options: Pick<InsertOptions, 'pieces' | 'inside'> = {}
): void {
  const map = p.chars;
  const { pieces } = options;
  const what = pieces === undefined ? options : { ...options, outsidePrompt: true };
  const before = pendingInsertion(ctx, map[offset - 1]);
  const after = pendingInsertion(ctx, map[offset]);

  // Strictly inside a pending insertion: into it, as its author's (the engine types inside the
  // `w:ins`); undo takes those units back with a direct deletion
  if (before && after === before && map[offset - 1]?.address.space === map[offset]?.address.space) {
    insertText(asAuthor(ctx, before), p, offset, text, what);
    return;
  }

  // Where two parts of one pending insertion meet (typed by two changes): into it too, but in the
  // edit's own space, so the units name who typed them (as the author, it would grow a part)
  if (before && after === before) {
    settle(ctx, insertText(ctx, p, offset, text, what), before.id);
    return;
  }

  const next = map[offset];
  if (before && !after) {
    // At a pending insertion's end: on the base side, with the next character's face
    settle(
      ctx,
      insertText(ctx, p, offset, text, {
        face: next?.item.face ?? null,
        outsideOf: before.id,
        ...what,
      })
    );
    return;
  }

  // Inside text under a pending format: the base face, and the pending format covers it too.
  // Not at a prompt: typed text replaces it in the prompt's face alone (rule 7, `promptSite`)
  const neighbour = offset > 0 ? map[offset - 1] : map[offset];
  const prompt = pieces === undefined && [map[offset], map[offset - 1]].some((c) => c?.item.prompt);
  const formats =
    neighbour && !neighbour.deleted && !prompt
      ? neighbour.item.fmts.filter((f) => statusOf(ctx.doc, f.suggestion) === 'pending')
      : [];
  settle(
    ctx,
    insertText(
      ctx,
      p,
      offset,
      text,
      neighbour && formats.length > 0 ? { face: neighbour.item.face, fmts: formats, ...what } : what
    )
  );
}

export function directDelete(ctx: Ctx, p: OrderedViewParagraph, start: number, end: number): void {
  // Every unit, struck or not, base or another person's pending insertion, break or row, leaves
  // with one settled deletion, as the engine removes it: taken back, each shows again as it was
  const base = deleteText(ctx, p, start, end, null, { evenStruck: true });
  settle(ctx, base);
  const firstChar = p.chars[start];
  const lastChar = p.chars[end - 1];
  if (base === null || !firstChar || !lastChar) return;
  const first = itemAt(ctx.story, firstChar.address);
  const last = itemAt(ctx.story, lastChar.address);
  if (!first || !last) return;

  // The bookmarks it covers whole go too (rule 12); a comment's range stays
  for (const marker of coveredMarkers(ctx, first, last))
    relabel(ctx.doc, ctx.story, marker, { dels: [...marker.dels, base] });
}

/** A format: inside a pending insertion, a settled record over its units, as the engine draws. */
export function directFormat(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  write: OrderedFormatWrite
): void {
  settle(ctx, formatText(ctx, p, start, end, write, null));
}

/** Enter: inside a pending insertion too, the break is base. */
export function directEnter(ctx: Ctx, p: OrderedViewParagraph, offset: number): void {
  settle(ctx, splitParagraph(ctx, p, offset));
}

/** Backspace joins with a settled deletion of the mark, a pending break's too. */
export function directJoin(
  ctx: Ctx,
  previous: OrderedViewParagraph,
  p: OrderedViewParagraph
): void {
  settle(ctx, joinParagraphs(ctx, previous, p, { evenStruck: true }));
}

export function directFormatMark(
  ctx: Ctx,
  p: OrderedViewParagraph,
  write: OrderedFormatWrite
): void {
  settle(ctx, formatMark(ctx, p, write, null));
}

export function directParagraphOps(
  ctx: Ctx,
  p: OrderedViewParagraph,
  ops: readonly OrderedParagraphOp[]
): void {
  let into: OrderedSuggestion | null = null;
  for (const op of ops) {
    into = suggestionOrNull(ctx.doc, paragraphOp(ctx, p, op, into));
  }

  settle(ctx, into?.id ?? null);
}

/** A settled deletion of the row, struck or not, a pending one too. */
export function directDeleteRow(ctx: Ctx, p: OrderedViewParagraph): void {
  settle(ctx, deleteRow(ctx, p, { evenStruck: true }));
}
