// Direct edits (the organiser with tracking off): suggestions settled at once and marked direct,
// so `revert` takes them back exactly (tasks 6.7, 9.12). On pending content a direct edit lays
// its settled record on the units and leaves the pending suggestion as it was.

import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import {
  insertedPending,
  isPending,
  itemAt,
  statusOf,
  suggestionOrNull,
} from '../model/document.ts';
import { relabel, settleDirect } from '../model/mutations.ts';
import type {
  OrderedDocument,
  OrderedFormatWrite,
  OrderedParagraphOp,
  OrderedSuggestion,
} from '../model/types.ts';
import { coveredMarkers } from '../placement.ts';
import { asAuthor, type Ctx } from './context.ts';
import { formatMark, joinParagraphs, paragraphOp, splitParagraph } from './marks.ts';
import { deleteRow } from './tables.ts';
import { deleteText, formatText, insertText, type InsertOptions } from './text.ts';

/** Accept a direct edit's own suggestion, marked as a direct edit's. */
export function settle(doc: OrderedDocument, suggestion: string | null): void {
  if (isPending(doc, suggestion)) settleDirect(doc, suggestion);
}

/** The pending insertion that shows a character with its own marker, if any. */
function pendingInsertion(ctx: Ctx, char: OrderedViewChar | undefined): OrderedSuggestion | null {
  if (!char || char.item.by === null || char.item.quiet) return null;
  if (!insertedPending(ctx.doc, char.item)) return null;
  return ctx.doc.suggestions.get(char.item.by) ?? null;
}

/**
 * The pending new row, note or text box a paragraph sits in, if any: a new note's or text box's
 * paragraph mark carries its insertion unmarked (`quiet`), a new row's paragraphs sit in its row.
 */
function pendingContainerOf(ctx: Ctx, p: OrderedViewParagraph): OrderedSuggestion | null {
  const row = [...p.blocks].reverse().find((block) => block.token === 'tr');
  const container = p.mark.quiet ? p.mark : row;
  if (!container || container.by === null || !insertedPending(ctx.doc, container)) {
    return null;
  }

  return ctx.doc.suggestions.get(container.by) ?? null;
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

  // In a paragraph of a pending row, note or text box: the text is part of it (no marker of its
  // own), as Word writes it there and a load reads it back
  const container = pendingContainerOf(ctx, p);
  if (
    container &&
    !map.some((c) => pendingInsertion(ctx, c) !== null && pendingInsertion(ctx, c) !== container)
  ) {
    insertText(asAuthor(ctx, container), p, offset, text, { quiet: true, ...what });
    return;
  }

  const next = map[offset];
  if (before && !after) {
    // At a pending insertion's end: on the base side, with the next character's face
    settle(
      ctx.doc,
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
    ctx.doc,
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
  settle(ctx.doc, base);
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
  settle(ctx.doc, formatText(ctx, p, start, end, write, null));
}

/** Enter: inside a pending insertion too, the break is base. */
export function directEnter(ctx: Ctx, p: OrderedViewParagraph, offset: number): void {
  settle(ctx.doc, splitParagraph(ctx, p, offset));
}

/** Backspace joins with a settled deletion of the mark, a pending break's too. */
export function directJoin(
  ctx: Ctx,
  previous: OrderedViewParagraph,
  p: OrderedViewParagraph
): void {
  settle(ctx.doc, joinParagraphs(ctx, previous, p, { evenStruck: true }));
}

export function directFormatMark(
  ctx: Ctx,
  p: OrderedViewParagraph,
  write: OrderedFormatWrite
): void {
  settle(ctx.doc, formatMark(ctx, p, write, null));
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

  settle(ctx.doc, into?.id ?? null);
}

/** A settled deletion of the row, struck or not, a pending one too. */
export function directDeleteRow(ctx: Ctx, p: OrderedViewParagraph): void {
  settle(ctx.doc, deleteRow(ctx, p, { evenStruck: true }));
}
