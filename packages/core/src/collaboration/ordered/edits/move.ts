// Moves and copies: edits whose items land in two places at once.

import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import { currentObjectNode } from '../compose/text-boxes.ts';
import { isolateUnits, relabel } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { currentFace, suggestionFor, withNewIds, type Ctx } from './context.ts';
import { refuseLocked } from '../placement.ts';
import { groupBySpace, insertText, type Piece } from './text.ts';

/** Refuse words a move or a copy cannot carry: a prompt, a field, or an object outside a run. */
function refuseUncopyable(chars: readonly OrderedViewChar[]): void {
  if (chars.some((c) => c.item.prompt === true)) {
    throw new OrderedRefusal('move-of-prompt');
  }

  if (chars.some((c) => c.item.kind === 'atom' && (!c.item.node || c.item.pieces || c.item.bare))) {
    throw new OrderedRefusal('move-of-container', 'a field or an object outside a run');
  }
}

/**
 * The pieces that type a copy of view characters: one for each run of characters of one item,
 * in its face as it shows now; an object gets a node of its own.
 */
function piecesOf(ctx: Ctx, chars: readonly OrderedViewChar[]): Piece[] {
  const pieces: Piece[] = [];
  for (let k = 0; k < chars.length; ) {
    const item = chars[k]?.item;
    if (!item) {
      break;
    }

    let n = k;
    while (n < chars.length && chars[n]?.item === item) {
      n += 1;
    }

    const face = currentFace(ctx, item);

    // An object with a text box takes the text box as it shows now
    const node = currentObjectNode(ctx.doc, ctx.story, item);
    if (item.kind === 'atom' && node && node.kind !== 'textValue') {
      for (let a = k; a < n; a += 1) {
        pieces.push({ text: item.text, face, content: withNewIds(ctx, node) });
      }
    } else {
      const text = chars
        .slice(k, n)
        .map((c) => c.ch)
        .join('');
      pieces.push({ text, face });
    }

    k = n;
  }

  return pieces;
}

/**
 * Move words: the source struck as moved-from text, a copy at the destination as moved-to
 * text, in one suggestion (Word records a move this way; the engine has no gesture). The source
 * is base text in one paragraph; the copy takes the source's face, and objects get nodes of
 * their own.
 */
export function moveText(
  ctx: Ctx,
  from: OrderedViewParagraph,
  start: number,
  end: number,
  to: OrderedViewParagraph,
  offset: number
): string {
  refuseLocked(ctx, from, start, end);
  refuseLocked(ctx, to, offset, offset);
  const chars = from.chars.slice(start, end);
  if (chars.length === 0) throw new OrderedRefusal('empty-range', 'a move of nothing');
  if (chars.some((c) => c.item.by !== null || c.deleted))
    throw new OrderedRefusal('move-of-suggestion', 'words another suggestion holds');
  refuseUncopyable(chars);
  if (to.id === from.id && offset > start && offset < end)
    throw new OrderedRefusal('move-into-itself');
  const suggestion = suggestionFor(ctx);
  const name = ctx.ids.move();
  const pieces = piecesOf(ctx, chars);

  // The destination first: striking the source leaves every view offset where it is
  insertText({ ...ctx, join: suggestion.id }, to, offset, '', {
    pieces,
    move: name,
    outsideOf: '',
  });
  for (const group of groupBySpace(chars))
    for (const item of isolateUnits(ctx.doc, ctx.story, group.space, group.start, group.end))
      relabel(ctx.doc, ctx.story, item, { dels: [...item.dels, suggestion.id], move: name });
  return suggestion.id;
}

// ---- copies ----

/**
 * Copy words and objects of one paragraph to another place, as an insertion that keeps their
 * faces: a paste of the document's own content. An object's copy keeps its drawing id, as the
 * engine's paste of a picture keeps it.
 */
export function copyText(
  ctx: Ctx,
  from: OrderedViewParagraph,
  start: number,
  end: number,
  to: OrderedViewParagraph,
  offset: number
): string {
  refuseLocked(ctx, from, start, end);
  refuseLocked(ctx, to, offset, offset);
  const chars = from.chars.slice(start, end);
  if (chars.length === 0) throw new OrderedRefusal('empty-range', 'a copy of nothing');
  refuseUncopyable(chars);
  const pieces = piecesOf(ctx, chars);
  return insertText(ctx, to, offset, '', { pieces });
}
