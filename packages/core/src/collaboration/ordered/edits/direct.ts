// Direct edits (the organiser with tracking off): suggestions settled at once and marked direct,
// so `revert` takes them back exactly (tasks 6.7, 9.12). On pending content a direct edit lays
// its settled record on the units and leaves the pending suggestion as it was.

import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import { segmentsOf } from '../containers.ts';
import { controlsTakenWhole } from '../locked-content.ts';
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
import { asAuthor, strikeItems, suggestionFor, type Ctx } from './context.ts';
import { strikeControl } from './control-structure.ts';
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
  const read = new Map<OrderedItem, Surroundings>();
  for (const item of [...(doc.labelled.get(suggestion) ?? [])]) {
    const story = storyOfItem(doc, item);
    if (!story || !isUnit(item) || item.by !== suggestion || item.move !== null) {
      continue;
    }

    const holder =
      into === null ? pendingAround(doc, story, item, read) : { by: into, quiet: false };
    if (holder !== null) {
      relabel(doc, story, item, holder);
      recordAdded(doc, ctx.ids.changeId, unitsOf(story, [item]));
    }
  }
}

/** What is around an item in its paragraph: the pending inline control it is in, and the mark. */
interface Surroundings {
  readonly control: string | null;
  readonly mark: OrderedItem | null;
}

/**
 * The surroundings of every item of the paragraph `item` is in, in one walk: back to the
 * paragraph's start, then forward to its mark with the controls still open at each item.
 */
function readParagraph(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  read: Map<OrderedItem, Surroundings>
): void {
  let start = item;
  for (
    let prev = story.tree.prev(item);
    prev && prev.kind !== 'mark';
    prev = story.tree.prev(prev)
  ) {
    start = prev;
  }

  const items: OrderedItem[] = [];
  const controls: (string | null)[] = [];
  const open: OrderedItem[] = [];
  let mark: OrderedItem | null = null;
  for (let at: OrderedItem | null = start; at; at = story.tree.next(at)) {
    items.push(at);
    controls.push(openControl(doc, open));
    if (at.kind === 'mark') {
      mark = at;
      break;
    }

    if (at.kind === 'open' && at.pair !== undefined) {
      open.push(at);
    } else if (at.kind === 'close') {
      const k = open.map((o) => o.pair).lastIndexOf(at.pair);
      if (k >= 0) {
        open.splice(k, 1);
      }
    }
  }

  for (const [k, at] of items.entries()) {
    read.set(at, { control: controls[k] ?? null, mark });
  }
}

/** The innermost pending new inline control of the open items, if any. */
function openControl(doc: OrderedDocument, open: readonly OrderedItem[]): string | null {
  for (let k = open.length - 1; k >= 0; k -= 1) {
    const item = open[k]!;
    if (item.token === 'sdt' && !isSplitEdge(item) && item.by !== null && isPending(doc, item.by)) {
      return item.by;
    }
  }

  return null;
}

/**
 * The pending insertion that a direct edit's unit sits in, as Word writes it inside the
 * insertion's `w:ins`: the new inline control around it, which one `w:ins` holds whole; else a
 * new row, cell, note or text box, which covers it with no marker of its own (`quiet`). The
 * organiser's own pending insertion takes the unit too. `read` keeps each paragraph's walk.
 */
function pendingAround(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  read: Map<OrderedItem, Surroundings>
): { readonly by: string; readonly quiet: boolean } | null {
  if (!read.has(item)) {
    readParagraph(doc, story, item, read);
  }

  const around = read.get(item);
  if (around?.control) {
    return { by: around.control, quiet: false };
  }

  const container = innermostNewContainer(doc, item);
  if (container !== null) {
    return { by: container, quiet: true };
  }

  const mark = around?.mark;
  return mark?.quiet === true && isPending(doc, mark.by) ? { by: mark.by, quiet: true } : null;
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
  const read = new Map<OrderedItem, Surroundings>();
  readParagraph(doc, story, item, read);
  return read.get(item)?.control ?? null;
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

/** The runs of view characters `[start, end)` outside the controls `taken` names. */
function runsOutside(
  p: OrderedViewParagraph,
  start: number,
  end: number,
  taken: ReadonlyMap<string, OrderedItem>
): [number, number][] {
  const runs: [number, number][] = [];
  let from = start;
  for (let at = start; at < end; at += 1) {
    if (segmentsOf(p.chars[at]?.containers ?? '').some((segment) => taken.has(segment))) {
      if (at > from) {
        runs.push([from, at]);
      }

      from = at + 1;
    }
  }

  if (end > from) {
    runs.push([from, end]);
  }

  return runs;
}

/**
 * A direct deletion of view characters `[start, end)`. Within one paragraph (`wholeControls`) it
 * takes a control it covers whole whose content is locked, as the editor removes it.
 */
export function directDelete(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  options: { readonly wholeControls?: boolean } = {}
): void {
  const taken = options.wholeControls
    ? controlsTakenWhole(ctx, p, start, end)
    : new Map<string, OrderedItem>();
  let into: OrderedSuggestion | null = null;
  for (const open of taken.values()) {
    into ??= suggestionFor(ctx);
    strikeControl(ctx, open, into.id);
  }

  // Every unit, struck or not, base or another person's pending insertion, break or row, leaves
  // with one settled deletion, as the engine removes it: taken back, each shows again as it was
  let base = into?.id ?? null;
  for (const [from, to] of runsOutside(p, start, end, taken)) {
    base = deleteText(ctx, p, from, to, into, { evenStruck: true }) ?? base;
  }

  settle(ctx, base);
  const firstChar = p.chars[start];
  const lastChar = p.chars[end - 1];
  if (base === null || !firstChar || !lastChar) return;
  const first = itemAt(ctx.story, firstChar.address);
  const last = itemAt(ctx.story, lastChar.address);
  if (!first || !last) return;

  // The bookmarks it covers whole go too (rule 12), once where a control it took held them; a
  // comment's range stays
  strikeItems(ctx, ctx.story, coveredMarkers(ctx, first, last), base);
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
