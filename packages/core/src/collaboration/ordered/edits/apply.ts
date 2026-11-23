// One edit, applied in place inside an undo record: it either holds whole or is refused by
// name and changes nothing. Places are addresses; each edit resolves them in the view as it is
// now, then places its items by the engine's rules.

import type { ContentControlValueInput } from '../../../store/store/tree-op-content-controls.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { isLinkSegment, openOfSegment, segmentsOf } from '../containers.ts';
import { itemAt, suggestionOrNull } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedFormatWrite, OrderedParagraphOp } from '../model/types.ts';
import { withUndoRecord, type OrderedUndoRecorded } from '../model/undo-record.ts';
import { refuseLockedBlocks, siblings } from '../placement.ts';
import { breakPiece, type OrderedBreak, type OrderedBreakClear } from './breaks.ts';
import {
  contextOf,
  type Ctx,
  type OrderedEditContext,
  type OrderedReplacedFormat,
} from './context.ts';
import { setControlValue } from './controls.ts';
import { retargetLink, type OrderedLinkRange, type OrderedLinkTarget } from './links.ts';
import {
  directDelete,
  directDeleteRow,
  directEnter,
  directFormat,
  directFormatMark,
  directInsert,
  directJoin,
  directParagraphOps,
  settle,
} from './direct.ts';
import {
  formatMark,
  joinParagraphs,
  markInPrompt,
  paragraphOp,
  splitParagraph,
  strikeMark,
} from './marks.ts';
import { insertComment } from './comments.ts';
import { copyText, moveText } from './move.ts';
import { insertNote, typeInNewNote } from './notes.ts';
import {
  resolveOrderedPlace,
  shownItemAt,
  type OrderedPlace,
  type OrderedResolvedPlace,
} from './places.ts';
import { retractSuggestion } from './retract.ts';
import {
  columnOp,
  deleteRow,
  insertRow,
  insertTable,
  tableProperty,
  type OrderedTableOp,
} from './tables.ts';
import { deleteText, formatText, insertText } from './text.ts';

/**
 * One edit. Places are addresses (`OrderedPlace`): the unit a caret comes before, or a
 * paragraph mark for the end of its paragraph. A range is `[from, to)`.
 *
 * @public
 */
export type OrderedEdit =
  /** Typed text; `inside` names the inline control a caret on its edge types into. */
  | {
      readonly kind: 'insert';
      readonly at: OrderedPlace;
      readonly text: string;
      readonly inside?: OrderedPlace;
    }
  /** A tab or a break: one object at `at`; a line break may clear floating objects. */
  | {
      readonly kind: 'insertBreak';
      readonly at: OrderedPlace;
      readonly break: OrderedBreak;
      readonly clear?: OrderedBreakClear;
    }
  | { readonly kind: 'delete'; readonly from: OrderedPlace; readonly to: OrderedPlace }
  | {
      readonly kind: 'replace';
      readonly from: OrderedPlace;
      readonly to: OrderedPlace;
      readonly text: string;
    }
  | ({
      readonly kind: 'format';
      readonly from: OrderedPlace;
      readonly to: OrderedPlace;
    } & OrderedFormatWrite)
  /** Run properties on the mark of the paragraph that `at` ends (`at` is the mark's address). */
  | ({ readonly kind: 'formatMark'; readonly at: OrderedPlace } & OrderedFormatWrite)
  /** Enter. */
  | { readonly kind: 'split'; readonly at: OrderedPlace }
  /** Backspace at the start of the paragraph `at` begins (Delete at the end of the one before). */
  | { readonly kind: 'join'; readonly at: OrderedPlace }
  /** Paragraph presses on the paragraph `at` is in; one press is one suggestion. */
  | {
      readonly kind: 'paragraph';
      readonly at: OrderedPlace;
      readonly ops: readonly OrderedParagraphOp[];
    }
  | { readonly kind: 'insertRow'; readonly at: OrderedPlace; readonly where: 'above' | 'below' }
  | { readonly kind: 'deleteRow'; readonly at: OrderedPlace }
  | { readonly kind: 'insertColumn'; readonly at: OrderedPlace; readonly where: 'left' | 'right' }
  | { readonly kind: 'deleteColumn'; readonly at: OrderedPlace }
  /** A new table before the paragraph `at` is in; each column `columnWidth` twips (2000). */
  | {
      readonly kind: 'insertTable';
      readonly at: OrderedPlace;
      readonly rows: number;
      readonly cols: number;
      readonly columnWidth?: number;
    }
  /**
   * A press on the caret's table, row or cell. A cell op covers `cells` (the open items of
   * cells of the caret's table) when given, as one selection, else the caret's cell.
   */
  | {
      readonly kind: 'tableProperty';
      readonly at: OrderedPlace;
      readonly op: OrderedTableOp;
      readonly cells?: readonly OrderedPlace[];
    }
  /** A control's new value; `control` is the address of the control's open item. */
  | {
      readonly kind: 'controlValue';
      readonly control: OrderedPlace;
      readonly value: string | ContentControlValueInput;
    }
  /**
   * The link at `at` (or the link whose open item is `link`) points at an address (`url`) or a
   * bookmark (`anchor`): exactly one. `tooltip` replaces its tooltip; `from` and `to`, both or
   * neither, name a part of its words, which becomes a link of its own.
   */
  | {
      readonly kind: 'linkTarget';
      readonly at: OrderedPlace;
      readonly link?: OrderedPlace;
      readonly url?: string;
      readonly anchor?: string;
      readonly tooltip?: string;
      readonly from?: OrderedPlace;
      readonly to?: OrderedPlace;
    }
  /** Copy the words and objects `[from, to)` of one paragraph to `at`, as an insertion. */
  | {
      readonly kind: 'copy';
      readonly from: OrderedPlace;
      readonly to: OrderedPlace;
      readonly at: OrderedPlace;
    }
  /** Move the words `[from, to)` of one paragraph to `at`. */
  | {
      readonly kind: 'move';
      readonly from: OrderedPlace;
      readonly to: OrderedPlace;
      readonly at: OrderedPlace;
    }
  /** A new note at `at`, its paragraph holding `text` after the note's own mark. */
  | {
      readonly kind: 'insertNote';
      readonly at: OrderedPlace;
      readonly note: 'footnote' | 'endnote';
      readonly text?: string;
    }
  /** Take back the author's own pending suggestion. */
  | { readonly kind: 'retract'; readonly suggestion: string }
  /** A new comment thread of `text` on `[from, to)`, or at a caret when they are one place. */
  | {
      readonly kind: 'comment';
      readonly from: OrderedPlace;
      readonly to: OrderedPlace;
      readonly text: string;
    };

/** What an edit did. @public */
export interface OrderedEditOutcome {
  /** The suggestions that hold the edit: made, extended, or (direct edits) accepted at once. */
  readonly suggestions: readonly string[];
  /** Policy 2: the format records the edit took over from other suggestions. */
  readonly replaced: readonly OrderedReplacedFormat[];
}

/** Whether text holds a C0 control: a tab or a break is an object, never a character. */
function holdsControl(text: string): boolean {
  for (let k = 0; k < text.length; k += 1) if (text.charCodeAt(k) < 0x20) return true;
  return false;
}

function checkText(text: unknown): string {
  if (typeof text !== 'string' || text.length === 0 || holdsControl(text))
    throw new OrderedRefusal('invalid-edit', 'text must be plain characters');
  return text;
}

/** The view of the shown paragraph at an index, as it is now. */
function viewAt(ctx: Ctx, index: number): OrderedViewParagraph {
  const view = orderedViewParagraph(ctx.doc, ctx.story.name, index);
  if (!view) {
    throw new OrderedRefusal('malformed-story', `no paragraph ${index}`);
  }

  return view;
}

/** The views of the shown paragraphs `[from, to]`, read once before an edit writes. */
function viewsBetween(ctx: Ctx, from: number, to: number): OrderedViewParagraph[] {
  const out: OrderedViewParagraph[] = [];
  for (let index = from; index <= to; index += 1) {
    out.push(viewAt(ctx, index));
  }

  return out;
}

function resolve(ctx: Ctx, place: OrderedPlace): OrderedResolvedPlace {
  return resolveOrderedPlace(ctx.doc, ctx.story.name, place);
}

function range(ctx: Ctx, from: OrderedPlace, to: OrderedPlace) {
  const a = resolve(ctx, from);
  const b = resolve(ctx, to);
  if (b.index < a.index || (b.index === a.index && b.offset <= a.offset))
    throw new OrderedRefusal('empty-range');
  return { a, b, views: viewsBetween(ctx, a.index, b.index) };
}

/** A format edit's press: what it sets and what it takes off, at least one property. */
function checkedFormatWrite(edit: OrderedFormatWrite): OrderedFormatWrite {
  const remove = edit.remove ?? [];
  if (
    !Array.isArray(edit.properties) ||
    !Array.isArray(remove) ||
    edit.properties.length + remove.length === 0
  )
    throw new OrderedRefusal('invalid-edit', 'a format needs properties');
  return remove.length === 0
    ? { properties: edit.properties }
    : { properties: edit.properties, remove };
}

/** One paragraph of a range: its view, its part of the range, and the paragraph after it. */
interface RangeStep {
  readonly p: OrderedViewParagraph;
  readonly start: number;
  readonly end: number;
  readonly next: OrderedViewParagraph | undefined;
}

/** A range paragraph by paragraph, each with its offsets, from views read before any write. */
function rangeSteps(
  ctx: Ctx,
  from: OrderedPlace,
  to: OrderedPlace
): { steps: RangeStep[]; start: OrderedResolvedPlace } {
  const { a, b, views } = range(ctx, from, to);
  const steps = views.map(
    (p, k): RangeStep => ({
      p,
      start: k === 0 ? a.offset : 0,
      end: k === views.length - 1 ? b.offset : p.chars.length,
      next: views[k + 1],
    })
  );
  return { steps, start: a };
}

/** The range goes on into `next`, a sibling: the mark between them is in it (read as it is). */
function joinsNext(ctx: Ctx, step: RangeStep): step is RangeStep & { next: OrderedViewParagraph } {
  return step.next !== undefined && siblings(ctx, step.p, step.next);
}

/** A deletion over a range, paragraph by paragraph, and the marks between. */
function deleteOverRange(
  ctx: Ctx,
  from: OrderedPlace,
  to: OrderedPlace
): { suggestion: string | null; start: OrderedResolvedPlace } {
  const walk = rangeSteps(ctx, from, to);
  let id: string | null = null;
  for (const step of walk.steps) {
    const { p, start, end } = step;
    const joins = joinsNext(ctx, step);
    if (ctx.direct) {
      if (end > start) {
        directDelete(ctx, p, start, end);
      }

      if (joins) {
        directJoin(ctx, p, step.next);
      }

      continue;
    }

    if (end > start) {
      id = deleteText(ctx, p, start, end, suggestionOrNull(ctx.doc, id)) ?? id;
    }

    if (joins && !markInPrompt(ctx, p, step.next)) {
      refuseLockedBlocks(ctx, p.blocks);
      id = strikeMark(ctx, p.mark, suggestionOrNull(ctx.doc, id)) ?? id;
    }
  }

  return { suggestion: id, start: walk.start };
}

/** A format over a range, paragraph by paragraph, and the marks between. */
function formatOverRange(
  ctx: Ctx,
  from: OrderedPlace,
  to: OrderedPlace,
  write: OrderedFormatWrite
): string | null {
  let id: string | null = null;
  for (const step of rangeSteps(ctx, from, to).steps) {
    const { p, start, end } = step;
    const joins = joinsNext(ctx, step);
    if (ctx.direct) {
      if (end > start) {
        directFormat(ctx, p, start, end, write);
      }

      if (joins) {
        directFormatMark(ctx, p, write);
      }

      continue;
    }

    if (end > start) {
      id = formatText(ctx, p, start, end, write, suggestionOrNull(ctx.doc, id)) ?? id;
    }

    if (joins) {
      id = formatMark(ctx, p, write, suggestionOrNull(ctx.doc, id)) ?? id;
    }
  }

  return id;
}

/** Settle the suggestion at once when the edit is direct (an edit with no direct form). */
function settledIfDirect(ctx: Ctx, suggestion: string | null): string | null {
  if (ctx.direct) {
    settle(ctx.doc, suggestion);
  }

  return suggestion;
}

/** A link target edit's target: exactly one of an address and a bookmark. */
function linkTargetOf(edit: Extract<OrderedEdit, { kind: 'linkTarget' }>): OrderedLinkTarget {
  if (typeof edit.url === 'string' && edit.anchor === undefined) {
    return { url: edit.url };
  }

  if (typeof edit.anchor === 'string' && edit.url === undefined) {
    return { anchor: edit.anchor };
  }

  throw new OrderedRefusal('invalid-edit', 'a link target is one address or one bookmark');
}

/** The part of a link a link target edit names (`from` and `to`), or null for the whole link. */
function linkPartOf(
  ctx: Ctx,
  edit: Extract<OrderedEdit, { kind: 'linkTarget' }>
): OrderedLinkRange | null {
  if (edit.from === undefined || edit.to === undefined) {
    return null;
  }

  const { a, b } = range(ctx, edit.from, edit.to);
  if (a.index !== b.index) {
    throw new OrderedRefusal('not-in-link', 'a link is in one paragraph');
  }

  return { p: a.paragraph, start: a.offset, end: b.offset };
}

function run(ctx: Ctx, edit: OrderedEdit): (string | null)[] {
  switch (edit.kind) {
    case 'insert': {
      const text = checkText(edit.text);
      const at = resolve(ctx, edit.at);
      const inside =
        edit.inside === undefined
          ? {}
          : { inside: shownItemAt(ctx.doc, ctx.story, edit.inside, 'open', ['sdt']) };
      if (ctx.direct) {
        directInsert(ctx, at.paragraph, at.offset, text, inside);
        return [];
      }

      return [insertText(ctx, at.paragraph, at.offset, text, inside)];
    }
    case 'insertBreak': {
      const piece = breakPiece(ctx, edit.break, edit.clear);
      const at = resolve(ctx, edit.at);
      if (ctx.direct) {
        directInsert(ctx, at.paragraph, at.offset, '', { pieces: [piece] });
        return [];
      }

      const options = { pieces: [piece], outsidePrompt: true };
      return [insertText(ctx, at.paragraph, at.offset, '', options)];
    }
    case 'delete':
      return [deleteOverRange(ctx, edit.from, edit.to).suggestion];
    case 'replace': {
      const text = checkText(edit.text);
      const { suggestion, start } = deleteOverRange(ctx, edit.from, edit.to);

      // The text goes where the range started, in the view as the deletion left it
      const p = viewAt(ctx, start.index);
      if (ctx.direct) {
        directInsert(ctx, p, start.offset, text);
        return [];
      }

      // The engine's replacement is two transactions: the typing goes before the struck words
      const joined = suggestion === null ? ctx : { ...ctx, join: suggestion, earlierStrike: true };
      return [suggestion, insertText(joined, p, start.offset, text)];
    }
    case 'format': {
      const write = checkedFormatWrite(edit);
      return [formatOverRange(ctx, edit.from, edit.to, write)];
    }
    case 'formatMark': {
      const write = checkedFormatWrite(edit);
      const at = resolve(ctx, edit.at);
      const mark = itemAt(ctx.story, edit.at);
      if (!mark || mark.kind !== 'mark' || at.paragraph.mark !== mark)
        throw new OrderedRefusal('hidden-target', 'a mark format names a shown mark');
      if (ctx.direct) {
        directFormatMark(ctx, at.paragraph, write);
        return [];
      }

      return [formatMark(ctx, at.paragraph, write, null)];
    }
    case 'split': {
      const at = resolve(ctx, edit.at);
      if (ctx.direct) {
        directEnter(ctx, at.paragraph, at.offset);
        return [];
      }

      return [splitParagraph(ctx, at.paragraph, at.offset)];
    }
    case 'join': {
      const at = resolve(ctx, edit.at);
      if (at.offset !== 0) throw new OrderedRefusal('not-paragraph-start');
      if (at.index === 0) throw new OrderedRefusal('no-paragraph-before');
      const previous = viewAt(ctx, at.index - 1);
      if (ctx.direct) {
        directJoin(ctx, previous, at.paragraph);
        return [];
      }

      return [joinParagraphs(ctx, previous, at.paragraph)];
    }
    case 'paragraph': {
      if (!Array.isArray(edit.ops) || edit.ops.length === 0)
        throw new OrderedRefusal('invalid-edit', 'a press needs ops');
      const at = resolve(ctx, edit.at);
      if (ctx.direct) {
        directParagraphOps(ctx, at.paragraph, edit.ops);
        return [];
      }

      let id: string | null = null;
      for (const op of edit.ops) {
        id = paragraphOp(ctx, at.paragraph, op, suggestionOrNull(ctx.doc, id));
      }

      return [id];
    }
    case 'insertRow':
      return [settledIfDirect(ctx, insertRow(ctx, resolve(ctx, edit.at).paragraph, edit.where))];
    case 'deleteRow': {
      const p = resolve(ctx, edit.at).paragraph;
      if (ctx.direct) {
        directDeleteRow(ctx, p);
        return [];
      }

      return [deleteRow(ctx, p)];
    }
    case 'insertColumn':
      return [
        settledIfDirect(ctx, columnOp(ctx, resolve(ctx, edit.at).paragraph, 'insert', edit.where)),
      ];
    case 'deleteColumn':
      return [settledIfDirect(ctx, columnOp(ctx, resolve(ctx, edit.at).paragraph, 'delete'))];
    case 'insertTable': {
      const p = resolve(ctx, edit.at).paragraph;
      return [settledIfDirect(ctx, insertTable(ctx, p, edit.rows, edit.cols, edit.columnWidth))];
    }
    case 'tableProperty': {
      const p = resolve(ctx, edit.at).paragraph;
      const cells = edit.cells?.map((place) => itemAt(ctx.story, place));
      if (cells?.some((cell) => !cell))
        throw new OrderedRefusal('unknown-address', 'a cell of the press is not in the story');
      return [
        settledIfDirect(
          ctx,
          tableProperty(
            ctx,
            p,
            edit.op,
            cells?.filter((c) => c !== null)
          )
        ),
      ];
    }
    case 'controlValue': {
      if (typeof edit.value !== 'string' && (typeof edit.value !== 'object' || !edit.value))
        throw new OrderedRefusal('invalid-edit', 'a value is text or a control value');
      const open = itemAt(ctx.story, edit.control);
      if (!open)
        throw new OrderedRefusal('unknown-control', `${edit.control.space}|${edit.control.offset}`);
      return [settledIfDirect(ctx, setControlValue(ctx, open, edit.value))];
    }
    case 'linkTarget': {
      const target = linkTargetOf(edit);
      if ((edit.from === undefined) !== (edit.to === undefined)) {
        throw new OrderedRefusal('invalid-edit', 'a part of a link names both of from and to');
      }

      const open = edit.link === undefined ? linkAt(ctx, edit.at) : itemAt(ctx.story, edit.link);
      if (!open || open.kind !== 'open' || open.token !== 'hyperlink')
        throw new OrderedRefusal('not-in-link', 'the place is not in a link');
      return [retargetLink(ctx, open, target, edit.tooltip, linkPartOf(ctx, edit))];
    }
    case 'move': {
      const { a, b } = range(ctx, edit.from, edit.to);
      if (a.index !== b.index)
        throw new OrderedRefusal('invalid-edit', 'a move of words in one paragraph');
      const at = resolve(ctx, edit.at);
      return [
        settledIfDirect(
          ctx,
          moveText(ctx, a.paragraph, a.offset, b.offset, at.paragraph, at.offset)
        ),
      ];
    }
    case 'copy': {
      const { a, b } = range(ctx, edit.from, edit.to);
      if (a.index !== b.index)
        throw new OrderedRefusal('invalid-edit', 'a copy of words in one paragraph');
      const at = resolve(ctx, edit.at);
      return [
        settledIfDirect(
          ctx,
          copyText(ctx, a.paragraph, a.offset, b.offset, at.paragraph, at.offset)
        ),
      ];
    }
    case 'insertNote': {
      const text = edit.text ?? '';
      if (text.length > 0) checkText(text);
      const at = resolve(ctx, edit.at);
      const sid = insertNote(ctx, at.paragraph, at.offset, edit.note);
      if (text.length > 0) typeInNewNote(ctx, sid, edit.note, text);
      return [settledIfDirect(ctx, sid)];
    }
    case 'retract':
      retractSuggestion(ctx.doc, edit.suggestion, ctx.attribution);
      return [edit.suggestion];
    case 'comment':
      insertComment(ctx, resolve(ctx, edit.from), resolve(ctx, edit.to), edit.text);
      return [];
    default:
      return notCovered(edit);
  }
}

/** The innermost link around the view character at a place. */
function linkAt(ctx: Ctx, place: OrderedPlace) {
  const at = resolve(ctx, place);
  const char = at.paragraph.chars[at.offset] ?? at.paragraph.chars[at.offset - 1];
  const segment = char ? segmentsOf(char.containers).reverse().find(isLinkSegment) : undefined;
  return segment === undefined ? null : openOfSegment(ctx.story, segment);
}

/** An edit of a kind the model does not know (a new picture, a crop, a rich paste). */
function notCovered(edit: never): never {
  const kind = (edit as { readonly kind?: unknown }).kind;
  throw new OrderedRefusal(
    'not-covered',
    `an edit of kind ${JSON.stringify(kind ?? null).slice(0, 64)}`
  );
}

/**
 * Apply one edit inside the undo record the caller opened: a change made of several edits
 * applies them in one record, so it holds whole or not at all. `onReplaced` is not called: the
 * outcome lists the takeovers, for the caller to report once its record holds.
 */
export function runOrderedEdit(context: OrderedEditContext, edit: OrderedEdit): OrderedEditOutcome {
  if (context.doc.undo === null) throw new Error('An edit runs inside an undo record.');
  return editOutcome(context, (ctx) => run(ctx, edit));
}

/**
 * Run an edit on its context: the suggestions it names and every suggestion it added, and the
 * records it took over.
 */
export function editOutcome(
  context: OrderedEditContext,
  edit: (ctx: Ctx) => readonly (string | null)[]
): OrderedEditOutcome {
  const replaced: OrderedReplacedFormat[] = [];
  const ctx = contextOf(context, replaced);
  const before = new Set(context.doc.suggestions.keys());
  const ids = edit(ctx).filter((id): id is string => id !== null);
  for (const id of context.doc.suggestions.keys()) if (!before.has(id)) ids.push(id);
  return { suggestions: [...new Set(ids)], replaced };
}

/**
 * Apply one edit to the context's story, inside a new undo record. It holds whole, or it is
 * refused (`OrderedRefusal`) and nothing changed. Records do not nest: none may be open.
 *
 * @public
 */
export function applyOrderedEdit(
  context: OrderedEditContext,
  edit: OrderedEdit
): OrderedUndoRecorded<OrderedEditOutcome> {
  const recorded = withUndoRecord(context.doc, () => runOrderedEdit(context, edit));
  for (const takeover of recorded.value.replaced) context.onReplaced?.(takeover);
  return recorded;
}
