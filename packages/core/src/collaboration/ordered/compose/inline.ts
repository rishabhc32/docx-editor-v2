// A paragraph's inline items back to runs, containers and revision wrappers. Pending
// suggestions are drawn as Word draws them (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` with
// their range markers, `w:rPrChange`); decided ones are applied.

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { withDeletedTextRestored } from '../../../store/store/revision-rebuild-helpers.ts';
import {
  recordStatesText,
  storesValueRecord,
} from '../../../store/store/tracked-content-control-value.ts';
import { build, textNode } from '../../../store/store/tree-op-tracked.ts';
import { MOVE_RANGE } from '../load/moves.ts';
import { isSplitEdge, movedAwayBy, statusOf } from '../model/document.ts';
import { fileMark, REVISION_ROLE } from '../model/revision-ids.ts';
import {
  isUnit,
  type OrderedFieldPiece,
  type OrderedFormatRecord,
  type OrderedItem,
  type OrderedPropertyRecord,
} from '../model/types.ts';
import {
  attributeOf,
  childWhere,
  fieldCharBalance,
  isRunProperties,
  isWml,
  renamed,
  withChildren,
  withId,
} from '../ooxml-nodes.ts';
import type { ComposeContext } from './context.ts';
import { isNotesStory } from './note-references.ts';
import { containerPropertiesAfter, drawnFormats, faceAfter } from './properties.ts';
import { keptLinkEdges } from './kept-links.ts';
import { hidePrompt } from './prompts.ts';
import { withTextBoxContent } from './text-boxes.ts';
import { objectNodeOf } from './objects.ts';

interface Box {
  readonly open: OrderedItem | null;
  readonly children: OoxmlNode[];
  /** A unit sits right in it, shown or not. */
  hadUnits: boolean;
  /** A unit shows in it (struck or not): markers alone are no content. */
  drawn: boolean;
  /** A decision changed whether a unit in it reads as text. */
  reached: boolean;
  /** A decision on a suggestion took a unit in it away (`ComposeContext.decidedAway`). */
  decided: boolean;
  /** A unit in it shows struck by a pending suggestion. */
  struck: boolean;
  /** A unit in it shows, neither struck nor a suggestion's insertion. */
  own: boolean;
  /** A unit in it shows as a suggestion's insertion an accept decided. */
  accepted: boolean;
  /** The words of insertions a reject took away in it. */
  rejected: string;
  /** The pending insertion of a new control, drawn as one `w:ins` around the control. */
  cover?: string;
}

interface RunBuild {
  readonly key: string;
  /** The room run it is drawn from (`roomRunOf`), which numbers it for a checkpoint; else empty. */
  readonly roomRun: string;
  readonly item: OrderedItem;
  readonly content: OoxmlNode[];
  readonly deleted: boolean;
  text: string;
  textShell: OoxmlElement | undefined;
  /** The text form's item is drawn in its own paragraph: the form keeps its id. */
  textHomed: boolean;
  /** Its first unit, which names who typed every unit of the run. */
  unit?: OrderedItem;
}

/** Whether an item is a unit: text or an atom. */

/** Struck run content: `w:t` as `w:delText`, `w:instrText` as `w:delInstrText`. */
function withTextDeleted(node: OoxmlNode, depth = 0): OoxmlNode {
  if (node.kind === 'textValue' || depth > 32) return node;
  const children = node.children.map((child) => withTextDeleted(child, depth + 1));
  const rebuilt = children.some((c, i) => c !== node.children[i])
    ? withChildren(node, children)
    : node;
  if (rebuilt.kind === 'text' && rebuilt.localName === 't')
    return renamed(rebuilt, 'deletedText', 'delText');
  if (rebuilt.namespaceUri === WML_NAMESPACE_URI && rebuilt.localName === 'instrText')
    return renamed(rebuilt, 'generic', 'delInstrText');
  return rebuilt;
}

/** Whether a field's run holds a `w:fldChar` of this type. */
function hasFieldChar(node: OoxmlNode, type: string, depth = 0): boolean {
  if (node.kind === 'textValue' || depth > 8) return false;
  if (
    node.localName === 'fldChar' &&
    node.attributes.some((a) => a.localName === 'fldCharType' && a.value === type)
  )
    return true;
  return node.children.some((c) => hasFieldChar(c, type, depth + 1));
}

/** The run content that makes a run part of a field's shown result, as the engine counts it. */
const RESULT_CONTENT = new Set(['text', 'deletedText', 'tab', 'hardBreak', 'textValue']);

/**
 * Which pieces of a whole complex field an edit's format reaches, as the engine's
 * `formatRunIds`: the runs of the outer field's result that hold a text, a tab or a break, else
 * the run of its own separator, else its begin run. The runs of the fields nested in it never
 * take it.
 */
export function formatRunsOf(pieces: readonly OrderedFieldPiece[]): boolean[] {
  const formatted = pieces.map(() => false);
  let depth = 0;
  let separate = -1;
  for (const [k, piece] of pieces.entries()) {
    if (piece.node.kind !== 'run') {
      continue;
    }

    for (const child of piece.node.children) {
      depth += fieldCharBalance(child);
      if (depth !== 1) {
        continue;
      }

      if (separate < 0 && hasFieldChar(child, 'separate')) {
        separate = k;
      } else if (separate >= 0 && RESULT_CONTENT.has(child.kind)) {
        formatted[k] = true;
      }
    }
  }

  if (!formatted.includes(true)) {
    formatted[Math.max(separate, 0)] = true;
  }

  return formatted;
}

/** A box for the paragraph (`open` null) or for an inline container. */
function newBox(open: OrderedItem | null, cover: string | null = null): Box {
  return {
    open,
    children: [],
    hadUnits: false,
    drawn: false,
    reached: false,
    decided: false,
    struck: false,
    own: false,
    accepted: false,
    rejected: '',
    ...(cover === null ? {} : { cover }),
  };
}

/** The markers between a field's pieces, by field: they draw inside the field while it shows. */
function fieldPassengers(buffer: readonly OrderedItem[]): Map<string, OrderedItem[]> {
  const passengers = new Map<string, OrderedItem[]>();
  for (const item of buffer) {
    if (!item.fieldSlot) continue;
    const list = passengers.get(item.fieldSlot.field);
    if (list) {
      list.push(item);
    } else {
      passengers.set(item.fieldSlot.field, [item]);
    }
  }

  return passengers;
}

/** A move range marker: its range (its half and `w:id`) and whether it starts it. */
function moveRangeEdge(
  item: OrderedItem
): { readonly range: string; readonly start: boolean } | null {
  const node = item.moveRange ? item.node : undefined;
  if (!node || node.kind === 'textValue') {
    return null;
  }

  const match = MOVE_RANGE.exec(node.localName);
  if (!match) {
    return null;
  }

  return { range: `${match[1]}:${attributeOf(node, 'id') ?? ''}`, start: match[2] === 'Start' };
}

/**
 * The markers of a move that are passed over: where two ranges of one move meet with nothing
 * drawn between them, as when a decision took out the words that split the move, and around a
 * range with no moved word drawn in it, as when its words all went (a prompt shown again in their
 * place is no moved word). Every pending move has range marker items (`model/move-ranges.ts`).
 */
function emptyMoveEdges(ctx: ComposeContext, buffer: readonly OrderedItem[]): Set<OrderedItem> {
  const empty = new Set<OrderedItem>();
  if (ctx.doc.movesWithRanges.size === 0) {
    return empty;
  }

  const drawn = buffer.filter((item) => item.kind !== 'mark' && !ctx.gone(item));

  // A range's end with its next start next to it: one range goes on
  let before: OrderedItem | null = null;
  for (const item of drawn) {
    const edge = moveRangeEdge(item);
    const previous = before === null ? null : moveRangeEdge(before);
    const meets = edge?.start === true && previous?.range === edge.range && !previous.start;
    if (before !== null && meets) {
      empty.add(before).add(item);
    }

    before = item;
  }

  // The start of each range open here, while no moved word is drawn after it
  const wordless = new Map<string, OrderedItem>();
  for (const item of drawn) {
    const edge = empty.has(item) ? null : moveRangeEdge(item);
    if (edge === null) {
      if (item.move !== null) {
        wordless.clear();
      }
    } else if (edge.start) {
      wordless.set(edge.range, item);
    } else {
      const start = wordless.get(edge.range);
      if (start !== undefined) {
        empty.add(start).add(item);
      }

      wordless.delete(edge.range);
    }
  }

  return empty;
}

/** A unit in the boxes: the innermost held it; each one it shows in, reached or struck. */
function noteUnit(ctx: ComposeContext, boxes: readonly Box[], item: OrderedItem): void {
  const innermost = boxes[boxes.length - 1];
  if (!innermost) {
    return;
  }

  innermost.hadUnits = true;

  // A decided insertion or deletion reaches every control around it
  if (ctx.decided(item)) for (const b of boxes) b.reached = true;
  if (ctx.decidedAway(item)) for (const b of boxes) b.decided = true;
  const rejected = ctx.rejectedWords(item);
  for (const b of boxes) {
    b.rejected += rejected;
  }

  if (ctx.gone(item)) return;
  const struck = ctx.struckBy(item) !== null;
  const accepted = ctx.acceptedInsertion(item);
  const own = ctx.ownWord(item);
  for (const b of boxes) {
    b.drawn = true;
    if (struck) b.struck = true;
    b.accepted ||= accepted;
    b.own ||= own;
  }
}

/** Draw a paragraph's inline items: its children. */
export function composeInline(ctx: ComposeContext, buffer: readonly OrderedItem[]): OoxmlNode[] {
  const boxes: Box[] = [newBox(null)];
  const box = (): Box => {
    const top = boxes[boxes.length - 1];
    if (!top) throw new RangeError('An inline drawing without a box.');
    return top;
  };
  let current: RunBuild | null = null;
  const passengers = fieldPassengers(buffer);
  const carried = new Set<OrderedItem>();
  const empty = emptyMoveEdges(ctx, buffer);

  // A rejected new link that still holds something stays, as in Word
  const kept = keptLinkEdges(ctx, buffer);
  const hidden = (item: OrderedItem): boolean => ctx.gone(item) && !kept.has(item);
  const flushText = (build: RunBuild): void => {
    if (!build.text) return;
    const text = textElement(ctx, build.text, build.textShell, build.deleted, build.textHomed);
    build.content.push(ctx.ofRoomText(text, build.roomRun, build.textShell));
    build.text = '';
  };
  const flush = (): void => {
    if (!current) return;
    flushText(current);
    const run = ctx.ofRoomRun(runNode(ctx, current.item, current.content), current.roomRun);
    place(ctx, box().children, ctx.typed(run, current.unit), current.item);
    current = null;
  };
  for (const item of buffer) {
    if (item.kind === 'mark' || carried.has(item)) continue; // a joined paragraph's mark
    if (item.kind === 'open') {
      if (hidden(item)) continue;
      flush();
      const cover = coverOf(ctx, item);
      if (cover !== null) ctx.covered.add(cover);
      boxes.push(newBox(item, cover));
      continue;
    }

    if (item.kind === 'close') {
      if (hidden(item)) continue;
      flush();
      const done = boxes.pop();
      if (!done?.open) continue;
      closeBox(ctx, done, box());
      continue;
    }

    if (isUnit(item)) noteUnit(ctx, boxes, item);
    if (ctx.gone(item) || empty.has(item)) {
      continue;
    }

    if (item.kind === 'zero' && !item.inRun) {
      flush();
      if (item.node && item.moveRange) {
        // A range start says who moved, as the move's wrappers do; its end has only a `w:id`
        const by = moveRangeEdge(item)?.start === true ? item.by : null;
        const node = item.node;
        box().children.push(
          by === null || node.kind === 'textValue' ? node : fileMark(ctx.doc, node, by)
        );
      } else if (item.node) {
        place(ctx, box().children, decidedRunShell(ctx, item, item.node), item);
      }

      continue;
    }

    if (item.bare) {
      flush();
      drawBare(ctx, item, box().children);
      continue;
    }

    if (item.pieces) {
      flush();
      const riding = passengers.get(`${item.space}|${item.offset}`) ?? [];
      for (const passenger of riding) carried.add(passenger);
      drawField(ctx, item, box().children, riding);
      continue;
    }

    const key = runKey(ctx, item);
    if (!current || current.key !== key) {
      flush();
      current = {
        key,
        // Only a checkpoint numbers the room's runs
        roomRun: ctx.options.runs === undefined ? '' : roomRunOf(ctx, item),
        item,
        content: [],
        deleted: struckText(ctx, item),
        text: '',
        textShell: undefined,
        textHomed: false,
      };
    }

    if (current.unit === undefined && isUnit(item)) {
      current.unit = item;
    }

    if (item.kind === 'text') {
      const homed = ctx.homed(item);
      if (
        current.text &&
        (current.textShell?.id !== item.textShell?.id || current.textHomed !== homed)
      )
        flushText(current);
      current.text += item.text;
      current.textShell = item.textShell;
      current.textHomed = homed;
    } else {
      flushText(current);

      // The object under its records, its text boxes holding their stories as they are now
      const node = withTextBoxContent(ctx, item, objectNodeOf(ctx, item));
      if (node) current.content.push(node);
    }
  }

  flush();

  // A container still open at the paragraph's end closes there, as at its close item
  while (boxes.length > 1) {
    const done = boxes.pop();
    if (done) closeBox(ctx, done, box());
  }

  return box().children;
}

/** The `w:sdtPr` keys a change of a control's title, tag or lock sets: no value. */
const METADATA_KEYS = new Set(
  ['alias', 'tag', 'lock'].map((name) => `${WML_NAMESPACE_URI} ${name}`)
);

/**
 * A record of a control's value (a list's last value, a date, a check box), not only of its
 * title, tag or lock: the value records follow the text the control shows. A record of its words
 * alone sets none: a value proposed over a prompt keeps the record the control stores.
 */
function isValueRecord(record: OrderedPropertyRecord): boolean {
  if (record.target !== 'sdtPr') return false;
  const merge = record.kind === 'element' ? record.merge : undefined;
  return merge === undefined || merge.some((key) => !METADATA_KEYS.has(key));
}

/**
 * The pending insertion a new inline control is drawn inside, as one `w:ins` around the whole
 * control (the engine's tracked control and Word's `w:ins` around `w:sdt`), so a reject takes
 * the control with its words. Not for a control an Enter split: a decision keeps both halves.
 */
export function coverOf(ctx: ComposeContext, open: OrderedItem): string | null {
  if (open.token !== 'sdt' || isSplitEdge(open) || open.by === null) return null;
  if (ctx.covered.has(open.by)) return null;
  return ctx.insertedBy(open);
}

/** What the units of a control told its value record: an inline control's box, a block's builder. */
export interface ValueSigns {
  /** A decision changed whether a unit in it reads as text. */
  readonly reached: boolean;
  /** A unit in it shows struck by a pending suggestion. */
  readonly struck: boolean;
  /** A unit in it shows, neither struck nor a suggestion's insertion. */
  readonly own: boolean;
  /** A unit in it shows as a suggestion's insertion an accept decided. */
  readonly accepted: boolean;
  /** The words of insertions a reject took away in it. */
  readonly rejected: string;
}

/**
 * Keep a control's value record in step with what it shows (`ComposeContext.reached`), as the
 * engine reads it back after a decision: a decision reached it or emptied it; it has a value
 * record that shows and no pending deletion (a record taken back, a direct edit undone, is
 * none); or words that filled it beside a record it stores were accepted
 * (`newWordsBesideRecord`). A record the words a reject took away stated is stale
 * (`ComposeContext.stale`). `control` is the control as drawn.
 */
export function noteValueRecord(
  ctx: ComposeContext,
  open: OrderedItem,
  control: OoxmlElement,
  signs: ValueSigns,
  emptied: boolean
): void {
  const valued =
    !signs.struck &&
    open.props.some(
      (record) => isValueRecord(record) && statusOf(ctx.doc, record.suggestion) !== 'rejected'
    );
  const filled = signs.accepted && !signs.own && storesValueRecord(control);
  if (signs.rejected !== '' && recordStatesText(control, signs.rejected)) {
    ctx.stale.add(control.id);
  }

  if (signs.reached || emptied || valued || filled) {
    ctx.reached.push(control.id);
  }
}

function closeBox(ctx: ComposeContext, done: Box, into: Box): void {
  const open = done.open;
  if (done.cover !== undefined) ctx.covered.delete(done.cover);
  if (!open) return;

  // A link whose every unit is gone goes too (the engine drops it); a control, a smart tag or
  // custom XML stays, and a control shows its prompt again
  if (done.children.length === 0 && done.hadUnits && open.token === 'hyperlink') return;
  const container = inlineContainer(ctx, open, done.children);
  const emptied = !done.drawn && done.hadUnits;
  if (open.token === 'sdt') {
    if (emptied) ctx.emptied.push(container.id);
    if (emptied && done.decided) ctx.decisionEmptied.add(container.id);

    noteValueRecord(ctx, open, container, done, emptied);
  }

  if (done.cover !== undefined) place(ctx, into.children, container, open);
  else into.children.push(container);
}

function textElement(
  ctx: ComposeContext,
  text: string,
  shell: OoxmlElement | undefined,
  deleted: boolean,
  homed: boolean
): OoxmlNode {
  if (!shell) return textNode(ctx.mint, text, deleted);
  const form = deleted ? withTextDeleted(shell) : shell;
  if (form.kind === 'textValue') return textNode(ctx.mint, text, deleted);
  const id = !homed || ctx.placed.has(form.id) ? ctx.mint() : form.id;
  ctx.placed.add(id);
  return withId(withChildren(form, [{ id: ctx.mint(), kind: 'textValue', value: text }]), id);
}

/** A drawn format record as a run's key names it. */
const recordKey = (record: OrderedFormatRecord): string => `${record.id}:${record.suggestion}`;

/**
 * The run the room draws a unit in, by what no file holds: its `run`, its face and the accepted
 * format records it draws. A pending record splits the run only while it shows, and the file
 * draws that split, so it names no room run: a take-back or a reject draws the run whole again. An
 * accepted record splits it for good, as the editor keeps the runs a format split, and a fold
 * writes it into the face: it names the room run.
 */
function roomRunOf(ctx: ComposeContext, unit: OrderedItem): string {
  const accepted = drawnFormats(ctx, unit.fmts).filter(
    (record) => statusOf(ctx.doc, record.suggestion) === 'accepted'
  );
  return [unit.run, unit.face?.id ?? '', ...accepted.map(recordKey)].join(' ');
}

function runKey(ctx: ComposeContext, item: OrderedItem): string {
  const by = ctx.insertedBy(item) ?? '';
  const del = ctx.struckBy(item) ?? '';

  // A move's name parts runs only inside its revision (as `place` draws it): a decided move's
  // words join their run, struck or not
  const moved = by !== '' || (del !== '' && movedAwayBy(ctx.doc, item, del));
  const move = moved ? (item.move ?? '') : '';
  return [
    item.run,
    item.face ? item.face.id : '',
    drawnFormats(ctx, item.fmts).map(recordKey).join(','),
    by,
    del,
    move,
    item.instr ? 'instr' : '',
  ].join('|');
}

function runNode(ctx: ComposeContext, item: OrderedItem, content: OoxmlNode[]): OoxmlElement {
  const shell = ctx.doc.runShells.get(item.run);
  const homed = ctx.homed(item);
  const rPr = ctx.fresh(faceAfter(ctx, item.fmts, item.face), homed);
  const id = ctx.runId(homed ? shell?.id : undefined);
  const run = withChildren(shell ?? build(id, 'run', 'r', [], []), [
    ...(rPr ? [rPr] : []),
    ...content,
  ]);
  return withId(run, id);
}

/** An object in the paragraph itself: an equation, a simple field with its runs. */
function drawBare(ctx: ComposeContext, item: OrderedItem, into: OoxmlNode[]): void {
  const node = item.node;
  if (!node) return;
  if (node.kind !== 'textValue' && node.localName === 'fldSimple') {
    // A simple field: its revision goes around the runs in it (a `w:del` cannot hold the field
    // itself), and a format reaches those runs
    const inner: OoxmlNode[] = [];
    for (const piece of item.pieces ?? []) {
      if (piece.node.kind === 'run')
        place(ctx, inner, fieldRun(ctx, item, piece, piece.node, true), item);
      else inner.push(piece.node);
    }

    into.push(ctx.typed(withChildren(node, inner), item));
    return;
  }

  place(ctx, into, node.kind === 'textValue' ? node : ctx.typed(node, item), item);
}

/**
 * A whole complex field: its own runs as the file had them, an edit's format on the runs
 * `formatRunsOf` names. The markers between its pieces draw in their places.
 */
function drawField(
  ctx: ComposeContext,
  item: OrderedItem,
  into: OoxmlNode[],
  passengers: readonly OrderedItem[]
): void {
  const pieces = item.pieces ?? [];
  const ride = (before: (slot: number) => boolean): void => {
    for (const passenger of passengers) {
      const slot = passenger.fieldSlot?.before ?? 0;
      if (before(slot) && passenger.node && !ctx.gone(passenger))
        place(ctx, into, passenger.node, passenger);
    }
  };
  const formatted = item.fmts.some((record) => record.kind === 'properties')
    ? formatRunsOf(pieces)
    : [];
  for (const [k, piece] of pieces.entries()) {
    ride((slot) => slot === k);
    let node = piece.node;
    if (node.kind === 'run') {
      node = ctx.typed(withId(node, ctx.runId(node.id)), item);
    }

    place(ctx, into, fieldRun(ctx, item, piece, node, formatted[k] === true), item);
  }

  ride((slot) => slot >= pieces.length);
}

/**
 * A field's run as the field's state draws it: struck text as deleted text, and formats. A
 * file's format record reaches the runs that had it, each with its own face; an edit's format
 * reaches the run when it is `formatted`.
 */
function fieldRun(
  ctx: ComposeContext,
  item: OrderedItem,
  piece: OrderedFieldPiece,
  run: OoxmlNode,
  formatted: boolean
): OoxmlNode {
  if (run.kind !== 'run') return run;
  const faces = piece.recordFaces;
  const files = piece.recordFiles;
  const fmts: OrderedFormatRecord[] = [];
  for (const record of item.fmts) {
    if (record.kind === 'properties') {
      if (formatted) {
        fmts.push(record);
      }

      continue;
    }

    if (!faces?.has(record.id)) continue;

    // The run's own face under the record, and its own change record
    const { fileRecord: _shared, ...rest } = record;
    const own = files?.get(record.id);
    fmts.push({ ...rest, face: faces.get(record.id) ?? null, ...(own ? { fileRecord: own } : {}) });
  }

  let node: OoxmlNode = run;
  if (fmts.length > 0) {
    const base = childWhere(run, isRunProperties);
    const face = ctx.fresh(faceAfter(ctx, fmts, base ?? null));
    node = withChildren(run, [...(face ? [face] : []), ...run.children.filter((c) => c !== base)]);
  }

  return struckText(ctx, item) ? withTextDeleted(node) : withDeletedTextRestored(node);
}

/**
 * Whether the item's text is written in its deleted form (`w:delText`). Moved-from text keeps
 * its plain form inside `w:moveFrom`, as Word writes a move.
 */
function struckText(ctx: ComposeContext, item: OrderedItem): boolean {
  const del = ctx.struckBy(item);
  return del !== null && !movedAwayBy(ctx.doc, item, del);
}

/**
 * A run with nothing in it but its properties, kept whole, once a decision reached the file's
 * change of its properties: its properties as the decision leaves them (the new ones for an
 * accept, the old ones for a reject). While the change is pending the run stays as the file
 * wrote it.
 */
function decidedRunShell(ctx: ComposeContext, item: OrderedItem, node: OoxmlNode): OoxmlNode {
  if (
    node.kind === 'textValue' ||
    item.fmts.every((record) => statusOf(ctx.doc, record.suggestion) === 'pending')
  )
    return node;
  const rPr = faceAfter(ctx, item.fmts, item.face);
  const children = node.children.filter((child) => !isRunProperties(child));
  return withChildren(node, rPr && rPr.children.length > 0 ? [rPr, ...children] : children);
}

/**
 * A node into its container, wrapped in the revisions over it. Neighbours under one outer
 * wrapper share it, and their inner wrappers of one revision join too, as Word writes them.
 */
function place(ctx: ComposeContext, into: OoxmlNode[], node: OoxmlNode, item: OrderedItem): void {
  const by = ctx.insertedBy(item);
  const del = ctx.struckBy(item);
  let wrapped = node;
  let key = '';
  let outer: string | null = null;
  if (del) {
    // Moved-from text is a pending move's source; a deletion of moved-to text, or of a decided
    // move's words, is a plain deletion
    const moved = movedAwayBy(ctx.doc, item, del);
    const kind = moved ? 'moveFrom' : 'del';
    const role = moved
      ? REVISION_ROLE.moveFromIn(ctx.paragraphKey)
      : REVISION_ROLE.deleteIn(ctx.paragraphKey);
    const id = ctx.revisionId(del, role);
    wrapped = wrapper(ctx, kind, del, role, [wrapped]);
    key = `${kind}:${id}:${del}`;
    outer = key;
    ctx.wrappers.set(wrapped, { key, outer: key });
  }

  if (by) {
    const kind = item.move !== null ? 'moveTo' : 'ins';
    const role = insertionRole(ctx, item);
    const id = ctx.revisionId(by, role);
    wrapped = wrapper(ctx, kind, by, role, [wrapped]);
    key = `${kind}:${id}:${by}|${key}`;
    outer = `${kind}:${id}:${by}`;
  }

  if (outer === null || wrapped.kind === 'textValue') {
    into.push(wrapped);
    return;
  }

  const info = { key, outer };
  const last = into[into.length - 1];
  if (last && last.kind !== 'textValue' && ctx.wrappers.get(last)?.outer === outer) {
    const joined = joinWrappers(ctx, last, wrapped);
    ctx.wrappers.set(joined, info);
    into[into.length - 1] = joined;
    return;
  }

  ctx.wrappers.set(wrapped, info);
  into.push(wrapped);
}

/**
 * The role of an insertion's wrapper: a move's, or an insertion's, which keeps one revision id
 * across paragraphs (a notes part's its own), where struck text takes one per paragraph.
 */
function insertionRole(ctx: ComposeContext, item: OrderedItem): string {
  if (item.move !== null) return REVISION_ROLE.moveTo;
  return isNotesStory(ctx.story) ? REVISION_ROLE.noteInsert : REVISION_ROLE.insert;
}

/**
 * Two wrappers of one outer revision as one; an inner wrapper of one revision at the seam joins
 * too.
 */
function joinWrappers(ctx: ComposeContext, left: OoxmlElement, right: OoxmlElement): OoxmlElement {
  const lastInner = left.children[left.children.length - 1];
  const [firstInner, ...rest] = right.children;
  if (
    lastInner &&
    firstInner &&
    lastInner.kind !== 'textValue' &&
    firstInner.kind !== 'textValue' &&
    lastInner !== left &&
    ctx.wrappers.get(lastInner) !== undefined &&
    ctx.wrappers.get(lastInner)?.key === ctx.wrappers.get(firstInner)?.key
  ) {
    const inner = withChildren(lastInner, [...lastInner.children, ...firstInner.children]);
    const innerInfo = ctx.wrappers.get(lastInner);
    if (innerInfo) ctx.wrappers.set(inner, innerInfo);
    return withChildren(left, [...left.children.slice(0, -1), inner, ...rest]);
  }

  return withChildren(left, [...left.children, ...right.children]);
}

const WRAPPER_KINDS = {
  ins: 'revisionInsert',
  del: 'revisionDelete',
  moveFrom: 'revisionMoveFrom',
  moveTo: 'revisionMoveTo',
} as const;

function wrapper(
  ctx: ComposeContext,
  kind: keyof typeof WRAPPER_KINDS,
  suggestion: string,
  role: string,
  children: OoxmlNode[]
): OoxmlElement {
  return build(ctx.mint(), WRAPPER_KINDS[kind], kind, ctx.revisionMark(suggestion, role), children);
}

function inlineContainer(
  ctx: ComposeContext,
  open: OrderedItem,
  children: OoxmlNode[]
): OoxmlElement {
  const shell = open.node;
  if (!shell || shell.kind === 'textValue')
    return build(ctx.mint(), 'generic', 'customXml', [], children);
  if (open.token !== 'sdt') return withChildren(shell, [...shell.children, ...children]);
  const original = ctx.doc.originals.get(shell.id);
  let props = containerPropertiesAfter(ctx, open, shell.children);
  if (ctx.unprompted.has(open.space)) {
    props = hidePrompt(ctx, props, original, {
      inline: true,
      retain: !ctx.ownFilled.has(open.space),
    });
  } else if (ctx.filledPrompts.has(open.space)) {
    props = hidePrompt(ctx, props, original, { hide: false, inline: true });
  }

  return withChildren(shell, [...props, controlContent(ctx, original, children)]);
}

/** A control's `w:sdtContent`, of the kind the reader gave the original. */
export function controlContent(
  ctx: ComposeContext,
  original: OoxmlNode | undefined,
  children: readonly OoxmlNode[]
): OoxmlElement {
  // The content keeps the kind the reader gave it (generic where its children break the typed
  // kind's rules, as a field in a plain text control does)
  const content =
    original && original.kind !== 'textValue'
      ? original.children.find((c) => isWml(c, 'sdtContent') || c.kind === 'contentControlContent')
      : undefined;
  const kind = content && content.kind !== 'textValue' ? content.kind : 'contentControlContent';
  return build(ctx.mint(), kind, 'sdtContent', [], children);
}
