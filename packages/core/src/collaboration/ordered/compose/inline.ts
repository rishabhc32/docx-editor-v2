// A paragraph's inline items back to runs, containers and revision wrappers. Pending
// suggestions are drawn as Word draws them (`w:ins`, `w:del`, `w:moveFrom`, `w:moveTo` with
// their range markers, `w:rPrChange`); decided ones are applied.

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { withDeletedTextRestored } from '../../../store/store/revision-rebuild-helpers.ts';
import { build, textNode } from '../../../store/store/tree-op-tracked.ts';
import { isSplitEdge, movedAwayBy, statusOf } from '../model/document.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import {
  isUnit,
  type OrderedFieldPiece,
  type OrderedFormatRecord,
  type OrderedItem,
  type OrderedPropertyRecord,
} from '../model/types.ts';
import {
  childWhere,
  fieldCharBalance,
  isRunProperties,
  isWml,
  renamed,
  withChildren,
  withId,
  wmlAttribute,
} from '../ooxml-nodes.ts';
import type { ComposeContext } from './context.ts';
import { isNotesStory } from './note-references.ts';
import { faceAfter, containerPropertiesAfter } from './properties.ts';
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
  /** The pending insertion of a new control, drawn as one `w:ins` around the control. */
  cover?: string;
}

interface RunBuild {
  readonly key: string;
  readonly item: OrderedItem;
  readonly content: OoxmlNode[];
  readonly deleted: boolean;
  text: string;
  textShell: OoxmlElement | undefined;
  /** The text form's item is drawn in its own paragraph: the form keeps its id. */
  textHomed: boolean;
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

/** A unit in the boxes: the innermost held it; each one it shows in, reached or struck. */
function noteUnit(ctx: ComposeContext, boxes: readonly Box[], item: OrderedItem): void {
  const innermost = boxes[boxes.length - 1];
  if (innermost) innermost.hadUnits = true;

  // A decided insertion or deletion reaches every control around it
  if (ctx.decided(item)) for (const b of boxes) b.reached = true;
  if (ctx.decidedAway(item)) for (const b of boxes) b.decided = true;
  if (ctx.gone(item)) return;
  const struck = ctx.struckBy(item) !== null;
  for (const b of boxes) {
    b.drawn = true;
    if (struck) b.struck = true;
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

  // A rejected new link that still holds something stays, as in Word
  const kept = keptLinkEdges(ctx, buffer);
  const hidden = (item: OrderedItem): boolean => ctx.gone(item) && !kept.has(item);
  const flushText = (build: RunBuild): void => {
    if (!build.text) return;
    build.content.push(
      textElement(ctx, build.text, build.textShell, build.deleted, build.textHomed)
    );
    build.text = '';
  };
  const flush = (): void => {
    if (!current) return;
    flushText(current);
    place(ctx, box().children, runNode(ctx, current.item, current.content), current.item);
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
    if (ctx.gone(item)) continue;
    if (item.kind === 'zero' && !item.inRun) {
      flush();
      if (item.node && item.moveRange) box().children.push(item.node);
      else if (item.node) place(ctx, box().children, decidedRunShell(ctx, item, item.node), item);
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
        item,
        content: [],
        deleted: struckText(ctx, item),
        text: '',
        textShell: undefined,
        textHomed: false,
      };
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

  return withMoveRanges(ctx, box().children);
}

/** The `w:sdtPr` keys a change of a control's title, tag or lock sets: no value. */
const METADATA_KEYS = new Set(
  ['alias', 'tag', 'lock'].map((name) => `${WML_NAMESPACE_URI} ${name}`)
);

/**
 * A record of a control's value (its words, a list's last value, a date, a check box), not only
 * of its title, tag or lock: the value records follow the text the control shows.
 */
function isValueRecord(record: OrderedPropertyRecord): boolean {
  if (record.target !== 'sdtPr') return false;
  const merge = record.kind === 'element' ? record.merge : undefined;
  return merge === undefined || merge.length === 0 || merge.some((key) => !METADATA_KEYS.has(key));
}

/**
 * The pending insertion a new inline control is drawn inside, as one `w:ins` around the whole
 * control (the engine's tracked control and Word's `w:ins` around `w:sdt`), so a reject takes
 * the control with its words. Not for a control an Enter split: a decision keeps both halves.
 */
function coverOf(ctx: ComposeContext, open: OrderedItem): string | null {
  if (open.token !== 'sdt' || isSplitEdge(open) || open.by === null) return null;
  if (ctx.covered.has(open.by)) return null;
  return ctx.insertedBy(open);
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

    // A control with a value record keeps the record in step with what it shows, once no
    // pending deletion is in it
    const valued = !done.struck && open.props.some(isValueRecord);
    if (done.reached || emptied || valued) ctx.reached.push(container.id);
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

function runKey(ctx: ComposeContext, item: OrderedItem): string {
  return [
    item.run,
    item.face ? item.face.id : '',
    item.fmts.map((f) => `${f.id}:${f.suggestion}`).join(','),
    ctx.insertedBy(item) ?? '',
    ctx.struckBy(item) ?? '',
    item.move ?? '',
    item.instr ? 'instr' : '',
  ].join('|');
}

function runNode(ctx: ComposeContext, item: OrderedItem, content: OoxmlNode[]): OoxmlElement {
  const shell = ctx.doc.runShells.get(item.run);
  const homed = ctx.homed(item);
  const rPr = ctx.fresh(faceAfter(ctx, item.fmts, item.face), homed);
  const id = ctx.runId(homed ? shell?.id : undefined);
  return withId(
    withChildren(shell ?? build(id, 'run', 'r', [], []), [...(rPr ? [rPr] : []), ...content]),
    id
  );
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

    into.push(withChildren(node, inner));
    return;
  }

  place(ctx, into, node, item);
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
      node = withId(node, ctx.runId(node.id));
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
  let outer: { readonly key: string; readonly suggestion: string } | null = null;
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
    outer = { key, suggestion: del };
    ctx.wrappers.set(wrapped, { key, outer: key, suggestion: del, move: moved ? item.move : null });
  }

  if (by) {
    const kind = item.move !== null ? 'moveTo' : 'ins';
    const role = insertionRole(ctx, item);
    const id = ctx.revisionId(by, role);
    wrapped = wrapper(ctx, kind, by, role, [wrapped]);
    key = `${kind}:${id}:${by}|${key}`;
    outer = { key: `${kind}:${id}:${by}`, suggestion: by };
  }

  if (outer === null || wrapped.kind === 'textValue') {
    into.push(wrapped);
    return;
  }

  const info = { key, outer: outer.key, suggestion: outer.suggestion, move: item.move };
  const last = into[into.length - 1];
  if (last && last.kind !== 'textValue' && ctx.wrappers.get(last)?.outer === outer.key) {
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

/** A move wrapper's half, as its range markers name it (`w:moveFromRangeStart`, ...). */
const MOVE_HALVES: ReadonlyMap<string, 'From' | 'To'> = new Map([
  ['moveFrom', 'From'],
  ['moveTo', 'To'],
]);

/** A move's wrappers between its range markers, as Word writes a move. */
function withMoveRanges(ctx: ComposeContext, children: OoxmlNode[]): OoxmlNode[] {
  if (!children.some((c) => c.kind !== 'textValue' && ctx.wrappers.get(c)?.move)) return children;
  const out: OoxmlNode[] = [];
  for (let i = 0; i < children.length; ) {
    const child = children[i];
    if (child === undefined) break;
    const info = child.kind === 'textValue' ? undefined : ctx.wrappers.get(child);
    const half = child.kind === 'textValue' ? undefined : MOVE_HALVES.get(child.localName);
    if (!info?.move || half === undefined || ctx.doc.movesWithRanges.has(info.move)) {
      out.push(child);
      i += 1;
      continue;
    }

    let j = i;
    while (j < children.length) {
      const next = children[j];
      if (!next || next.kind === 'textValue' || ctx.wrappers.get(next)?.key !== info.key) break;
      j += 1;
    }

    const role = half === 'From' ? REVISION_ROLE.moveFromRange : REVISION_ROLE.moveToRange;
    const range = ctx.revisionMark(info.suggestion, role);
    const [idAttribute, ...rest] = range;
    if (idAttribute === undefined) break;
    const named = [idAttribute, wmlAttribute('name', info.move), ...rest];
    out.push(build(ctx.mint(), `move${half}RangeStart`, `move${half}RangeStart`, named, []));
    for (let k = i; k < j; k += 1) {
      const moved = children[k];
      if (moved !== undefined) out.push(moved);
    }

    out.push(build(ctx.mint(), `move${half}RangeEnd`, `move${half}RangeEnd`, [idAttribute], []));
    i = j;
  }

  return out;
}

function inlineContainer(ctx: ComposeContext, open: OrderedItem, drawn: OoxmlNode[]): OoxmlElement {
  // A move's half inside a link, a control or another inline container has its range markers
  // there (each holds paragraph content, `EG_PContent`, which range markup is part of), or a
  // reopened file cannot pair it with the other half
  const children = withMoveRanges(ctx, drawn);
  const shell = open.node;
  if (!shell || shell.kind === 'textValue')
    return build(ctx.mint(), 'generic', 'customXml', [], children);
  if (open.token !== 'sdt') return withChildren(shell, [...shell.children, ...children]);
  const original = ctx.doc.originals.get(shell.id);
  let props = containerPropertiesAfter(ctx, open, shell.children);
  if (ctx.unprompted.has(open.space)) {
    props = hidePrompt(ctx, props, original, { inline: true });
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
