// A paragraph as items. Every unit gets the width the engine's offset index gives it, so a
// paragraph's shown units read as the engine's model text, but for a field one unit cannot hold
// (`offsetIndexOf`). Revision markup becomes labels: the same path loads a plain file and a file
// with tracked changes.

import { atomicFieldSpansOf } from '../../../store/package/field-nodes.ts';
import {
  isInlineRunContainer,
  MAX_INLINE_CONTAINER_DEPTH,
} from '../../../store/package/ooxml-shared.ts';
import type {
  OoxmlElement,
  OoxmlNode,
  OoxmlParagraphNode,
} from '../../../store/package/ooxml-tree.ts';
import { paragraphModelTextOf } from '../../../store/store/paragraph-model-text.ts';
import {
  paragraphOffsetIndex,
  transientParagraphOffsetIndex,
  type ParagraphOffsetIndex,
} from '../../../store/store/tree-op-segments.ts';
import { formatRunsOf } from '../compose/inline.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import { equationAtom, type EquationAtom } from './equations.ts';
import {
  markSpace,
  paragraphSpace,
  type OrderedContainerToken,
  type OrderedFieldPiece,
  type OrderedFormatRecord,
} from '../model/types.ts';
import {
  childWhere,
  isControlContent,
  isParagraphProperties,
  isRunProperties,
  isWml,
  showsPlaceholder,
  withChildren,
  withMarkFace,
} from '../ooxml-nodes.ts';
import { loadMoveRangeMarker, movedBy } from './moves.ts';
import { faceOf, paragraphProperties, containerShell } from './properties.ts';
import { loadFieldPiece, loadRun, simpleFieldAtom } from './runs.ts';
import {
  checkDepth,
  markerLabels,
  markPrompt,
  moveName,
  pushClose,
  pushItem,
  pushOpen,
  pushZero,
  type Labels,
  type LoadState,
} from './state.ts';

/** A whole complex field being gathered: its item, its pieces and its format records. */
export interface FieldBuild {
  readonly at: number;
  /** The field item's address, `space|offset`. */
  readonly key: string;
  readonly pieces: OrderedFieldPiece[];
  readonly fmts: OrderedFormatRecord[];
  /** Field begins less ends in its pieces so far: the field is open while above zero. */
  depth: number;
}

/** What the walk of one paragraph's inline content knows. */
export interface ParagraphLoad {
  readonly node: OoxmlParagraphNode;
  readonly index: ParagraphOffsetIndex;
  /** The engine's model text: an atom's character is read from it. */
  readonly text: string;
  /** The space of the paragraph's units. */
  readonly space: string;
  /** The mark's space: struck runs take one revision id per paragraph. */
  readonly markKey: string;
  /** A node a whole complex field covers, to the field's begin node. */
  readonly coveredBy: ReadonlyMap<string, string>;
  readonly fields: Map<string, FieldBuild>;
  /** The field whose pieces are being read, until its end: a marker between them rides in it. */
  openField: FieldBuild | null;
}

const RANGE_MARKERS = new Set([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'permStart',
  'permEnd',
  'proofErr',
  'customXmlInsRangeStart',
  'customXmlInsRangeEnd',
  'customXmlDelRangeStart',
  'customXmlDelRangeEnd',
]);

const isFieldChar = (node: OoxmlNode) =>
  node.kind === 'fldChar' || (node.kind !== 'textValue' && node.localName === 'fldChar');

/** The character of an object the engine's model text does not spell. */
const OBJECT_CHAR = '￼';

/** The character an atom at `span` shows: the engine's model text, else the object character. */
export function atomCharAt(ctx: ParagraphLoad, span: { readonly start: number } | null): string {
  return span ? ctx.text.charAt(span.start) || OBJECT_CHAR : OBJECT_CHAR;
}

/** The engine's whole complex fields: the segment of each one's begin node. */
const wholeFields = (index: ParagraphOffsetIndex) =>
  index.segments.filter((segment) => isFieldChar(segment.node) && segment.removeNodeIds?.length);

/** A complex field is one unit: its begin node, and every node it covers, to the begin node. */
function fieldCover(index: ParagraphOffsetIndex): Map<string, string> {
  const coveredBy = new Map<string, string>();
  for (const segment of wholeFields(index)) {
    coveredBy.set(segment.node.id, segment.node.id);
    for (const id of segment.removeNodeIds ?? []) {
      coveredBy.set(id, segment.node.id);
    }
  }

  return coveredBy;
}

/** The content of each run, to the node that holds the run: the paragraph, a link, a revision. */
function runHomes(paragraph: OoxmlParagraphNode): Map<string, string> {
  const homes = new Map<string, string>();
  const visit = (node: OoxmlElement, depth: number): void => {
    for (const child of node.children) {
      if (child.kind === 'run') {
        for (const content of child.children) {
          homes.set(content.id, node.id);
        }
      } else if (child.kind !== 'textValue' && depth < MAX_INLINE_CONTAINER_DEPTH) {
        visit(child, depth + 1);
      }
    }
  };
  visit(paragraph, 0);
  return homes;
}

/**
 * The engine's offset index of a paragraph, but for a complex field whose runs are not all where
 * its begin run is: one unit cannot hold a link or a revision that starts or ends in the field (a
 * table of contents field whose result is a link, a suggestion in a field's result). Such a field
 * loads as its markers and its result, as a field over two paragraphs does, as the room holds it.
 */
function offsetIndexOf(paragraph: OoxmlParagraphNode): ParagraphOffsetIndex {
  const index = paragraphOffsetIndex(paragraph);
  const fields = wholeFields(index);
  if (fields.length === 0) {
    return index;
  }

  const homes = runHomes(paragraph);
  const split = new Set(
    fields
      .filter((field) =>
        field.removeNodeIds?.some((id) => homes.get(id) !== homes.get(field.node.id))
      )
      .map((field) => field.node.id)
  );
  if (split.size === 0) {
    return index;
  }

  const spans = atomicFieldSpansOf(paragraph).filter((span) => !split.has(span.node.id));
  return transientParagraphOffsetIndex(paragraph, spans);
}

/** Load one paragraph: its inline content, then its mark. */
export function loadParagraph(
  state: LoadState,
  node: OoxmlParagraphNode,
  labels: Labels,
  depth: number
): void {
  checkDepth(state, depth);
  const index = offsetIndexOf(node);
  const ctx: ParagraphLoad = {
    node,
    index,
    text: paragraphModelTextOf(node, index.segments),
    space: paragraphSpace(state.doc.file, node.id),
    markKey: markSpace(state.doc.file, node.id),
    coveredBy: fieldCover(index),
    fields: new Map(),
    openField: null,
  };
  const outer = state.home;
  state.home = node.id;
  const start = state.items.length;
  for (const child of node.children) {
    if (!isParagraphProperties(child)) loadInline(state, child, labels, ctx, depth + 1);
  }

  for (const { at, pieces, fmts } of ctx.fields.values()) {
    const item = state.items[at];
    if (item) {
      // The field's face is that of the run a format reaches, as the engine reads it
      const run = pieces[formatRunsOf(pieces).indexOf(true)]?.node;
      const face =
        run === undefined || run.kind === 'textValue'
          ? item.face
          : (childWhere(run, isRunProperties) ?? null);
      state.items[at] = { ...item, face, pieces, fmts };
    }
  }

  pushMark(state, ctx, labels, start);
  state.home = outer;
}

/** The paragraph's mark, which counts the units of the items from `start` on. */
function pushMark(state: LoadState, ctx: ParagraphLoad, labels: Labels, start: number): void {
  const pPr = childWhere(ctx.node, isParagraphProperties);
  const markRPr = pPr ? childWhere(pPr, isRunProperties) : undefined;
  const mark = markerLabels(state, markRPr, labels, state.markRoles);
  const { face, fmts } = faceOf(state, markRPr);
  const { own, props } = paragraphProperties(state, pPr);
  let unitCount = 0;
  for (let k = start; k < state.items.length; k += 1) {
    unitCount += state.items[k]?.length ?? 0;
  }

  pushItem(state, {
    space: ctx.markKey,
    kind: 'mark',
    node: withChildren(ctx.node, own ? [withMarkFace(own, face)] : []),
    face,
    fmts,
    props,
    ...mark,
    unitCount,
  });
}

/**
 * The container token of a link, a smart tag, inline custom XML, a direction override or an
 * editable simple field (its result runs are text), or null.
 */
function inlineToken(node: OoxmlElement): OrderedContainerToken | null {
  if (node.kind === 'hyperlink') {
    return 'hyperlink';
  }

  if (node.kind === 'fldSimple' || node.localName === 'fldSimple') {
    return 'fldSimple';
  }

  if (!isInlineRunContainer(node)) {
    return null;
  }

  switch (node.localName) {
    case 'smartTag':
    case 'customXml':
    case 'dir':
    case 'bdo':
      return node.localName;
    default:
      return null;
  }
}

/** Load one inline node: a revision wrapper, a run, a container, a marker or an object. */
export function loadInline(
  state: LoadState,
  node: OoxmlNode,
  labels: Labels,
  ctx: ParagraphLoad,
  depth: number
): void {
  if (node.kind === 'textValue') return;
  checkDepth(state, depth);
  if (loadRevisionWrapper(state, node, labels, ctx, depth)) {
    return;
  }

  if (node.kind === 'run') {
    loadRun(state, node, labels, ctx);
    return;
  }

  // A complex field that begins outside a run
  if (ctx.coveredBy.has(node.id)) {
    loadFieldPiece(state, node, null, labels, ctx);
    return;
  }

  const span = ctx.index.spanOf(node);
  const width = span ? span.end - span.start : 0;

  // A simple field or an equation that the engine counts as one unit
  if (
    span &&
    width === 1 &&
    (node.kind === 'fldSimple' ||
      node.localName === 'fldSimple' ||
      node.localName === 'oMath' ||
      node.localName === 'oMathPara')
  ) {
    loadOneUnitAtom(state, node, labels, ctx, span);
    return;
  }

  // An inline container: open and close items around its content
  const token = inlineToken(node);
  if (token !== null) {
    const open = pushOpen(state, node, token, withChildren(node, []), labels);
    loadInlineChildren(state, node, labels, ctx, depth);
    pushClose(state, open, labels);
    return;
  }

  if (node.kind === 'contentControl' || isWml(node, 'sdt')) {
    loadInlineControl(state, node, labels, ctx, depth);
    return;
  }

  if (loadMoveRangeMarker(state, node, labels, fieldSlotOf(ctx))) return;
  if (RANGE_MARKERS.has(node.localName) || width === 0 || !span) {
    pushZero(state, node, labels, fieldSlotOf(ctx));
    return;
  }

  // Anything else with a width: one unit, kept whole
  pushItem(state, {
    space: ctx.space,
    offset: span.start,
    kind: 'atom',
    text: atomCharAt(ctx, span),
    node,
    run: node.id,
    bare: true,
    ...labels,
  });
}

function loadInlineChildren(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad,
  depth: number
): void {
  for (const child of node.children) {
    loadInline(state, child, labels, ctx, depth + 1);
  }
}

/**
 * Load a revision wrapper's content with the labels the wrapper gives it. False when the node is
 * not a revision wrapper.
 */
function loadRevisionWrapper(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad,
  depth: number
): boolean {
  let inner: Labels;
  if (node.kind === 'revisionInsert' || isWml(node, 'ins')) {
    inner = { ...labels, by: state.revisions.suggestionFor(node, state.markRoles.run) };
  } else if (node.kind === 'revisionDelete' || isWml(node, 'del')) {
    const deleter = state.revisions.suggestionFor(node, REVISION_ROLE.deleteIn(ctx.markKey));
    inner = { ...labels, dels: [...labels.dels, deleter] };
  } else if (node.kind === 'revisionMoveFrom' || isWml(node, 'moveFrom')) {
    const deleter = movedBy(state, node, state.moveFrom, REVISION_ROLE.moveFromIn(ctx.markKey));
    inner = { ...labels, dels: [...labels.dels, deleter], move: moveName(state.moveFrom, node) };
  } else if (node.kind === 'revisionMoveTo' || isWml(node, 'moveTo')) {
    const by = movedBy(state, node, state.moveTo, REVISION_ROLE.moveTo);
    inner = { ...labels, by, move: moveName(state.moveTo, node) };
  } else {
    return false;
  }

  loadInlineChildren(state, node, inner, ctx, depth);
  return true;
}

/** A simple field or an equation that the engine counts as one unit, at `span`. */
function loadOneUnitAtom(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad,
  span: { readonly start: number }
): void {
  // A simple field the reader knows by its kind alone is kept whole
  let atom: EquationAtom | ReturnType<typeof simpleFieldAtom> = {
    node,
    pieces: undefined,
    by: null,
    dels: [],
    fmts: [],
  };
  if (node.localName === 'fldSimple') {
    atom = simpleFieldAtom(state, node, ctx);
  } else if (node.localName === 'oMath' || node.localName === 'oMathPara') {
    atom = equationAtom(state, node, ctx.markKey);
  }

  pushItem(state, {
    space: ctx.space,
    offset: span.start,
    kind: 'atom',
    text: atomCharAt(ctx, span),
    node: atom.node,
    ...(atom.pieces ? { pieces: atom.pieces } : {}),
    run: node.id,
    bare: true,
    by: atom.by ?? labels.by,
    dels: [...labels.dels, ...atom.dels],
    fmts: atom.fmts,
    move: labels.move,
    within: labels.within,
  });
}

/** A marker read while a field's pieces are being read rides in that field. */
export function fieldSlotOf(ctx: ParagraphLoad): { fieldSlot?: { field: string; before: number } } {
  const field = ctx.openField;
  return field ? { fieldSlot: { field: field.key, before: field.pieces.length } } : {};
}

function loadInlineControl(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad,
  depth: number
): void {
  const content = childWhere(node, isControlContent);
  const { shell } = containerShell(state, node, (c) => !isControlContent(c));
  const open = pushOpen(state, node, 'sdt', shell, labels);
  const first = state.items.length;
  for (const child of content?.children ?? []) loadInline(state, child, labels, ctx, depth + 1);

  // A control showing its prompt: the prompt is not content
  if (showsPlaceholder(node)) {
    markPrompt(state, first, () => true);
  }

  pushClose(state, open, labels);
}
