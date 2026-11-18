// Runs as items: characters, objects in runs, markers inside runs, and the pieces of fields.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { withDeletedTextRestored } from '../../../store/store/revision-rebuild-helpers.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import type { OrderedFieldPiece, OrderedFormatRecord } from '../model/types.ts';
import { childWhere, isRunProperties, isWml, shellOf, withChildren } from '../ooxml-nodes.ts';
import { atomCharAt, fieldSlotOf, type ParagraphLoad } from './paragraph.ts';
import { faceOf } from './properties.ts';
import { pushItem, pushZero, type Labels, type LoadState } from './state.ts';

const isTextElement = (node: OoxmlElement) =>
  node.kind === 'text' ||
  node.kind === 'deletedText' ||
  isWml(node, 't') ||
  isWml(node, 'delText') ||
  isWml(node, 'instrText');

/** The text element's shell, in its shown form (`w:delText` as `w:t`). */
function textShellOf(element: OoxmlElement): OoxmlElement {
  const restored = withDeletedTextRestored(shellOf(element));
  return restored.kind === 'textValue' ? shellOf(element) : restored;
}

/** Load one run. */
export function loadRun(
  state: LoadState,
  run: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad
): void {
  const rPr = childWhere(run, isRunProperties);
  const { face, fmts } = faceOf(state, rPr);
  state.doc.runShells.set(run.id, shellOf(run));
  if (!run.children.some((c) => !isRunProperties(c) && c.kind !== 'textValue')) {
    // A run with nothing in it but its properties: kept whole, with no width. A file's change of
    // its properties stays its suggestion's, so a decision reaches it (compose, `decidedRunShell`)
    pushZero(state, run, labels, { face, fmts, ...fieldSlotOf(ctx) });
    return;
  }

  const inRun = { run: run.id, face, fmts };
  for (const child of run.children) {
    if (child.kind === 'textValue' || isRunProperties(child)) continue;
    if (ctx.coveredBy.has(child.id)) {
      loadFieldPiece(state, child, run, labels, ctx);
      continue;
    }

    if (isTextElement(child)) {
      loadText(state, child, run, labels, ctx, face, fmts);
      continue;
    }

    const span = ctx.index.spanOf(child);
    if (span && span.end - span.start === 1) {
      pushItem(state, {
        space: ctx.space,
        offset: span.start,
        kind: 'atom',
        text: atomCharAt(ctx, span),
        node: child,
        ...inRun,
        ...labels,
      });
    } else {
      pushZero(state, child, labels, { inRun: true, ...inRun });
    }
  }
}

function loadText(
  state: LoadState,
  element: OoxmlElement,
  run: OoxmlElement,
  labels: Labels,
  ctx: ParagraphLoad,
  face: OoxmlElement | null,
  fmts: readonly OrderedFormatRecord[]
): void {
  const inRun = { run: run.id, face, fmts };
  if (!element.children.some((v) => v.kind === 'textValue' && v.value.length > 0)) {
    // An empty text element: kept in its run, with no width
    pushZero(state, element, labels, { inRun: true, ...inRun });
    return;
  }

  const textShell = textShellOf(element);
  for (const value of element.children) {
    if (value.kind !== 'textValue' || value.value.length === 0) continue;
    const span = ctx.index.spanOf(value);
    if (!span || span.end - span.start !== value.value.length) {
      // Text the engine does not count (an orphan instruction): kept whole, with no width
      pushZero(state, element, labels, { inRun: true, ...inRun });
      return;
    }

    pushItem(state, {
      space: ctx.space,
      offset: span.start,
      length: value.value.length,
      kind: 'text',
      text: value.value,
      textShell,
      ...inRun,
      ...labels,
      ...(isWml(element, 'instrText') ? { instr: true } : {}),
    });
  }
}

/**
 * Add a run's file format records to a field's records, one per suggestion, and give the run's
 * own face under each by the id of the field's record (the record a later policy may hand to
 * another suggestion keeps its id, so the face stays found).
 */
function addFormats(
  into: OrderedFormatRecord[],
  found: readonly OrderedFormatRecord[]
): Pick<OrderedFieldPiece, 'recordFaces' | 'recordFiles'> {
  if (found.length === 0) return {};
  const faces = new Map<string, OoxmlElement | null>();
  const files = new Map<string, OoxmlElement>();
  for (const record of found) {
    let kept = into.find((other) => other.suggestion === record.suggestion);
    if (!kept) {
      kept = record;
      into.push(record);
    }

    if (record.kind !== 'face') continue;
    faces.set(kept.id, record.face);

    // Each run keeps its own record: the faces before the change differ run by run
    if (record.fileRecord) files.set(kept.id, record.fileRecord);
  }

  return { recordFaces: faces, ...(files.size > 0 ? { recordFiles: files } : {}) };
}

/** Field begins less field ends in a node. */
function fieldCharBalance(node: OoxmlNode, depth = 0): number {
  if (node.kind === 'textValue' || depth > 8) return 0;
  if (node.localName === 'fldChar') {
    switch (node.attributes.find((a) => a.localName === 'fldCharType')?.value) {
      case 'begin':
        return 1;
      case 'end':
        return -1;
      default:
        return 0;
    }
  }

  let balance = 0;
  for (const child of node.children) balance += fieldCharBalance(child, depth + 1);
  return balance;
}

/**
 * A node a whole complex field covers: it joins the field's one unit, as a piece of its run.
 * The field's item is made at its first piece; the paragraph's load gives it its pieces.
 */
export function loadFieldPiece(
  state: LoadState,
  element: OoxmlElement,
  run: OoxmlElement | null,
  labels: Labels,
  ctx: ParagraphLoad
): void {
  const begin = ctx.coveredBy.get(element.id);
  if (begin === undefined) return;

  // A format record on a field run: the field's format, and the run keeps its old face
  const rPr = run ? childWhere(run, isRunProperties) : undefined;
  const { face, fmts: found } = faceOf(state, rPr);
  let field = ctx.fields.get(begin);
  if (!field) {
    const span = ctx.index.spanOf(begin);
    const at = pushItem(state, {
      space: ctx.space,
      offset: span?.start ?? 0,
      kind: 'atom',
      text: atomCharAt(ctx, span),
      run: run?.id ?? element.id,
      face,
      ...labels,
    });
    field = { at, key: `${ctx.space}|${span?.start ?? 0}`, pieces: [], fmts: [], depth: 0 };
    ctx.fields.set(begin, field);
  }

  field.depth += fieldCharBalance(element);
  ctx.openField = field.depth > 0 ? field : null;
  const records = addFormats(field.fmts, found);
  const last = field.pieces[field.pieces.length - 1];
  if (run && last && last.sourceRun === run.id && last.node.kind !== 'textValue') {
    // Pieces of one run stay one run
    field.pieces[field.pieces.length - 1] = {
      ...last,
      node: withChildren(last.node, [...last.node.children, element]),
      ...records,
    };
    return;
  }

  const piece: OrderedFieldPiece = run
    ? {
        node: withChildren(run, [...(face ? [face] : []), element]),
        sourceRun: run.id,
        ...records,
      }
    : { node: element };
  field.pieces.push(piece);
}

/**
 * A simple field the engine counts as one unit: its shell, its runs as pieces (revisions out,
 * each run with its own face), and the revisions as labels: the `w:ins` or `w:del` around its
 * runs, and the format records in them.
 */
export function simpleFieldAtom(
  state: LoadState,
  element: OoxmlElement,
  ctx: ParagraphLoad
): {
  node: OoxmlElement;
  pieces: OrderedFieldPiece[];
  by: string | null;
  dels: string[];
  fmts: OrderedFormatRecord[];
} {
  let by: string | null = null;
  const dels: string[] = [];
  const fmts: OrderedFormatRecord[] = [];
  const pieces: OrderedFieldPiece[] = [];
  const take = (node: OoxmlNode): void => {
    if (node.kind !== 'run') {
      pieces.push({ node });
      return;
    }

    const rPr = childWhere(node, isRunProperties);
    const { face, fmts: found } = faceOf(state, rPr);
    const records = addFormats(fmts, found);
    const restored = withDeletedTextRestored(
      withChildren(node, [...(face ? [face] : []), ...node.children.filter((c) => c !== rPr)])
    );
    pieces.push({ node: restored, ...records });
  };
  for (const child of element.children) {
    if (child.kind !== 'textValue' && (isWml(child, 'ins') || isWml(child, 'del'))) {
      if (isWml(child, 'ins')) by = state.revisions.suggestionFor(child, state.markRoles.run);
      else {
        const deleter = state.revisions.suggestionFor(child, REVISION_ROLE.deleteIn(ctx.markKey));
        if (!dels.includes(deleter)) dels.push(deleter);
      }

      for (const run of child.children) take(run);
    } else take(child);
  }

  return { node: shellOf(element), pieces, by, dels, fmts };
}
