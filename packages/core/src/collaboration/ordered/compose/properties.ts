// Faces and properties under their records: the accepted ones applied in the order they were
// made (not the order of the decisions), the pending ones drawn on top with Word's change
// record (`w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, ...), written by the engine's writer.

import { orderedContentControlProperties } from '../../../store/package/content-control-nodes.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import {
  AUTHORABLE_RUN_PROPERTIES,
  authoredProperties,
  mergedProperties,
} from '../../../store/store/direct-properties.ts';
import { mergedPropertyChildren, RUN_VOCABULARY } from '../../../store/store/tree-op-properties.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import {
  CT_TCPR_SEQUENCE,
  CT_TBLPR_SEQUENCE,
  CT_TRPR_SEQUENCE,
} from '../../../store/store/tree-op-table-properties.ts';
import {
  withPropertyChangeRecord,
  type PropertyChangeContainer,
} from '../../../store/store/tree-op-tracked-properties.ts';
import { isPending, statusOf } from '../model/document.ts';
import { fileMark, REVISION_ROLE } from '../model/revision-ids.ts';
import type {
  OrderedDocument,
  OrderedElementRecord,
  OrderedFormatRecord,
  OrderedFormatWrite,
  OrderedItem,
  OrderedParagraphOpRecord,
  OrderedPropertyRecord,
  OrderedPropertyTarget,
  OrderedStyleWrite,
} from '../model/types.ts';
import { attributeOf, isWml, withAttributes, withChildren, wmlAttribute } from '../ooxml-nodes.ts';
import { isChangeRecord, type ComposeContext, type DrawnTakeover } from './context.ts';
import { withoutRemoved, withParagraphOp } from './paragraph-ops.ts';

type AnyRecord = OrderedFormatRecord | OrderedPropertyRecord;

const bySeq = (a: AnyRecord, b: AnyRecord) => a.seq - b.seq;

/** The accepted records and the pending ones drawn, each in the order they were made. */
export function acceptedAndPending<T extends AnyRecord>(
  ctx: ComposeContext,
  records: readonly T[]
): { accepted: T[]; pending: T[] } {
  const accepted = records
    .filter((r) => statusOf(ctx.doc, r.suggestion) === 'accepted')
    .sort(bySeq);
  const pending =
    ctx.options.pending === 'reject'
      ? []
      : records
          .filter(
            (r) =>
              statusOf(ctx.doc, r.suggestion) === 'pending' &&
              !('acceptedOnly' in r && r.acceptedOnly === true)
          )
          .sort(bySeq);
  return { accepted, pending };
}

/** Run properties with `properties` merged in, as the engine merges a format press. */
export function faceWithProperties(
  face: readonly OoxmlNode[],
  write: OrderedFormatWrite,
  nextId: () => string
): OoxmlNode[] {
  const container = build(nextId(), 'runProperties', 'rPr', [], face);
  const authored = authoredProperties(container, AUTHORABLE_RUN_PROPERTIES);
  return mergedPropertyChildren(
    face,
    withoutRemoved(mergedProperties(authored, write.properties), write.remove),
    RUN_VOCABULARY,
    nextId
  );
}

function applyFormat(
  record: OrderedFormatRecord,
  children: readonly OoxmlNode[],
  nextId: () => string
): readonly OoxmlNode[] {
  if (record.kind === 'face')
    return record.face?.children.filter((c) => !isWml(c, 'rPrChange')) ?? [];
  const merged = faceWithProperties(children, record, nextId);
  return record.style === undefined ? merged : withStyleWrite(merged, record.style, nextId);
}

/** Run properties with a character style written, as the engine's link ops write it. */
export function withStyleWrite(
  children: readonly OoxmlNode[],
  write: OrderedStyleWrite,
  nextId: () => string
): OoxmlNode[] {
  const names = (c: OoxmlNode, value?: string) =>
    c.kind !== 'textValue' &&
    isWml(c, 'rStyle') &&
    (value === undefined || c.attributes.some((a) => a.localName === 'val' && a.value === value));
  if ('remove' in write) return children.filter((c) => !names(c, write.remove));
  const style = build(nextId(), 'generic', 'rStyle', [wmlAttribute('val', write.set)], []);

  // `w:rStyle` is the first child of `w:rPr` (the engine's `withCharacterStyle`)
  return [style, ...children.filter((c) => !names(c))];
}

/**
 * Note a drawn record in `taken` when it is a takeover (policy 2): a pending record taken from a
 * suggestion still pending, with the properties before it and with it.
 */
export function noteTakeover(
  ctx: ComposeContext,
  taken: DrawnTakeover[],
  record: AnyRecord,
  before: readonly OoxmlNode[],
  after: readonly OoxmlNode[]
): void {
  const { suggestion, replaced } = record;
  if (replaced === undefined || !isPending(ctx.doc, suggestion) || !isPending(ctx.doc, replaced)) {
    return;
  }

  const kept = (c: OoxmlNode) => !isChangeRecord(c);
  taken.push({ replaced, before: before.filter(kept), after: after.filter(kept) });
}

/**
 * The file's own change record, when it is the one record drawn: pending, alone, with nothing
 * accepted under it, and giving the whole properties. Written back as the file had it (a
 * record's copy can hold what the engine's writer leaves out, such as a cell's old `w:cellDel`),
 * with its suggestion's attribution (`fileMark`). A record that sets only some children
 * (`merge`) shows the container's own others, and load made those the nodes of the file's
 * record: it takes a record of copies, as a room's record does.
 */
export function fileRecordAlone(
  doc: OrderedDocument,
  accepted: readonly AnyRecord[],
  pending: readonly AnyRecord[]
): OoxmlElement | undefined {
  const [only, ...more] = pending;
  if (
    !only ||
    more.length > 0 ||
    accepted.length > 0 ||
    ('merge' in only && only.merge !== undefined)
  ) {
    return undefined;
  }

  return 'fileRecord' in only && only.fileRecord !== undefined
    ? fileMark(doc, only.fileRecord, only.suggestion)
    : undefined;
}

/**
 * Accepted records made before the first pending one, and those made after it. The engine
 * keeps one change record holding the properties before the first pending change; a direct
 * change made later shows on top of it.
 */
export function aroundPending<T extends AnyRecord>(
  accepted: readonly T[],
  pending: readonly AnyRecord[]
): [T[], T[]] {
  const first = pending[0]?.seq;
  if (first === undefined) return [[...accepted], []];
  return [accepted.filter((r) => r.seq < first), accepted.filter((r) => r.seq > first)];
}

/** A format on its author's own pending words, shown with no change record. */
const isUntracked = (record: OrderedFormatRecord): boolean =>
  record.kind === 'properties' && record.untracked === true;

/**
 * The suggestion that holds the record a direct change was made over: the one it was made over,
 * or the one that took that record over later (policy 2). None once the record was taken back.
 */
function holderOf(records: readonly AnyRecord[], suggestion: string): string | undefined {
  if (records.some((record) => record.suggestion === suggestion)) return suggestion;
  return records.find((record) => 'replaced' in record && record.replaced === suggestion)
    ?.suggestion;
}

/**
 * Whether a direct record was made over a pending record whose holder is now rejected (or drawn
 * so): Word keeps one record, the properties before the first pending change, so the reject
 * takes the direct change away too. `keeps` names a rejected suggestion whose reject keeps them.
 */
export function madeOverRejected(
  ctx: ComposeContext,
  record: AnyRecord,
  records: readonly AnyRecord[],
  keeps: (suggestion: string) => boolean = () => false
): boolean {
  const over = 'over' in record ? (record.over ?? []) : [];
  return over.some((suggestion) => {
    const holder = holderOf(records, suggestion);
    return holder !== undefined && ctx.rejectedOrDrawnSo(holder) && !keeps(holder);
  });
}

/** The decided records drawn: the accepted ones, less those a reject took away. */
const shownAccepted = <T extends AnyRecord>(
  ctx: ComposeContext,
  accepted: readonly T[],
  records: readonly AnyRecord[]
): T[] => accepted.filter((record) => !madeOverRejected(ctx, record, records));

/** A run's (or a mark's) run properties under its format records, the change record included. */
export function faceAfter(
  ctx: ComposeContext,
  fmts: readonly OrderedFormatRecord[],
  base: OoxmlElement | null
): OoxmlElement | null {
  if (fmts.length === 0) return base;
  const shown = acceptedAndPending(ctx, fmts);
  const accepted = shownAccepted(ctx, shown.accepted, fmts);

  // A format on its author's own pending words is part of their insertion: no change record
  const inserted = shown.pending.filter(isUntracked);
  const pending = shown.pending.filter((record) => !isUntracked(record));
  const [earlier, later] = aroundPending([...accepted, ...inserted].sort(bySeq), pending);
  let children: readonly OoxmlNode[] = (base?.children ?? []).filter((c) => !isWml(c, 'rPrChange'));
  for (const record of earlier) children = applyFormat(record, children, ctx.mint);
  const old = children;
  const taken: DrawnTakeover[] = [];
  for (const record of pending) {
    const before = children;
    children = applyFormat(record, children, ctx.mint);
    noteTakeover(ctx, taken, record, before, children);
  }

  // A decided change made after a pending one shows on top, and the record keeps the original
  for (const record of later) children = applyFormat(record, children, ctx.mint);
  const latest = pending[pending.length - 1];
  const own = fileRecordAlone(ctx.doc, shown.accepted, pending);
  if (own) children = [...children, own];
  else if (latest) {
    // One record: the original, under the latest author (the engine's rule)
    children = ctx.withRecordAttributes(
      withPropertyChangeRecord({
        container: 'runProperties',
        prior: old,
        next: children,
        revision: ctx.attribution(latest.suggestion),
        mint: ctx.mint,
        nextRevisionId: () => ctx.revisionId(latest.suggestion, REVISION_ROLE.format),
        // A file's record stays as the file wrote it, even one that changes nothing
        recordUnchanged: latest.kind === 'face',
      }),
      latest.suggestion,
      REVISION_ROLE.format
    );
    ctx.keepTakeovers(children, taken);
  }

  if (children.length === 0 && !base) return null;
  return withChildren(base ?? build(ctx.mint(), 'runProperties', 'rPr', [], []), children);
}

/** The given element's children for a property record, the kept children `keep` selects. */
function applyProperties(
  record: OrderedElementRecord,
  children: readonly OoxmlNode[],
  keep: (c: OoxmlNode) => boolean
): readonly OoxmlNode[] {
  if (record.merge !== undefined) {
    return withMergedChildren(record.target, record.element, record.merge, children);
  }

  const given = (record.element?.children ?? []).filter(
    (c) => !keep(c) && !(c.kind !== 'textValue' && c.localName.endsWith('Change'))
  );
  return [...given, ...children.filter(keep)];
}

/** The key of a properties child a merge record names: its namespace URI and local name. */
export const propertyKey = (node: OoxmlElement): string => `${node.namespaceUri} ${node.localName}`;

/** The schema's order of a table's, a row's or a cell's properties children. */
const SEQUENCES: Partial<Record<OrderedPropertyTarget, readonly string[]>> = {
  tblPr: CT_TBLPR_SEQUENCE,
  tblPrEx: CT_TBLPR_SEQUENCE,
  trPr: CT_TRPR_SEQUENCE,
  tcPr: CT_TCPR_SEQUENCE,
};

/** `given` placed among `kept` where `sequence` has them; children it does not name stay. */
function inSequence(
  kept: readonly OoxmlNode[],
  given: readonly OoxmlNode[],
  sequence: readonly string[]
): OoxmlNode[] {
  const rank = (c: OoxmlNode) =>
    c.kind === 'textValue' || c.namespaceUri !== WML_NAMESPACE_URI
      ? -1
      : sequence.indexOf(c.localName);
  const out = [...kept];
  for (const node of given) {
    const at = out.findIndex((c) => rank(c) > rank(node));
    out.splice(at < 0 ? out.length : at, 0, node);
  }

  return out;
}

/**
 * The children of `element` that `merge` names over the ones before them, in the schema's order.
 */
export function withMergedChildren(
  target: OrderedPropertyTarget,
  element: OoxmlElement | null,
  merge: readonly string[],
  children: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  const keys = new Set(merge);
  const kept = children.filter((c) => c.kind === 'textValue' || !keys.has(propertyKey(c)));
  const given = (element?.children ?? []).filter(
    (c) => c.kind !== 'textValue' && keys.has(propertyKey(c))
  );
  const sequence = SEQUENCES[target];
  if (sequence) {
    return inSequence(kept, given, sequence);
  }

  const merged = [...kept, ...given];
  return target === 'sdtPr' ? orderedContentControlProperties(merged) : merged;
}

/** A paragraph mark's `w:pPr` under its property records (its run properties drawn apart). */
export function paragraphPropertiesAfter(
  ctx: ComposeContext,
  mark: OrderedItem,
  base: OoxmlElement | null
): OoxmlElement | null {
  const records = mark.props.filter(
    (p): p is OrderedElementRecord | OrderedParagraphOpRecord => p.target === 'pPr'
  );
  if (records.length === 0) return base;
  const { accepted, pending } = acceptedAndPending(ctx, records);
  let current: OoxmlElement | null = base
    ? withChildren(
        base,
        base.children.filter((c) => !isWml(c, 'pPrChange'))
      )
    : null;
  const [earlier, later] = aroundPending(shownAccepted(ctx, accepted, records), pending);
  for (const record of earlier) current = withDecidedMarkRecord(ctx, current, record, keptByMark);
  const own = fileRecordAlone(ctx.doc, accepted, pending);
  const taken: DrawnTakeover[] = [];
  for (const record of pending) {
    const before = current?.children ?? [];
    if (record.kind === 'paragraphOp') current = withPendingParagraphOp(ctx, current, record);
    else current = withPendingMarkRecord(ctx, current, record, own);
    noteTakeover(ctx, taken, record, before, current?.children ?? []);
  }

  // A decided change made after a pending one shows on top; the pending record stays
  const keepRecord = (c: OoxmlNode) => keptByMark(c) || isWml(c, 'pPrChange');
  for (const record of later) current = withDecidedMarkRecord(ctx, current, record, keepRecord);
  ctx.keepTakeovers(current?.children ?? [], taken);
  return current && current.children.length === 0 && !base ? null : current;
}

/** The children of a mark's `w:pPr` that its property records leave as they are. */
const keptByMark = (c: OoxmlNode): boolean =>
  isWml(c, 'rPr') || c.kind === 'runProperties' || isWml(c, 'sectPr');

/** A mark's `w:pPr` with these children: `current`, or a new one when it has none. */
function withMarkChildren(
  ctx: ComposeContext,
  current: OoxmlElement | null,
  children: readonly OoxmlNode[]
): OoxmlElement {
  return withChildren(current ?? build(ctx.mint(), 'paragraphProperties', 'pPr', [], []), children);
}

/** A mark's `w:pPr` after a decided record: its press run untracked, or its properties. */
function withDecidedMarkRecord(
  ctx: ComposeContext,
  current: OoxmlElement | null,
  record: OrderedElementRecord | OrderedParagraphOpRecord,
  keep: (c: OoxmlNode) => boolean
): OoxmlElement | null {
  if (record.kind === 'paragraphOp')
    return withParagraphOp(ctx.story.part, current, record.op, ctx.mint, null);
  return withMarkChildren(ctx, current, applyProperties(record, current?.children ?? [], keep));
}

/**
 * A mark's `w:pPr` after a pending press: the engine runs it, tracked, and keeps one record (the
 * original, under the latest author).
 */
function withPendingParagraphOp(
  ctx: ComposeContext,
  current: OoxmlElement | null,
  record: OrderedParagraphOpRecord
): OoxmlElement | null {
  const tracked = record.untracked
    ? null
    : {
        revision: ctx.attribution(record.suggestion),
        revisionId: ctx.revisionId(record.suggestion, REVISION_ROLE.paragraphProperties),
      };
  const next = withParagraphOp(ctx.story.part, current, record.op, ctx.mint, tracked);
  if (!next || !tracked) return next;
  const role = REVISION_ROLE.paragraphProperties;
  return withChildren(next, ctx.withRecordAttributes(next.children, record.suggestion, role));
}

/**
 * A mark's `w:pPr` after a pending property record, with its change record: the file's own
 * (`own`) when it is the one record drawn, else the engine's.
 */
function withPendingMarkRecord(
  ctx: ComposeContext,
  current: OoxmlElement | null,
  record: OrderedElementRecord,
  own: OoxmlElement | undefined
): OoxmlElement {
  const prior = current?.children ?? [];
  const next = applyProperties(record, prior, keptByMark);
  const given = next.filter((c) => !keptByMark(c));
  if (own) return withMarkChildren(ctx, current, [...given, ...next.filter(keptByMark), own]);
  const recorded = ctx.withRecordAttributes(
    withPropertyChangeRecord({
      container: 'paragraphProperties',
      prior: prior.filter((c) => !keptByMark(c)),
      next: given,
      revision: ctx.attribution(record.suggestion),
      mint: ctx.mint,
      nextRevisionId: () => ctx.revisionId(record.suggestion, REVISION_ROLE.paragraphProperties),
      recordUnchanged: record.recordUnchanged === true,
    }),
    record.suggestion,
    REVISION_ROLE.paragraphProperties
  );
  return withMarkChildren(ctx, current, [
    ...recorded.filter((c) => !isWml(c, 'pPrChange')),
    ...next.filter(keptByMark),
    ...recorded.filter((c) => isWml(c, 'pPrChange')),
  ]);
}

/**
 * A revision element's name with its `w:id`, `w:author` and `w:date`, the triple the engine
 * decides it by.
 */
const recordKey = (node: OoxmlElement): string =>
  [
    node.localName,
    ...['id', 'author', 'date'].map(
      (name) =>
        node.attributes.find((a) => a.namespaceUri === WML_NAMESPACE_URI && a.localName === name)
          ?.value ?? ''
    ),
  ].join('\u0000');

/**
 * A paragraph's `w:pPr` with its file's own numbering records (`load/properties.ts`) as their
 * suggestions stand: a decision either way removes the record from `w:numPr` and keeps the list,
 * as the engine's and Word's do; a pending one takes its suggestion's attribution (`fileMark`).
 * A copy in a change record stays, as the engine leaves it.
 */
export function withNumberingRecords(
  ctx: ComposeContext,
  mark: OrderedItem,
  pPr: OoxmlElement | null
): OoxmlElement | null {
  const suggestions = new Map(
    mark.props.flatMap((p) =>
      p.kind === 'element' && p.target === 'numPr' && p.fileRecord !== undefined
        ? [[recordKey(p.fileRecord), p.suggestion] as const]
        : []
    )
  );
  const numPr = pPr?.children.find((c) => isWml(c, 'numPr'));
  if (!pPr || suggestions.size === 0 || !numPr || numPr.kind === 'textValue') {
    return pPr;
  }

  const drawn = numPr.children.flatMap((c) => {
    if (c.kind === 'textValue' || !(isWml(c, 'ins') || isWml(c, 'numberingChange'))) {
      return [c];
    }

    const suggestion = suggestions.get(recordKey(c));
    if (suggestion === undefined) {
      return [c];
    }

    const decided = ctx.options.pending === 'reject' || statusOf(ctx.doc, suggestion) !== 'pending';
    return decided ? [] : [fileMark(ctx.doc, c, suggestion)];
  });
  if (drawn.length === numPr.children.length && drawn.every((c, k) => c === numPr.children[k])) {
    return pPr;
  }

  return withChildren(
    pPr,
    pPr.children.map((c) => (c === numPr ? withChildren(numPr, drawn) : c))
  );
}

/** A `w:tblGridChange` with its `w:id` only (`CT_Markup`); any other node as it is. */
const asMarkup = (node: OoxmlNode): OoxmlNode =>
  node.kind !== 'textValue' && isWml(node, 'tblGridChange')
    ? withAttributes(
        node,
        node.attributes.filter((a) => a.localName === 'id')
      )
    : node;

/** The properties children of a container's open item that records target. */
type ContainerTarget = Exclude<OrderedPropertyTarget, 'pPr' | 'numPr' | 'sectPr'>;

/**
 * The engine's change record container of each target. A control's value has none; a row's
 * exceptions have only the file's own record, written back as the file had it.
 */
const CONTAINERS: { readonly [T in ContainerTarget]: PropertyChangeContainer | null } = {
  tblPr: 'tableProperties',
  tblGrid: 'tableGrid',
  tblPrEx: null,
  trPr: 'rowProperties',
  tcPr: 'cellProperties',
  sdtPr: null,
};

const KINDS: { readonly [T in ContainerTarget]: OoxmlElement['kind'] } = {
  tblPr: 'tableProperties',
  tblGrid: 'tableGrid',
  tblPrEx: 'generic',
  trPr: 'generic',
  tcPr: 'generic',
  sdtPr: 'contentControlProperties',
};

const isTableTarget = (name: string): name is ContainerTarget => name in CONTAINERS;

const elementRecords = (records: readonly OrderedPropertyRecord[]): OrderedElementRecord[] =>
  records.filter((record): record is OrderedElementRecord => record.kind === 'element');

/**
 * Whether a suggestion's only record is a table's grid: a drag that changed no cell's width,
 * which the engine records as a lone `w:tblGridChange`. Word keeps the new grid when it rejects
 * one (Word for Mac, 2026-10-29: `~$word-lone-grid-reject.docx`), as the engine does
 * (`unboundTableHistories`), so a reject draws it as accepted.
 */
function isLoneGrid(ctx: ComposeContext, open: OrderedItem, suggestion: string): boolean {
  const items = ctx.doc.labelled.get(suggestion);
  if (items?.size !== 1 || !items.has(open) || open.by === suggestion) return false;
  const records = open.props.filter((p) => p.suggestion === suggestion);
  const [only] = records;
  return (
    records.length === 1 &&
    only?.kind === 'element' &&
    only.target === 'tblGrid' &&
    only.untracked !== true &&
    !open.dels.includes(suggestion)
  );
}

/** A cell's `w:tcW` among properties children, with its type and integral width. */
function cellWidthOf(
  children: readonly OoxmlNode[]
): { readonly node: OoxmlElement; readonly type: string; readonly value: number } | undefined {
  const node = children.find((c): c is OoxmlElement => isWml(c, 'tcW'));
  const raw = node && attributeOf(node, 'w');
  const type = node && attributeOf(node, 'type');
  if (!node || (type !== 'dxa' && type !== 'pct') || !/^\d{1,9}$/.test(raw ?? '')) return undefined;
  return { node, type, value: Number(raw) };
}

/**
 * A cell record's original with the width a column lent after the record was taken lent from it
 * too (the engine's `withRecordLending`): Word writes no record for a lent width, so a reject of
 * the earlier change keeps the column whole.
 */
function withLentWidth(
  old: readonly OoxmlNode[],
  before: readonly OoxmlNode[],
  after: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  const [was, now, recorded] = [cellWidthOf(before), cellWidthOf(after), cellWidthOf(old)];
  if (!was || !now || !recorded || was.type !== now.type || recorded.type !== now.type) return old;
  const value = recorded.value + now.value - was.value;
  if (value === recorded.value || value <= 0) return old;
  const attributes = recorded.node.attributes.map((a) =>
    a.localName === 'w' ? { ...a, value: String(value) } : a
  );
  return old.map((c) => (c === recorded.node ? withAttributes(recorded.node, attributes) : c));
}

/**
 * An open item's properties children (`w:tblPr`, `w:tblGrid`, `w:trPr`, `w:tcPr`) under their
 * records.
 */
export function containerPropertiesAfter(
  ctx: ComposeContext,
  open: OrderedItem,
  children: readonly OoxmlNode[]
): readonly OoxmlNode[] {
  if (open.props.length === 0) return children;
  const made = new Set<string>();
  const drawn = withMissingContainers(ctx, open, children, made).map((child) =>
    containerAfter(ctx, open, child)
  );
  if (made.size === 0) return drawn;

  // A container made here that no record fills (all decided away) is left out again
  return drawn.filter(
    (child) => !made.has(child.id) || (child.kind !== 'textValue' && child.children.length > 0)
  );
}

/**
 * The children with an empty container, in its place, for each target a record names that the
 * item does not have yet. `made` takes the ids of the containers made.
 */
function withMissingContainers(
  ctx: ComposeContext,
  open: OrderedItem,
  children: readonly OoxmlNode[],
  made: Set<string>
): readonly OoxmlNode[] {
  const targets = new Set(open.props.map((p) => p.target));
  const missing = [...targets].filter(
    (t): t is ContainerTarget => isTableTarget(t) && !children.some((c) => isWml(c, t))
  );
  if (missing.length === 0) return children;
  const container = (t: ContainerTarget) => {
    const element = build(ctx.mint(), KINDS[t], t, [], []);
    made.add(element.id);
    return element;
  };

  // A row's `w:trPr` follows its `w:tblPrEx` (`CT_Row`)
  const after = children.findIndex((c) => isWml(c, 'tblPrEx')) + 1;
  return [
    ...missing.filter((t) => t === 'tblPrEx').map(container),
    ...children.slice(0, after),
    ...missing.filter((t) => t !== 'tblGrid' && t !== 'tblPrEx').map(container),
    ...children.slice(after),
    ...missing.filter((t) => t === 'tblGrid').map(container),
  ];
}

/** Records keep none of a container's children: each gives the whole container. */
const keepNone = (): boolean => false;

/** One properties child of an open item under its records, or the child as it is. */
function containerAfter(ctx: ComposeContext, open: OrderedItem, child: OoxmlNode): OoxmlNode {
  if (child.kind === 'textValue' || !isTableTarget(child.localName)) return child;
  const target = child.localName;
  const records = elementRecords(open.props.filter((p) => p.target === target));
  if (records.length === 0) return child;
  const shown = acceptedAndPending(ctx, records);

  // A pending change Word writes no record for (a column's lent width), made before the first
  // one it records, is part of the original that record holds: a reject of the later change
  // keeps it, as Word's reject of a file with it does
  const firstRecorded = shown.pending.find((p) => p.untracked !== true)?.seq ?? Infinity;
  const unrecorded = shown.pending.filter((p) => p.untracked === true && p.seq < firstRecorded);
  const pending = shown.pending.filter((p) => !unrecorded.includes(p));

  // A direct change shows on top of the pending ones it was made over, and goes with a reject;
  // a rejected lone grid record keeps its grid, and what was made over it
  const loneGrid = (suggestion: string) => isLoneGrid(ctx, open, suggestion);
  const kept = records.filter(
    (record) => ctx.rejectedOrDrawnSo(record.suggestion) && loneGrid(record.suggestion)
  );
  const decided = shown.accepted.filter(
    (record) => !madeOverRejected(ctx, record, records, loneGrid)
  );
  const [earlier, later] = aroundPending([...decided, ...kept].sort(bySeq), pending);
  const under = [...earlier, ...unrecorded].sort(bySeq);
  let inner: readonly OoxmlNode[] = child.children;
  for (const record of under) inner = applyProperties(record, inner, keepNone);
  let old = inner;
  const taken: DrawnTakeover[] = [];
  for (const record of [...pending, ...later]) {
    const before = inner;
    inner = applyProperties(record, inner, keepNone);
    noteTakeover(ctx, taken, record, before, inner);

    // A width lent after it is lent from the record's original too, as the engine lends it;
    // so is a grid a column changed after it (a row height's grid record, Word's form)
    if (record.untracked === true && target === 'tcPr') old = withLentWidth(old, before, inner);
    if (record.untracked === true && target === 'tblGrid')
      old = applyProperties(record, old, keepNone);
  }

  // The record names the author of the latest change it records
  const latest = pending.filter((p) => p.untracked !== true).at(-1);
  const own = fileRecordAlone(ctx.doc, under, pending);
  const container = CONTAINERS[target];
  if (own) inner = [...inner, own];
  else if (latest && container !== null) {
    inner = ctx.withRecordAttributes(
      withPropertyChangeRecord({
        container,
        prior: old,
        next: inner,
        revision: ctx.attribution(latest.suggestion),
        mint: ctx.mint,
        // One table op is one record id, over every container it touched
        nextRevisionId: () => ctx.revisionId(latest.suggestion, REVISION_ROLE.tableProperties),
        recordUnchanged: target === 'tblGrid' || pending.some((p) => p.recordUnchanged === true),
      }),
      latest.suggestion,
      REVISION_ROLE.tableProperties
    );

    // `w:tblGridChange` is `CT_Markup`: its `w:id` and nothing else
    if (target === 'tblGrid') inner = inner.map(asMarkup);
    ctx.keepTakeovers(inner, taken);
  }

  return withChildren(child, inner);
}
