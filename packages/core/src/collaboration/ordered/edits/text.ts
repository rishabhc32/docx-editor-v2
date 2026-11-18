// Typed text, struck text and formatted text. No item moves: an insertion adds items at the
// place the engine's rules give (`placement.ts`), a deletion and a format add labels.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { keepsPropertyChildren } from '../../../store/store/tree-op-properties.ts';
import type { OrderedViewChar, OrderedViewParagraph } from '../compose/context.ts';
import { faceWithProperties } from '../compose/properties.ts';
import { createOrderedItem, insertedPending, itemAt, madeBy, statusOf } from '../model/document.ts';
import { isolateUnits, nextSeq, recordAdded, relabel, replaceItems } from '../model/mutations.ts';
import type {
  OrderedFieldPiece,
  OrderedFormatRecord,
  OrderedFormatWrite,
  OrderedItem,
  OrderedItemLabels,
  OrderedPropertyRecord,
  OrderedStory,
  OrderedSuggestion,
  OrderedUnitRange,
} from '../model/types.ts';
import { isWml, withChildren } from '../ooxml-nodes.ts';
import {
  inRunMarkers,
  insertAtSite,
  promptSite,
  refuseLocked,
  siteAt,
  withComplexScript,
  type Site,
} from '../placement.ts';
import {
  currentFace,
  pendingOwn,
  recordsFormat,
  shownFace,
  shownParagraphProperties,
  suggestionFor,
  takeBack,
  withNewIds,
  type Ctx,
} from './context.ts';
import { strikeStruckNoteBodies, takeBackNotesOf } from './note-body.ts';

/**
 * Text or one object, with its own face: a move's copy, a note reference, a new object or a
 * field.
 */
export interface Piece {
  /** The text, or an object's one model character. */
  readonly text: string;
  readonly face?: OoxmlElement | null;
  /** The object's node, when the piece is one object. */
  readonly content?: OoxmlNode;
  /** The object sits in the paragraph itself, not in a run (an equation). */
  readonly bare?: boolean;
  /** A whole field's runs, without run properties: they take the piece's face (group 9.5). */
  readonly field?: readonly OrderedFieldPiece[];
}

/** A new field's runs under the face its item takes: each run holds its own copy of it. */
function fieldPiecesWithFace(
  ctx: Ctx,
  pieces: readonly OrderedFieldPiece[],
  face: OoxmlElement | null
): OrderedFieldPiece[] {
  return pieces.map((piece) =>
    face && piece.node.kind === 'run'
      ? {
          ...piece,
          node: withChildren(piece.node, [withNewIds(ctx, face), ...piece.node.children]),
        }
      : piece
  );
}

export interface InsertOptions {
  /** The face the new text takes, in place of the one the placement rules give. */
  readonly face?: OoxmlElement | null;
  /** Pending formats that cover the new text too (a direct edit under a pending format). */
  readonly fmts?: readonly OrderedFormatRecord[];
  /** Inside a pending container of its suggestion: the container's marker covers it. */
  readonly quiet?: boolean;
  /** Never into this suggestion: the text goes on the base side, outside it. */
  readonly outsideOf?: string;
  /** The destination of a move with this name. */
  readonly move?: string;
  /** Several pieces, each with its own face or object (a move's copy, a note reference). */
  readonly pieces?: readonly Piece[];
  /** A tab or a break: it never goes into a control's prompt (rule 7 is for typed text). */
  readonly outsidePrompt?: boolean;
}

/** Units of one space, in a row. */
interface Group {
  readonly space: string;
  readonly start: number;
  end: number;
}

export function groupBySpace(chars: readonly OrderedViewChar[]): Group[] {
  const out: Group[] = [];
  for (const char of chars) {
    const last = out[out.length - 1];
    if (last && last.space === char.address.space && last.end === char.address.offset)
      last.end += 1;
    else
      out.push({
        space: char.address.space,
        start: char.address.offset,
        end: char.address.offset + 1,
      });
  }

  return out;
}

// ---- insertions ----

/** Insert text (or the pieces) at a view offset. Returns the suggestion that holds it. */
export function insertText(
  ctx: Ctx,
  p: OrderedViewParagraph,
  offset: number,
  text: string,
  options: InsertOptions = {}
): string {
  refuseLocked(ctx, p, offset, offset);
  const site =
    (options.outsidePrompt ? null : promptSite(ctx, p, offset)) ?? siteAt(ctx, p, offset);
  const owner = site.neighbour ? pendingOwn(ctx, site.neighbour.item.by) : null;

  // Typing inside or beside your own pending text extends that suggestion
  const extend = owner && owner.id !== options.outsideOf ? owner : null;
  const suggestion = suggestionFor(ctx, extend);
  const face = faceFor(ctx, site, options);
  const grown =
    options.pieces === undefined ? growTyped(ctx, site, suggestion, face, text, options) : null;
  if (grown) {
    noteOwnAddition(ctx, suggestion.id, [grown]);
    return suggestion.id;
  }

  const space = ctx.ids.units();
  const pieces: readonly Piece[] = options.pieces ?? [{ text }];
  let at = 0;
  const items = pieces.map((piece) => {
    const atom = piece.content !== undefined || piece.field !== undefined;
    const length = atom ? 1 : piece.text.length;
    const itemFace = piece.face !== undefined ? piece.face : face;
    const item = createOrderedItem({
      space,
      offset: at,
      length,
      kind: atom ? 'atom' : 'text',
      text: atom ? piece.text || '￼' : piece.text,
      node: piece.content ?? null,
      ...(piece.field ? { pieces: fieldPiecesWithFace(ctx, piece.field, itemFace) } : {}),
      face: itemFace,
      // One run per piece: pieces can differ in face
      run: at === 0 ? space : `${space}~${at}`,
      by: suggestion.id,
      // The rows and cells around the paragraph hold the new text too
      within: p.mark.within,
      ...(options.fmts ? { fmts: options.fmts } : {}),
      ...(options.move === undefined ? {} : { move: options.move }),
      ...(options.quiet ? { quiet: true } : {}),
      ...(piece.bare && piece.content ? { bare: true } : {}),
    });
    at += length;
    return item;
  });
  insertAtSite(ctx, site, items);
  noteOwnAddition(ctx, suggestion.id, [{ story: ctx.story.name, space, from: 0, to: at }]);
  return suggestion.id;
}

/** The face new text takes: the edit's own, the site's, or the neighbour's as it shows now. */
function faceFor(ctx: Ctx, site: Site, options: InsertOptions): OoxmlElement | null {
  if (options.face !== undefined) {
    return options.face;
  }

  if (site.face !== undefined) {
    return site.face;
  }

  return site.neighbour ? currentFace(ctx, site.neighbour.item) : null;
}

/**
 * Units an edit added to a pending suggestion of the author's that an earlier change made
 * (typing beside or over one's own words, Enter or a new note inside them): noted, so the undo
 * of this change takes back only them.
 */
export function noteOwnAddition(
  ctx: Ctx,
  suggestion: string,
  ranges: readonly OrderedUnitRange[]
): void {
  if (!suggestion.startsWith(`S:${ctx.ids.changeId}.`))
    recordAdded(ctx.doc, ctx.ids.changeId, ranges);
}

/** The units of new items of a story. */
export function unitsOf(story: OrderedStory, items: readonly OrderedItem[]): OrderedUnitRange[] {
  return items.map((item) => ({
    story: story.name,
    space: item.space,
    from: item.offset,
    to: item.offset + item.length,
  }));
}

/** A face's properties as a string to compare, ids aside. */
function faceKey(node: OoxmlNode | null): string {
  if (node === null) {
    return '';
  }

  if (node.kind === 'textValue') {
    return JSON.stringify(node.value);
  }

  const attributes = node.attributes.map((x) => `${x.localName}=${x.value}`).join(',');
  const children = node.children.map(faceKey).join(',');
  return `${node.namespaceUri}:${node.localName}(${attributes})[${children}]`;
}

/** Two faces with the same properties, ids aside. */
export function sameFace(a: OoxmlElement | null, b: OoxmlElement | null): boolean {
  const strip = (node: OoxmlElement | null) => (node && node.children.length === 0 ? null : node);
  return faceKey(strip(a)) === faceKey(strip(b));
}

/**
 * Text typed right after the last unit of your own pending typed text, with its face and
 * nothing between: that item takes it (the next offsets of its space), so a run of keystrokes
 * is one item. Returns the units it added, or null when it does not apply.
 */
function growTyped(
  ctx: Ctx,
  site: Site,
  suggestion: OrderedSuggestion,
  face: OoxmlElement | null,
  text: string,
  options: InsertOptions
): OrderedUnitRange | null {
  if (!site.after || text.length === 0) return null;
  if (options.move !== undefined || options.fmts || options.quiet) return null;
  const last = itemAt(ctx.story, { space: site.after.item.space, offset: site.after.offset });
  if (!last || last.kind !== 'text' || !last.space.startsWith('i:')) return null;
  if (last.by !== suggestion.id || last.move !== null || last.fmts.length > 0) return null;
  if (last.dels.length > 0 || last.retracted || last.quiet) return null;
  if (site.after.offset !== last.offset + last.length - 1) return null;
  const spaceItems = ctx.story.spaces.get(last.space) ?? [];
  if (spaceItems[spaceItems.length - 1] !== last) return null;

  // Nothing the new text would go past may follow (a bookmark marker, a run's marker)
  const next = ctx.story.tree.next(last);
  if (site.pastMarkers && next && next.kind === 'zero' && !next.block) return null;
  if (!sameFace(face, last.face)) return null;
  replaceItems(
    ctx.doc,
    ctx.story,
    [last],
    [{ ...last, length: last.length + text.length, text: last.text + text, leaf: undefined }]
  );
  const from = last.offset + last.length;
  return { story: ctx.story.name, space: last.space, from, to: from + text.length };
}

// ---- deletions ----

/**
 * Policy 1 (design D5): a pending deletion already strikes the unit. Word keeps one `w:del` per
 * run.
 */
export function struckAlready(ctx: Ctx, dels: readonly string[]): boolean {
  return dels.some((deleter) => statusOf(ctx.doc, deleter) === 'pending');
}

/**
 * Strike the view characters `[start, end)` of a paragraph. Your own pending text is taken
 * back; a prompt is not content (rule 7). Returns the suggestion that holds the edit, if any.
 */
export function deleteText(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  into: OrderedSuggestion | null,
  options: { readonly evenStruck?: boolean } = {}
): string | null {
  refuseLocked(ctx, p, start, end);
  const chars = p.chars.slice(start, end);
  let suggestion = into;
  let result: string | null = into?.id ?? null;
  const struck: OrderedItem[] = [];
  const taken: OrderedItem[] = [];
  for (const group of groupBySpace(chars)) {
    for (const item of isolateUnits(ctx.doc, ctx.story, group.space, group.start, group.end)) {
      if (item.prompt) continue;
      const own = pendingOwn(ctx, item.by);
      if (own && insertedPending(ctx.doc, item)) {
        takeBack(ctx, ctx.story, item);
        taken.push(item);
        result ??= own.id;
        continue;
      }

      if (!options.evenStruck && struckAlready(ctx, item.dels)) continue;
      suggestion ??= suggestionFor(ctx);
      relabel(ctx.doc, ctx.story, item, { dels: [...item.dels, suggestion.id] });
      struck.push(item);
      result = suggestion.id;
    }
  }

  // A marker inside a run between struck units is struck with them (rule 3)
  if (suggestion) {
    const id = suggestion.id;
    for (const marker of inRunMarkers(ctx, chars, (x) => x.dels.includes(id), false))
      relabel(ctx.doc, ctx.story, marker, { dels: [...marker.dels, id] });

    // A note whose references are all struck now: its text too, as Word writes it
    strikeStruckNoteBodies(ctx, struck, id);
  }

  // A new note whose own reference the author took back goes with it
  takeBackNotesOf(ctx, taken);
  return result;
}

// ---- formats ----

/** The pending suggestions of these records, each once. */
function pendingSuggestionsOf(
  ctx: Ctx,
  records: readonly (OrderedFormatRecord | OrderedPropertyRecord)[]
): string[] {
  const found = new Set<string>();
  for (const record of records)
    if (statusOf(ctx.doc, record.suggestion) === 'pending') found.add(record.suggestion);
  return [...found];
}

/** Whether properties hold their change record (`w:rPrChange`, `w:pPrChange`). */
const holdsChangeRecord = (properties: OoxmlElement | null, name: string): boolean =>
  properties?.children.some((child) => isWml(child, name)) === true;

/**
 * A format or property record Word writes a change record for: not one shown with none
 * (`untracked`: a format on the author's own pending words, a column's lent width), nor one
 * applied only once accepted.
 */
const writesRecord = (record: OrderedFormatRecord | OrderedPropertyRecord): boolean =>
  !('untracked' in record && record.untracked === true) &&
  !('acceptedOnly' in record && record.acceptedOnly === true);

/**
 * Policy 2 (design D5): `kept`, a pending record of another suggestion, taken over by `by`. It
 * keeps its id, so runs keyed by record keep their faces. The takeover is reported to the host
 * once.
 */
export function takenOver<T extends OrderedFormatRecord | OrderedPropertyRecord>(
  ctx: Ctx,
  kept: T,
  by: string
): T {
  if (!ctx.replaced.some((r) => r.suggestion === kept.suggestion && r.by === by)) {
    ctx.replaced.push({ suggestion: kept.suggestion, by });
  }

  return { ...kept, suggestion: by, replaced: kept.suggestion };
}

/**
 * Records with `record` added (policy 2): it takes over every pending record of another
 * suggestion, except the ones `stays` keeps as they are.
 */
export function withTakeOver<T extends OrderedFormatRecord | OrderedPropertyRecord>(
  ctx: Ctx,
  records: readonly T[],
  record: T,
  stays: (kept: T) => boolean
): T[] {
  const kept = records.map((other) =>
    other.suggestion === record.suggestion ||
    statusOf(ctx.doc, other.suggestion) !== 'pending' ||
    stays(other)
      ? other
      : takenOver(ctx, other, record.suggestion)
  );
  return [...kept, record];
}

/**
 * A unit's format records with one more. Policy 2: the new record takes over the pending
 * records of other suggestions, as Word keeps one `w:rPrChange` per run, the original under the
 * latest author.
 */
export function withFormat(
  ctx: Ctx,
  item: OrderedItem,
  record: OrderedFormatRecord
): OrderedFormatRecord[] {
  const fmts = item.fmts;

  // A direct edit changes the text under a pending format and leaves the format record; it
  // names the pending formats whose change record it was made over: their reject takes it away
  if (ctx.direct) {
    const recorded = holdsChangeRecord(shownFace(ctx, item), 'rPrChange');
    const over = recorded ? pendingSuggestionsOf(ctx, fmts.filter(writesRecord)) : [];
    return [
      ...fmts,
      over.length > 0 && record.kind === 'properties' ? { ...record, over } : record,
    ];
  }

  return withTakeOver(ctx, fmts, record, (kept) => !writesRecord(kept));
}

/**
 * Policy 2 for property records: a change of a paragraph's (or a table's, a row's, a cell's)
 * properties takes over the pending records of other suggestions on the same properties, as
 * Word keeps one change record per properties element. The takeovers are reported.
 */
export function withPropertyRecord(
  ctx: Ctx,
  item: OrderedItem,
  record: OrderedPropertyRecord
): OrderedPropertyRecord[] {
  const props = item.props;

  // A direct edit changes the properties under a pending record and leaves the record, and
  // names the pending records whose change record it was made over: their reject takes it away
  // (`over`). A paragraph press that changes nothing writes no record. A change Word records
  // nowhere (a column's lent width) leaves the record too
  if (ctx.direct) {
    const pending = props.filter((kept) => kept.target === record.target && writesRecord(kept));
    const shows =
      record.target !== 'pPr' ||
      holdsChangeRecord(shownParagraphProperties(ctx, item), 'pPrChange');
    const over = shows ? pendingSuggestionsOf(ctx, pending) : [];
    const namesOver = record.kind === 'element' || record.kind === 'paragraphOp';
    return [...props, over.length > 0 && namesOver ? { ...record, over } : record];
  }

  if (record.kind === 'element' && record.untracked === true) return [...props, record];
  return withTakeOver(
    ctx,
    props,
    record,
    (kept) => kept.target !== record.target || !writesRecord(kept)
  );
}

/**
 * A unit's format records with a press's record. A tracked press that gives the unit back the
 * face it had before its pending records cancels them all when the press's author wrote the
 * latest of them: the engine then drops the run's one change record, as Word does, so no
 * suggestion holds the unit any more (another author's record stays, the engine's rule). They are
 * kept apart, so a take-back of the press gives them back (`retract.ts`).
 */
export function withPressFormat(
  ctx: Ctx,
  item: OrderedItem,
  record: OrderedFormatRecord
): Partial<OrderedItemLabels> {
  const fmts = withFormat(ctx, item, record);
  const pending = (f: OrderedFormatRecord) => statusOf(ctx.doc, f.suggestion) === 'pending';
  const latest = item.fmts
    .filter(pending)
    .sort((a, b) => a.seq - b.seq)
    .at(-1);
  const owned = latest !== undefined && madeBy(ctx.doc, latest.suggestion, ctx.attribution);
  if (ctx.direct || !owned || recordsFormat(ctx, fmts, item.face)) return { fmts };
  return {
    fmts: fmts.filter((f) => !pending(f)),
    cancelledFmts: [...item.cancelledFmts, ...fmts.filter(pending)],
  };
}

/**
 * Whether a format press leaves a unit as it is, by the engine's rule: the press restates every
 * value the unit shows, and no change record is there to take over (a tracked press on a unit
 * with a pending record writes the record again under its own author, as Word keeps one record
 * per run, the original under the latest author). Such a press makes no change and no record.
 */
export function formatKeeps(
  ctx: Ctx,
  item: OrderedItem,
  face: OoxmlElement | null,
  write: OrderedFormatWrite
): boolean {
  if (!ctx.direct && item.fmts.some((f) => statusOf(ctx.doc, f.suggestion) === 'pending'))
    return false;
  const children = face?.children ?? [];
  let scratch = 0;
  const merged = faceWithProperties(children, write, () => `keeps#${(scratch += 1)}`);
  return keepsPropertyChildren(children, merged);
}

/** A format record of a suggestion's press, with its own id. */
function formatRecord(
  suggestion: string,
  seq: number,
  write: OrderedFormatWrite
): Extract<OrderedFormatRecord, { kind: 'properties' }> {
  return { kind: 'properties', id: `f:${suggestion}#${seq}`, suggestion, seq, ...write };
}

/**
 * A format on the author's own pending words: a record of their own suggestion that shows with
 * no change record (`untracked`), as the engine formats its own insertion. It names the change
 * that made it, so that change's undo takes it off and leaves the words.
 */
export function ownFormat(
  ctx: Ctx,
  own: OrderedSuggestion,
  seq: number,
  write: OrderedFormatWrite
): OrderedFormatRecord {
  return { ...formatRecord(own.id, seq, write), untracked: true, change: ctx.ids.changeId };
}

/**
 * Format the view characters `[start, end)` of a paragraph. Your own pending text takes a
 * format of your own suggestion with no change record (`ownFormat`); text the press leaves as it
 * is (`formatKeeps`) takes no record. Returns the suggestion that holds the edit, if any.
 */
export function formatText(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  end: number,
  write: OrderedFormatWrite,
  into: OrderedSuggestion | null
): string | null {
  refuseLocked(ctx, p, start, end, write.properties);
  const chars = p.chars.slice(start, end);
  let suggestion = into;
  let result: string | null = into?.id ?? null;
  const seq = nextSeq(ctx.doc);
  for (const group of groupBySpace(chars)) {
    for (const item of isolateUnits(ctx.doc, ctx.story, group.space, group.start, group.end)) {
      // A right-to-left run takes the complex-script partner too (rule 10)
      const face = currentFace(ctx, item, false);
      const written = { ...write, properties: withComplexScript(face, write.properties) };
      if (formatKeeps(ctx, item, face, written)) continue;
      const own = pendingOwn(ctx, item.by);
      if (own && insertedPending(ctx.doc, item)) {
        relabel(ctx.doc, ctx.story, item, {
          fmts: [...item.fmts, ownFormat(ctx, own, seq, written)],
        });
        result ??= own.id;
        continue;
      }

      suggestion ??= suggestionFor(ctx);
      const record = formatRecord(suggestion.id, seq, written);
      relabel(ctx.doc, ctx.story, item, withPressFormat(ctx, item, record));
      result = suggestion.id;
    }
  }

  if (suggestion && chars.length > 0) {
    const id = suggestion.id;
    const record = formatRecord(id, seq, write);
    for (const marker of inRunMarkers(
      ctx,
      chars,
      (x) => x.fmts.some((f) => f.suggestion === id),
      true
    ))
      relabel(ctx.doc, ctx.story, marker, { fmts: withFormat(ctx, marker, record) });
  }

  return result;
}
