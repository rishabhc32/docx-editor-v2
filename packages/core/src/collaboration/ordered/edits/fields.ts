// Fields and bookmarks (group 9.5): each writes what the engine's own op writes.

import { bookmarkNameKey } from '../../../store/package/bookmark-names.ts';
import { FIELD_ATOM_CHAR, isFldChar, isFldSimpleNode } from '../../../store/package/field-nodes.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { bookmarkPairNodes } from '../../../store/package/toc-build.ts';
import { pageFieldContentBuilders } from '../../../store/store/tree-op-fields.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { refuseFieldInField } from '../field-depth.ts';
import {
  createOrderedItem,
  insertedPending,
  isHidden,
  isPending,
  itemAt,
} from '../model/document.ts';
import { insertItems, splitItem } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedFieldPiece, OrderedItem } from '../model/types.ts';
import { attributeOf, withChildren } from '../ooxml-nodes.ts';
import { firstItemOf, isTableOfContents, refuseLocked } from '../placement.ts';
import { strikeItems, suggestionFor, withNewIds, type Ctx } from './context.ts';
import { directInsert, settle } from './direct.ts';
import { freshDecimal, newMarkerItem } from './marker.ts';
import type { OrderedResolvedPlace } from './places.ts';
import { engineScratch } from './templates.ts';
import { insertText, type Piece } from './text.ts';

/** The page fields the engine types (`insertPageField`). @public */
export type OrderedPageField = 'PAGE' | 'NUMPAGES' | 'SECTIONPAGES' | 'PAGE_X_OF_Y';

const MAX_WALK_DEPTH = 16;

/** The first node in a tree (depth first) that `match` accepts. */
function findIn(node: OoxmlNode, match: (n: OoxmlNode) => boolean, depth = 0): OoxmlNode | null {
  if (match(node)) return node;
  if (node.kind === 'textValue' || depth >= MAX_WALK_DEPTH) return null;
  for (const child of node.children) {
    const found = findIn(child, match, depth + 1);
    if (found) return found;
  }

  return null;
}

const textOf = (node: OoxmlNode): string =>
  node.kind === 'textValue'
    ? node.value
    : node.children.map((c) => (c.kind === 'textValue' ? c.value : '')).join('');

// ---- page fields ----

/**
 * The pieces a page field types: one object for each field, its nodes in one run as the engine
 * builds them, and the literal between the two fields of "Page X of Y".
 */
function pageFieldPieces(ctx: Ctx, field: OrderedPageField): Piece[] {
  const mint = () => ctx.ids.node();
  const out: Piece[] = [];
  let open: OoxmlNode[] | null = null;
  for (const builder of pageFieldContentBuilders(field)) {
    const node = builder(mint);
    if (open === null && isFldChar(node, 'begin')) open = [];
    if (open === null) {
      out.push({ text: textOf(node) });
      continue;
    }

    open.push(node);
    if (!isFldChar(node, 'end')) continue;
    const run = build(ctx.ids.node(), 'run', 'r', [], open);
    out.push({ text: FIELD_ATOM_CHAR, field: [{ node: run, sourceRun: run.id }] });
    open = null;
  }

  return out;
}

/**
 * Type a page field at a place, as the engine's `insertPageField` does: one object (a whole
 * complex field), so a reject takes it all back. Returns the suggestion that holds it, or null
 * for a direct edit (settled where it lands).
 */
export function insertPageField(
  ctx: Ctx,
  at: OrderedResolvedPlace,
  field: OrderedPageField
): string | null {
  const pieces = pageFieldPieces(ctx, field);
  if (ctx.direct) {
    directInsert(ctx, at.paragraph, at.offset, '', pieces);
    return null;
  }

  return insertText(ctx, at.paragraph, at.offset, '', { pieces, outsidePrompt: true });
}

// ---- a field's instruction and result ----

/** The node an op names a whole field by: its begin `w:fldChar`, or its `w:fldSimple`. */
export function fieldNodeIdOf(item: OrderedItem): string | null {
  if (item.kind !== 'atom' || !item.pieces) return null;
  if (item.node && isFldSimpleNode(item.node)) return item.node.id;
  for (const piece of item.pieces) {
    const begin = findIn(piece.node, (n) => isFldChar(n, 'begin'));
    if (begin) return begin.id;
  }

  return null;
}

/**
 * A whole field the edit may rewrite: shown, in no pending insertion, deletion or format (a
 * pending move's source is struck), and not where the engine keeps content from edits. A table
 * of contents' own page pass (`inToc`) rewrites the fields of its generated result, which every
 * other edit leaves alone.
 */
function refuseUnrewritableField(
  ctx: Ctx,
  item: OrderedItem,
  at: OrderedResolvedPlace,
  inToc = false
): void {
  if (fieldNodeIdOf(item) === null) throw new OrderedRefusal('invalid-edit', 'not a whole field');
  if (isHidden(ctx.doc, item)) throw new OrderedRefusal('hidden-target', item.space.slice(0, 160));

  // The engine leaves a field inside `w:ins` or `w:del` as it is: it is revision content. One in
  // a pending move's destination is moved content, which the engine rewrites in place
  if (
    (insertedPending(ctx.doc, item) && !inPendingMoveTo(ctx, item)) ||
    item.dels.some((id) => isPending(ctx.doc, id)) ||
    item.fmts.some((record) => isPending(ctx.doc, record.suggestion))
  )
    throw new OrderedRefusal('invalid-edit', 'a field inside a suggestion');
  if (!(inToc && at.paragraph.blocks.some(isTableOfContents)))
    refuseLocked(ctx, at.paragraph, at.offset, at.offset + 1);
}

/** In a pending move's destination: moved content, inserted by the move's suggestion. */
function inPendingMoveTo(ctx: Ctx, item: OrderedItem): boolean {
  return item.move !== null && insertedPending(ctx.doc, item);
}

/** A field's runs as its paragraph holds them: alone, or in its `w:fldSimple`. */
function fieldChildren(item: OrderedItem, runs: OoxmlNode[]): OoxmlNode[] {
  const shell = item.node;
  return shell && shell.kind !== 'textValue' && isFldSimpleNode(shell)
    ? [withChildren(shell, runs)]
    : runs;
}

/** The field alone in a scratch paragraph. */
function fieldParagraph(ctx: Ctx, item: OrderedItem): OoxmlElement {
  const runs = (item.pieces ?? []).map((piece) => piece.node);
  return build(ctx.ids.node(), 'paragraph', 'p', [], fieldChildren(item, runs));
}

/** A field's new node and pieces, with the edit's ids; a run keeps its file records. */
interface FieldForm {
  readonly node: OoxmlNode | null;
  readonly pieces: readonly OrderedFieldPiece[];
}

function formOf(ctx: Ctx, item: OrderedItem, children: readonly OoxmlNode[]): FieldForm {
  const simple = item.node && isFldSimpleNode(item.node) ? children[0] : undefined;
  const runs = simple && simple.kind !== 'textValue' ? simple.children : children;
  const old = new Map((item.pieces ?? []).map((piece) => [piece.node.id, piece]));
  const pieces = runs.flatMap((node): OrderedFieldPiece[] => {
    if (node.kind === 'textValue') return [];
    const was = old.get(node.id);
    const copy = withNewIds(ctx, node);
    return [
      {
        node: copy,
        ...(copy.kind === 'run' ? { sourceRun: copy.id } : {}),
        ...(was?.recordFaces ? { recordFaces: was.recordFaces } : {}),
        ...(was?.recordFiles ? { recordFiles: was.recordFiles } : {}),
      },
    ];
  });
  const node =
    simple && simple.kind !== 'textValue' ? withNewIds(ctx, withChildren(simple, [])) : item.node;
  return { node, pieces };
}

/**
 * The field as the engine's op leaves it, run on a scratch paragraph that holds only the field;
 * null when the op changes nothing (the engine skips a field it cannot rewrite, or one that
 * shows the text already).
 */
function engineRewrite(
  ctx: Ctx,
  item: OrderedItem,
  op: (paragraphId: string, fieldNodeId: string) => TreeDocOp
): FieldForm | null {
  const fieldNodeId = fieldNodeIdOf(item);
  if (fieldNodeId === null) throw new OrderedRefusal('invalid-edit', 'not a whole field');
  const paragraph = fieldParagraph(ctx, item);
  const scratch = engineScratch(ctx.story.part, [paragraph], op(paragraph.id, fieldNodeId));
  const after = findNode(scratch, paragraph.id);
  if (!after || after.kind === 'textValue')
    throw new OrderedRefusal('engine-refused', `${op(paragraph.id, fieldNodeId).op}: no field`);
  if (after.children === paragraph.children) return null;
  return formOf(ctx, item, after.children);
}

/**
 * Replace a whole field's item with a new form: the old item struck, the new one inserted right
 * after it with the same face, format records and containers. Returns the suggestion. A
 * suggesting change saves as Word's tracked field update (the old runs in `w:del`, the new ones
 * in `w:ins`), and the organiser's undo (`revert`) shows the old field again exactly. In a
 * pending move's destination the new one stays in the move, as the engine's rewrite in place
 * does: it goes with a rejected move, and the organiser's undo deletes it as text typed inside
 * another person's insertion.
 */
function replaceField(ctx: Ctx, item: OrderedItem, form: FieldForm): string {
  const suggestion = suggestionFor(ctx);
  const moved = inPendingMoveTo(ctx, item);
  strikeItems(ctx, ctx.story, [item], suggestion.id);
  const space = ctx.ids.units();
  const copy = createOrderedItem({
    space,
    kind: 'atom',
    text: item.text,
    node: form.node,
    pieces: form.pieces,
    face: item.face,
    run: space,
    by: moved ? item.by : suggestion.id,
    ...(moved ? { move: item.move } : {}),
    fmts: item.fmts,
    within: item.within,
    ...(item.bare ? { bare: true } : {}),
  });
  insertItems(ctx.doc, ctx.story, item, [copy]);
  refuseFieldInField(ctx, [copy]);
  if (ctx.direct) settle(ctx.doc, suggestion.id);
  return suggestion.id;
}

/** A field's new instruction (`PAGE` or `NUMPAGES`), as the engine's `setFieldCode` writes it. */
export function setFieldCode(
  ctx: Ctx,
  item: OrderedItem,
  at: OrderedResolvedPlace,
  code: string
): string | null {
  refuseUnrewritableField(ctx, item, at);
  const form = engineRewrite(ctx, item, (paragraphId, fieldNodeId) => ({
    op: 'setFieldCode',
    paragraphId,
    fieldNodeId,
    code,
  }));
  return form === null ? null : replaceField(ctx, item, form);
}

/** The last `w:t` in a node, in document order. */
function lastTextElement(node: OoxmlNode, depth = 0): OoxmlElement | null {
  if (node.kind === 'textValue' || depth >= MAX_WALK_DEPTH) return null;
  for (let k = node.children.length - 1; k >= 0; k -= 1) {
    const child = node.children[k];
    const found = child ? lastTextElement(child, depth + 1) : null;
    if (found) return found;
  }

  return node.localName === 't' ? node : null;
}

/** A copy of a node with the text of `found` (an element in it) set to `text`. */
function withTextOf(ctx: Ctx, node: OoxmlNode, found: OoxmlElement, text: string): OoxmlNode {
  if (node.kind === 'textValue') {
    return node;
  }

  if (node === found) {
    return withChildren(node, [{ id: ctx.ids.node(), kind: 'textValue', value: text }]);
  }

  return withChildren(
    node,
    node.children.map((child) => withTextOf(ctx, child, found, text))
  );
}

/** The pieces with the last `w:t`'s text set, as the engine rewrites a contents page number. */
function withLastText(ctx: Ctx, item: OrderedItem, text: string): FieldForm | null {
  const pieces = item.pieces ?? [];
  for (let k = pieces.length - 1; k >= 0; k -= 1) {
    const piece = pieces[k];
    const found = piece ? lastTextElement(piece.node) : null;
    if (found === null) {
      continue;
    }

    if (textOf(found) === text) {
      return null;
    }

    const runs = pieces.map((p, at) => (at === k ? withTextOf(ctx, p.node, found, text) : p.node));
    return formOf(ctx, item, fieldChildren(item, runs));
  }

  throw new OrderedRefusal('invalid-edit', 'a field with no text to rewrite');
}

/**
 * A field's new result, as the engine's `refreshFieldResults` writes it; with `last`, only the
 * field's last `w:t`, as `rewriteTocPageNumbers` writes a contents row's page number. The text
 * comes in the change: a copy never computes it again.
 */
export function setFieldResult(
  ctx: Ctx,
  item: OrderedItem,
  at: OrderedResolvedPlace,
  text: string,
  last: boolean
): string | null {
  refuseUnrewritableField(ctx, item, at, last);
  const form = last
    ? withLastText(ctx, item, text)
    : engineRewrite(ctx, item, (paragraphId, fieldNodeId) => ({
        op: 'refreshFieldResults',
        updates: [{ paragraphId, fieldNodeId, text }],
      }));
  return form === null ? null : replaceField(ctx, item, form);
}

/**
 * New text for the units of one text element, its run, face and text form kept: the engine's
 * `rewriteTocPageNumbers` sets the last `w:t` of a contents row in place. The old units are
 * struck and the new ones go right after them, under the same suggestion.
 */
export function setRunText(ctx: Ctx, units: readonly OrderedItem[], text: string): string | null {
  const first = units[0];
  const last = units[units.length - 1];
  if (!first || !last) throw new OrderedRefusal('empty-range');
  if (
    units.some(
      (u) =>
        u.kind !== 'text' ||
        u.instr === true ||
        u.run !== first.run ||
        u.textShell?.id !== first.textShell?.id ||
        isHidden(ctx.doc, u)
    )
  )
    throw new OrderedRefusal('invalid-edit', 'the text of one text element');
  if (units.map((u) => u.text).join('') === text) return null;
  const suggestion = suggestionFor(ctx);
  strikeItems(ctx, ctx.story, units, suggestion.id);
  if (text.length > 0) {
    const space = ctx.ids.units();
    const copy = createOrderedItem({
      space,
      kind: 'text',
      length: text.length,
      text,
      face: first.face,
      run: first.run,
      ...(first.textShell ? { textShell: first.textShell } : {}),
      by: suggestion.id,
      fmts: first.fmts,
      within: first.within,
    });
    insertItems(ctx.doc, ctx.story, last, [copy]);
  }

  if (ctx.direct) settle(ctx.doc, suggestion.id);
  return suggestion.id;
}

// ---- bookmarks ----

/** A bookmark marker's element (`w:bookmarkStart` or `w:bookmarkEnd`), or null. */
export function bookmarkMarker(item: OrderedItem): OoxmlElement | null {
  const node = item.node;
  if (item.kind !== 'zero' || !node || node.kind === 'textValue') return null;
  return node.localName === 'bookmarkStart' || node.localName === 'bookmarkEnd' ? node : null;
}

/**
 * A Word decimal for a new bookmark, past every bookmark id the document holds: minted from the
 * change, so every copy writes the same one in any order.
 */
function bookmarkId(ctx: Ctx): string {
  const used = new Set<string>();
  for (const story of ctx.doc.stories.values())
    for (const item of story.tree.items()) {
      const marker = bookmarkMarker(item);
      const id = marker ? attributeOf(marker, 'id') : undefined;
      if (id !== undefined) used.add(id);
    }

  return freshDecimal(ctx, (id) => used.has(id));
}

/**
 * The item a marker at a place goes right after, as the engine's `markerSiteAt` places it: past
 * the unit before the place and every container that ends there, before any marker after them.
 */
function markerAnchor(ctx: Ctx, at: OrderedResolvedPlace): OrderedItem | null {
  const tree = ctx.story.tree;
  if (at.offset === 0) return tree.prev(firstItemOf(at.paragraph));
  const char = at.paragraph.chars[at.offset - 1];
  if (!char) throw new OrderedRefusal('unknown-address', `offset ${at.offset}`);
  const item = itemAt(ctx.story, char.address);
  if (!item) throw new OrderedRefusal('unknown-address', char.address.space.slice(0, 160));
  const unit = splitItem(ctx.doc, ctx.story, item, char.address.offset + 1)?.[0] ?? item;
  let anchor = unit;
  for (let next = tree.next(unit); next; next = tree.next(next)) {
    if (isHidden(ctx.doc, next)) continue;
    if (next.kind === 'close') anchor = next;
    else if (next.kind !== 'zero') break;
  }

  return anchor;
}

/**
 * The ids of the shown bookmark starts of a name in the story: a removed end is one of theirs,
 * so a change never takes away another bookmark's end, as the editor names the ends by them.
 */
function endIdsOf(ctx: Ctx, key: string): Set<string> {
  const ids = new Set<string>();
  for (const item of ctx.story.tree.items()) {
    const marker = bookmarkMarker(item);
    if (marker?.localName !== 'bookmarkStart' || isHidden(ctx.doc, item)) continue;
    const id = attributeOf(marker, 'id');
    if (id !== undefined && bookmarkNameKey(attributeOf(marker, 'name') ?? '') === key) ids.add(id);
  }

  return ids;
}

/** Refuse a removal of a marker that is not the bookmark's own, or that does not show. */
function checkRemovals(ctx: Ctx, key: string, name: string, remove: readonly OrderedItem[]): void {
  const ends = remove.some((item) => bookmarkMarker(item)?.localName === 'bookmarkEnd')
    ? endIdsOf(ctx, key)
    : new Set<string>();
  for (const item of remove) {
    const marker = bookmarkMarker(item);
    const own =
      marker?.localName === 'bookmarkStart'
        ? bookmarkNameKey(attributeOf(marker, 'name') ?? '') === key
        : marker !== null && ends.has(attributeOf(marker, 'id') ?? '');
    if (!marker || !own)
      throw new OrderedRefusal('invalid-edit', `not a marker of the bookmark ${name.slice(0, 64)}`);
    if (isHidden(ctx.doc, item))
      throw new OrderedRefusal('hidden-target', item.space.slice(0, 160));
  }
}

/**
 * Add a bookmark named `name` from one place to another (equal places: a point), as the
 * engine's `insertBookmark` writes it, and take away the markers `remove` names (the same name,
 * elsewhere, which the new one moves; or a bookmark deleted). A direct edit only: a bookmark
 * carries no author or date (ECMA-376 `CT_Bookmark`, `CT_MarkupRange`), so Word writes it
 * directly with Track Changes on.
 */
export function editBookmark(
  ctx: Ctx,
  name: string,
  range: { readonly from: OrderedResolvedPlace; readonly to: OrderedResolvedPlace } | null,
  remove: readonly OrderedItem[]
): string {
  if (!ctx.direct)
    throw new OrderedRefusal('invalid-edit', 'a bookmark is no revision: a direct edit only');
  const key = bookmarkNameKey(name);
  checkRemovals(ctx, key, name, remove);
  const suggestion = suggestionFor(ctx);
  strikeItems(ctx, ctx.story, remove, suggestion.id);
  if (range !== null) {
    refuseLocked(ctx, range.from.paragraph, range.from.offset, range.from.offset);
    refuseLocked(ctx, range.to.paragraph, range.to.offset, range.to.offset);
    const pair = bookmarkPairNodes(() => ctx.ids.node(), name, bookmarkId(ctx));

    // Both sites are read before either marker goes in: at one point the start then lands in
    // front of the end, the order a point bookmark is written in
    const startAfter = markerAnchor(ctx, range.from);
    const endAfter = markerAnchor(ctx, range.to);
    const by = { by: suggestion.id };
    insertItems(ctx.doc, ctx.story, endAfter, [
      newMarkerItem(ctx, range.to.paragraph, pair.end, by),
    ]);
    insertItems(ctx.doc, ctx.story, startAfter, [
      newMarkerItem(ctx, range.from.paragraph, pair.start, by),
    ]);
  }

  settle(ctx.doc, suggestion.id);
  return suggestion.id;
}

/**
 * A bookmark around a whole paragraph, as the engine adds a contents row's heading bookmark:
 * its start before everything in the paragraph, its end right before the paragraph's mark.
 */
export function bookmarkParagraph(
  ctx: Ctx,
  at: OrderedResolvedPlace,
  name: string,
  suggestion: string
): void {
  const p = at.paragraph;
  refuseLocked(ctx, p, 0, p.chars.length);
  const pair = bookmarkPairNodes(() => ctx.ids.node(), name, bookmarkId(ctx));
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(firstItemOf(p)), [
    newMarkerItem(ctx, p, pair.start, { by: suggestion }),
  ]);
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(p.mark), [
    newMarkerItem(ctx, p, pair.end, { by: suggestion }),
  ]);
}
