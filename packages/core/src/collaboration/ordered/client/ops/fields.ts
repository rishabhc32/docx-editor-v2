// Field, bookmark and table of contents ops (group 9.5): what the client makes of them.

import { bookmarkNameKey } from '../../../../store/package/bookmark-names.ts';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import {
  detectBodyTocs,
  findDetectedToc,
  tocFieldRange,
} from '../../../../store/package/toc-detect.ts';
import { supportedPageFieldCode } from '../../../../store/store/tree-op-field-code.ts';
import { locateFieldResults } from '../../../../store/store/tree-op-field-results.ts';
import { applyTreeOp } from '../../../../store/store/tree-op-apply.ts';
import { paragraphOffsetIndex } from '../../../../store/store/tree-op-segments.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import type { OrderedChangePlace } from '../../change/types.ts';
import type { OrderedTocBookmark, OrderedTocRow } from '../../edits/toc.ts';
import { bookmarkMarker, fieldNodeIdOf } from '../../edits/fields.ts';
import { isHidden, mainStory } from '../../model/document.ts';
import type { OrderedItem, OrderedStory } from '../../model/types.ts';
import { attributeOf, childWhere, isRunProperties } from '../../ooxml-nodes.ts';
import { addressOf, placeAt, type ListParagraph } from '../paragraphs.ts';
import { resultRunFormatProperties } from '../properties.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import type { OrderedChangeBody, OrderedTranslateBatch } from './types.ts';

type OpOf<K extends TreeDocOp['op']> = Extract<TreeDocOp, { readonly op: K }>;

const MAX_WALK_DEPTH = 16;

const outOfStep = (detail: string) =>
  new OrderedClientRefusal('ordered-refused', 'out-of-step', detail.slice(0, 160));

/** Whether a node or one inside it has this id. */
function holds(node: OoxmlNode, id: string, depth = 0): boolean {
  if (node.id === id) return true;
  if (node.kind === 'textValue' || depth >= MAX_WALK_DEPTH) return false;
  return node.children.some((child) => holds(child, id, depth + 1));
}

/** The shown whole field of a paragraph that the editor names by this node id. */
function fieldAtom(batch: OrderedTranslateBatch, p: ListParagraph, nodeId: string): OrderedItem {
  const found = p.view.items.find(
    (item) => !isHidden(batch.doc, item) && fieldNodeIdOf(item) === nodeId
  );
  if (!found) throw outOfStep(`no field ${nodeId}`);
  return found;
}

/** The shown item of a paragraph whose element has this node id. */
function itemWithNode(batch: OrderedTranslateBatch, p: ListParagraph, nodeId: string): OrderedItem {
  const found = p.view.items.find(
    (item) =>
      !isHidden(batch.doc, item) && item.node?.kind !== 'textValue' && item.node?.id === nodeId
  );
  if (!found) throw outOfStep(`no node ${nodeId}`);
  return found;
}

/**
 * The part the editor shows for a story, as the client last published it from the list: the
 * node ids an op names there are the list's.
 */
function shownPart(batch: OrderedTranslateBatch, story: OrderedStory): OoxmlPart {
  const part = batch.shown()?.parts.get(story.name);
  if (!part) throw outOfStep(`no shown part ${story.name}`);
  return part;
}

/** The shown paragraph element with this node id. */
function shownParagraph(part: OoxmlPart, id: string): OoxmlElement {
  const node = findNode(part, id);
  if (!node || node.kind !== 'paragraph') throw outOfStep(`no shown paragraph ${id}`);
  return node;
}

// ---- fields ----

export function translatePageField(
  op: OpOf<'insertPageField'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const p = batch.paragraph(op.paragraphId);
  return [
    { kind: 'insert-field', part: p.story.name, place: placeAt(p, op.offset), field: op.field },
  ];
}

export function translateFieldCode(
  op: OpOf<'setFieldCode'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const code = supportedPageFieldCode(op.code);
  if (code === null)
    throw new OrderedClientRefusal('ordered-refused', 'invalid-change', 'a PAGE or NUMPAGES code');
  const p = batch.paragraph(op.paragraphId);
  const field = addressOf(fieldAtom(batch, p, op.fieldNodeId));
  return [{ kind: 'field-code', part: p.story.name, field, code }];
}

/**
 * The fields a refresh rewrites, as the engine decides it on the shown paragraph: one it can
 * rewrite whose result is not the text already. The others the engine skips, and so does the
 * change.
 */
export function translateFieldResults(
  op: OpOf<'refreshFieldResults'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const byPart = new Map<string, { field: ReturnType<typeof addressOf>; text: string }[]>();
  for (const update of op.updates) {
    const p = batch.paragraph(update.paragraphId);
    const paragraph = shownParagraph(shownPart(batch, p.story), update.paragraphId);
    const located = locateFieldResults(paragraph).find(
      (field) => field.fieldNodeId === update.fieldNodeId
    );
    if (!located || !located.rewritable || located.cachedText === update.text) continue;
    const field = addressOf(fieldAtom(batch, p, update.fieldNodeId));
    const list = byPart.get(p.story.name) ?? [];
    list.push({ field, text: update.text });
    byPart.set(p.story.name, list);
  }

  return [...byPart].map(([part, fields]) => ({ kind: 'field-result', part, fields }));
}

/**
 * The format of a field result's run, which the editor names by its id: a format of the
 * field's whole item, with what the press changes on that run.
 */
export function translateResultRunFormat(
  op: OpOf<'setRunProperties'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const p = batch.paragraph(op.paragraphId);
  const part = shownPart(batch, p.story);
  const out: OrderedChangeBody[] = [];
  for (const runId of op.targetRunIds ?? []) {
    const atom = p.view.chars
      .slice(op.start, op.end)
      .map((char) => char.item)
      .find(
        (item) => item.kind === 'atom' && (item.pieces ?? []).some((pc) => holds(pc.node, runId))
      );
    const run = findNode(part, runId);
    if (!atom || !run || run.kind !== 'run') throw outOfStep(`no result run ${runId}`);
    const face = childWhere(run, isRunProperties);
    const properties = resultRunFormatProperties(face ?? null, op.properties);
    if (properties.length === 0) continue;
    const start = p.view.chars.findIndex((char) => char.item === atom);
    out.push({
      kind: 'format',
      part: p.story.name,
      from: placeAt(p, start),
      to: placeAt(p, start + 1),
      properties,
    });
  }

  return out;
}

// ---- bookmarks ----

/** The `w:id` of a bookmark marker item. */
function markerIdOf(item: OrderedItem): string | undefined {
  const marker = bookmarkMarker(item);
  return marker ? attributeOf(marker, 'id') : undefined;
}

/** The shown markers of the bookmarks a name means in a story: starts, and their ends by id. */
function namedMarkers(batch: OrderedTranslateBatch, story: OrderedStory, name: string) {
  const key = bookmarkNameKey(name);
  const starts: OrderedItem[] = [];
  const ends: OrderedItem[] = [];
  for (const item of story.tree.items()) {
    const marker = bookmarkMarker(item);
    if (!marker || isHidden(batch.doc, item)) continue;
    if (marker.localName === 'bookmarkEnd') ends.push(item);
    else if (bookmarkNameKey(attributeOf(marker, 'name') ?? '') === key) starts.push(item);
  }

  const ids = new Set(starts.map(markerIdOf));
  return { starts, ends: ends.filter((end) => ids.has(markerIdOf(end))) };
}

/** The markers of a name in one story, past its first bookmark when `keepFirst`. */
function markersToRemove(
  batch: OrderedTranslateBatch,
  story: OrderedStory,
  name: string,
  keepFirst: boolean
): OrderedItem[] {
  const { starts, ends } = namedMarkers(batch, story, name);
  if (!keepFirst) return [...starts, ...ends];
  const [first] = starts;
  const firstId = first ? markerIdOf(first) : undefined;
  const keptEnd = ends.find((end) =>
    bookmarkMarker(end)?.attributes.some((a) => a.localName === 'id' && a.value === firstId)
  );
  return [...starts.slice(1), ...ends.filter((end) => end !== keptEnd)];
}

/** The same name in every other story: the editor takes it out of them in the same write. */
function otherStories(
  batch: OrderedTranslateBatch,
  part: string,
  name: string
): OrderedChangeBody[] {
  const out: OrderedChangeBody[] = [];
  for (const story of batch.doc.stories.values()) {
    if (story.name === part) continue;
    const remove = markersToRemove(batch, story, name, false).map(addressOf);
    if (remove.length > 0) out.push({ kind: 'bookmark', part: story.name, name, remove });
  }

  return out;
}

export function translateInsertBookmark(
  op: OpOf<'insertBookmark'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const from = batch.paragraph(op.start.paragraphId);
  const to = batch.paragraph(op.end.paragraphId);
  const part = from.story.name;
  const remove = markersToRemove(batch, from.story, op.name, false).map(addressOf);
  return [
    ...otherStories(batch, part, op.name),
    {
      kind: 'bookmark',
      part,
      name: op.name,
      from: placeAt(from, op.start.offset),
      to: placeAt(to, op.end.offset),
      remove,
    },
  ];
}

export function translateDeleteBookmark(
  op: OpOf<'deleteBookmark'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  if (op.pair !== undefined) {
    // A deletion's own pair, which it encloses: named by its two markers' node ids
    const story = [...batch.doc.stories.values()].find((s) =>
      [...s.tree.items()].some((item) => item.node?.id === op.pair?.startNodeId)
    );
    if (!story) throw outOfStep(`no bookmark ${op.name}`);
    const ids = new Set([op.pair.startNodeId, op.pair.endNodeId]);
    const remove = [...story.tree.items()]
      .filter((item) => bookmarkMarker(item) !== null && ids.has(item.node?.id ?? ''))
      .filter((item) => !isHidden(batch.doc, item))
      .map(addressOf);
    if (remove.length !== 2) throw outOfStep(`no bookmark pair ${op.name}`);
    return [{ kind: 'bookmark', part: story.name, name: op.name, remove }];
  }

  const out: OrderedChangeBody[] = [];
  const keepFirst = op.keepFirst === true;
  let first = true;
  for (const story of batch.doc.stories.values()) {
    const markers = namedMarkers(batch, story, op.name);
    if (markers.starts.length === 0 && markers.ends.length === 0) continue;

    // The bookmark the name means is the first story's that holds it, which keeps its first pair
    const remove = markersToRemove(batch, story, op.name, keepFirst && first).map(addressOf);
    first = false;
    if (remove.length > 0) out.push({ kind: 'bookmark', part: story.name, name: op.name, remove });
  }

  if (out.length === 0) throw outOfStep(`no bookmark ${op.name}`);
  return out;
}

/**
 * A bookmark's markers taken out of one story, as the editor's package edit takes them out of a
 * part its ops do not address (`Document.deleteBookmark` from a script, for a bookmark in a
 * header, a footer or a note). Nothing when the story holds none.
 */
export function translateBookmarkRemoval(
  batch: OrderedTranslateBatch,
  partName: string,
  name: string
): readonly OrderedChangeBody[] {
  const story = batch.doc.stories.get(partName);
  if (!story) throw outOfStep(`no story ${partName.slice(0, 160)}`);
  const remove = markersToRemove(batch, story, name, false).map(addressOf);
  return remove.length === 0 ? [] : [{ kind: 'bookmark', part: story.name, name, remove }];
}

// ---- a table of contents ----

const markOf = (batch: OrderedTranslateBatch, id: string) => addressOf(batch.paragraph(id).mark);

type TocOp = OpOf<'insertToc'> | OpOf<'replaceTocResult'>;

function rowsOf(batch: OrderedTranslateBatch, op: TocOp): OrderedTocRow[] {
  return op.entries.map((entry) => ({
    level: entry.level,
    text: entry.text,
    heading: markOf(batch, entry.headingParagraphId),
    bookmark: entry.bookmarkName,
    page: entry.pageNumberText,
  }));
}

function bookmarksOf(batch: OrderedTranslateBatch, op: TocOp): OrderedTocBookmark[] {
  return op.bookmarksToCreate.map((bookmark) => ({
    heading: markOf(batch, bookmark.paragraphId),
    name: bookmark.name,
  }));
}

/**
 * A new table of contents. Its rows and page numbers come from the editor's layout: the op
 * carries them, and so does the change, so the other copies write the same text and never lay
 * out to find it.
 */
export function translateInsertToc(
  op: OpOf<'insertToc'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const p = batch.paragraph(op.beforeParagraphId);
  return [
    {
      kind: 'insert-toc',
      part: p.story.name,
      before: addressOf(p.mark),
      alias: op.alias,
      instruction: op.instruction,
      rows: rowsOf(batch, op),
      bookmarks: bookmarksOf(batch, op),
    },
  ];
}

/** The main document's story: a table of contents is in its body. */
function bodyStory(batch: OrderedTranslateBatch): OrderedStory {
  const story = mainStory(batch.doc);
  if (!story) {
    throw outOfStep('no main document');
  }

  return story;
}

/** A detected table of contents in the shown body, by the id the editor names it by. */
function shownToc(batch: OrderedTranslateBatch, tocId: string) {
  const story = bodyStory(batch);
  const part = shownPart(batch, story);
  const toc = findDetectedToc(detectBodyTocs(part), tocId);
  const range = toc ? tocFieldRange(toc) : undefined;
  if (!toc || !range) throw outOfStep(`no table of contents ${tocId}`);
  return { story, part, toc, range };
}

export function translateTocResult(
  op: OpOf<'replaceTocResult'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const { story, toc, range } = shownToc(batch, op.tocId);
  const separate = itemWithNode(
    batch,
    batch.paragraph(range.separateParagraphId),
    range.separateNodeId
  );
  const end = itemWithNode(batch, batch.paragraph(toc.endParagraphId), range.endNodeId);
  return [
    {
      kind: 'toc-result',
      part: story.name,
      separate: addressOf(separate),
      end: addressOf(end),
      instruction: toc.instruction.raw,
      rows: rowsOf(batch, op),
      bookmarks: bookmarksOf(batch, op),
    },
  ];
}

/** The `w:t` elements of a node with their text, by id. */
function textElements(node: OoxmlNode, out = new Map<string, OoxmlElement>(), depth = 0) {
  if (node.kind === 'textValue' || depth >= MAX_WALK_DEPTH) return out;
  if (node.localName === 't') out.set(node.id, node);
  for (const child of node.children) textElements(child, out, depth + 1);
  return out;
}

const textOf = (node: OoxmlElement) =>
  node.children.map((c) => (c.kind === 'textValue' ? c.value : '')).join('');

/** The `w:t` of `before` whose text the write changed in `after`, matched by node id. */
function changedTextElement(before: OoxmlNode, after: OoxmlNode): OoxmlElement | null {
  const old = textElements(before);
  for (const [id, element] of textElements(after)) {
    const target = old.get(id);
    if (target && textOf(target) !== textOf(element)) {
      return target;
    }
  }

  return null;
}

/**
 * Page numbers rewritten in a contents' rows, as the engine rewrites them on the shown body:
 * the `w:t` it changes in each row is a field's (a `PAGEREF` result, whose item is whole) or
 * plain text, which keeps its run and face.
 */
export function translateTocPages(
  op: OpOf<'rewriteTocPageNumbers'>,
  batch: OrderedTranslateBatch
): readonly OrderedChangeBody[] {
  const { story, part } = shownToc(batch, op.tocId);
  const result = applyTreeOp(part, op);
  if (!result.ok)
    throw new OrderedClientRefusal('ordered-refused', 'engine-refused', result.reason);
  const fields: { field: ReturnType<typeof addressOf>; text: string; last: true }[] = [];
  const pages: { from: OrderedChangePlace; to: OrderedChangePlace; text: string }[] = [];
  for (const update of op.updates) {
    const before = findNode(part, update.paragraphId);
    const after = findNode(result.part, update.paragraphId);
    if (!before || !after || before === after) continue;
    const target = changedTextElement(before, after);
    if (!target) {
      continue;
    }

    const p = batch.paragraph(update.paragraphId);
    const atom = p.view.items.find(
      (item) => item.kind === 'atom' && (item.pieces ?? []).some((pc) => holds(pc.node, target.id))
    );
    if (atom) {
      fields.push({ field: addressOf(atom), text: update.pageNumberText, last: true });
      continue;
    }

    const value = target.children.find((c) => c.kind === 'textValue');
    const span =
      value && before.kind === 'paragraph' ? paragraphOffsetIndex(before).spanOf(value) : null;
    if (!span) throw outOfStep(`no page number in ${update.paragraphId}`);
    const end = span.start + textOf(target).length;
    pages.push({ from: placeAt(p, span.start), to: placeAt(p, end), text: update.pageNumberText });
  }

  return [
    ...(fields.length > 0 ? [{ kind: 'field-result' as const, part: story.name, fields }] : []),
    ...(pages.length > 0 ? [{ kind: 'toc-pages' as const, part: story.name, pages }] : []),
  ];
}
