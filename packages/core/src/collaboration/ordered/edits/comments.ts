// The comments a change adds: a new thread, a reply, and a resolve with an optional last note. A
// comment is no suggestion: nothing decides it, and its markers stay where they were put. Its
// element is built once, as the engine builds a comment, and every write is in the change's undo
// record.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { commentElementOf } from '../../../store/store/comment-element.ts';
import { stampedParagraph } from '../../../store/store/comment-writes.ts';
import type { OrderedIdMinter } from '../change/ids.ts';
import {
  commentMarkerOf,
  commentRangeItems,
  threadRootOf,
  type OrderedComment,
  type OrderedCommentMarker,
} from '../model/comments.ts';
import { createOrderedItem } from '../model/document.ts';
import { insertItems } from '../model/mutations.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { recordUndo } from '../model/undo-record.ts';
import type {
  OrderedAttribution,
  OrderedDocument,
  OrderedItem,
  OrderedItemAddress,
  OrderedStory,
} from '../model/types.ts';
import { withChildren } from '../ooxml-nodes.ts';
import type { Ctx } from './context.ts';
import { freshDecimal, itemBefore, wml } from './marker.ts';
import { mintParagraph } from './para-ids.ts';
import type { OrderedResolvedPlace } from './places.ts';

/** Who writes a comment, and where its new ids come from. */
export interface OrderedCommentWriter {
  readonly doc: OrderedDocument;
  readonly attribution: OrderedAttribution;
  readonly ids: OrderedIdMinter;
}

/** The prefix a built comment's `w14:paraId` takes; compose declares it on the part's root. */
const PARA_ID_PREFIX = 'w14';

/** The markers of a new comment's range: its start, its end, and its reference in a run. */
interface RangeMarkers {
  readonly start: OrderedItem;
  readonly end: OrderedItem;
  readonly reference: OrderedItem;
}

/**
 * A new comment's markers, each under its typed kind as a file's load gives it: the editor
 * finds a comment's range by these kinds. `start` and `end` give the new items the containers
 * around them and whether they sit between blocks.
 */
function rangeMarkers(
  ids: OrderedIdMinter,
  id: string,
  start: Pick<OrderedItem, 'within' | 'block'>,
  end: Pick<OrderedItem, 'within' | 'block'>
): RangeMarkers {
  const marker = (node: OoxmlElement, near: Pick<OrderedItem, 'within' | 'block'>) =>
    createOrderedItem({
      space: ids.marker(),
      kind: 'zero',
      node,
      within: near.within,
      ...(near.block === true ? { block: true } : {}),
    });
  const face = wml(ids, 'runProperties', 'rPr', {}, [
    wml(ids, 'generic', 'rStyle', { val: 'CommentReference' }),
  ]);
  return {
    start: marker(wml(ids, 'commentRangeStart', 'commentRangeStart', { id }), start),
    end: marker(wml(ids, 'commentRangeEnd', 'commentRangeEnd', { id }), end),
    reference: createOrderedItem({
      space: ids.marker(),
      kind: 'zero',
      node: wml(ids, 'commentReference', 'commentReference', { id }),
      within: end.within,
      inRun: true,
      run: ids.node(),
      face,
    }),
  };
}

/** A comment id past every id the file and the room's comments use. */
function newCommentId(writer: OrderedCommentWriter): string {
  const { commentIds, comments } = writer.doc;
  return freshDecimal(writer, (used) => commentIds.has(used) || comments.has(used));
}

/**
 * A new `w14:paraId`, past every id the comments and this session's new paragraphs hold. The
 * session keeps it until the change's undo.
 */
function newParaId(writer: OrderedCommentWriter): string {
  const comments = writer.doc.comments;
  const taken = (paraId: string) => {
    for (const comment of comments.values()) {
      if (comment.paraIds.includes(paraId)) {
        return true;
      }
    }

    return false;
  };
  return mintParagraph(writer, taken).paraId;
}

/** Put a comment in `doc.comments`; the undo puts back what was there. */
function putComment(doc: OrderedDocument, comment: OrderedComment): void {
  const before = doc.comments.get(comment.id);
  doc.comments.set(comment.id, comment);
  recordUndo(doc, () => {
    if (before === undefined) {
      doc.comments.delete(comment.id);
    } else {
      doc.comments.set(comment.id, before);
    }
  });
}

/** Note a new comment's range; the undo takes it back. */
function putRange(
  doc: OrderedDocument,
  id: string,
  story: OrderedStory,
  markers: RangeMarkers
): void {
  const address = (item: OrderedItem): OrderedItemAddress => ({
    space: item.space,
    offset: item.offset,
  });
  doc.commentRanges.set(id, {
    story: story.name,
    start: address(markers.start),
    end: address(markers.end),
    reference: address(markers.reference),
  });
  recordUndo(doc, () => doc.commentRanges.delete(id));
}

/** Add a comment of `text` (a paragraph for each line) by the writer, answering `parent`. */
function addComment(
  writer: OrderedCommentWriter,
  id: string,
  text: string,
  parent: string | null
): void {
  const { author, date, person } = writer.attribution;
  const paragraphs = text.split('\n').map((line) => ({ paraId: newParaId(writer), text: line }));
  const paraIds = paragraphs.map((paragraph) => paragraph.paraId);
  putComment(writer.doc, {
    id,
    author,
    initials: null,
    date: date ?? null,
    paraIds,
    paraId: paraIds[paraIds.length - 1] ?? null,
    parent,
    done: false,
    element: commentElementOf(
      id,
      { author, ...(date === undefined ? {} : { date }) },
      paragraphs,
      PARA_ID_PREFIX
    ),
    person,
  });
}

/**
 * A new comment thread of `text` on the units from `from` up to `to` (a caret at `from` when
 * they are one place): its range start before the first, its range end and its reference before
 * the last place. Returns the comment's id.
 */
export function insertComment(
  ctx: Ctx,
  from: OrderedResolvedPlace,
  to: OrderedResolvedPlace,
  text: string
): string {
  if (to.index < from.index || (to.index === from.index && to.offset < from.offset)) {
    throw new OrderedRefusal('empty-range', 'a comment that ends before it starts');
  }

  const id = newCommentId(ctx);
  const markers = rangeMarkers(ctx.ids, id, from.paragraph.mark, to.paragraph.mark);
  const first = itemBefore(ctx, from);
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(first), [markers.start]);

  // Read after the start went in: at a caret, the end follows the start
  const last = itemBefore(ctx, to);
  insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(last), [markers.end, markers.reference]);
  putRange(ctx.doc, id, ctx.story, markers);
  addComment(ctx, id, text, null);
  return id;
}

/** Every comment of the thread that hangs from `root`, `root` first, in the model's order. */
function threadOf(doc: OrderedDocument, root: OrderedComment): OrderedComment[] {
  const thread = [root];
  for (const comment of doc.comments.values()) {
    if (comment.id !== root.id && threadRootOf(doc, comment.id).id === root.id) {
      thread.push(comment);
    }
  }

  return thread;
}

/** The story and the markers of a comment's range, or a refusal when it has none whole. */
function rangeOf(
  doc: OrderedDocument,
  id: string
): { story: OrderedStory; start: OrderedItem; end: OrderedItem; reference: OrderedItem | null } {
  const range = commentRangeItems(doc, id);
  if (range === null || range.start === null || range.end === null) {
    throw new OrderedRefusal('comment-without-range', id.slice(0, 64));
  }

  return { story: range.story, start: range.start, end: range.end, reference: range.reference };
}

/**
 * The comment with a `w14:paraId` on its last paragraph, which its thread entry is keyed by: a
 * file's comment that has none gets one, as the engine gives one to a comment it threads. The
 * caller puts the keyed comment in the model.
 */
function withParaId(writer: OrderedCommentWriter, comment: OrderedComment): OrderedComment {
  if (comment.paraId !== null) {
    return comment;
  }

  const children = comment.element.children;
  let last = children.length - 1;
  while (last >= 0 && children[last]?.kind !== 'paragraph') {
    last -= 1;
  }

  const paragraph = children[last];
  if (paragraph === undefined || paragraph.kind === 'textValue') {
    throw new OrderedRefusal('not-covered', `comment ${comment.id.slice(0, 64)} has no paragraph`);
  }

  const paraId = newParaId(writer);
  const stamped = stampedParagraph(paragraph, paraId, PARA_ID_PREFIX);
  return {
    ...comment,
    paraIds: [...comment.paraIds, paraId],
    paraId,
    element: withChildren(
      comment.element,
      children.map((child, k) => (k === last ? stamped : child))
    ),
  };
}

/**
 * The item a reply's `marker` goes after, beside its parent's marker `at`, as the engine orders
 * the markers at one place (`coincidentInsertIndex`): a start after the starts that follow the
 * parent's, an end before the ends that come before it, a reference after the references that
 * follow it. So a thread's starts and references run oldest first, and its ends the other way.
 */
function besideParent(
  story: OrderedStory,
  at: OrderedItem,
  marker: OrderedCommentMarker
): OrderedItem | null {
  const { tree } = story;
  const same = (item: OrderedItem | null): item is OrderedItem =>
    item !== null && commentMarkerOf(item.node)?.marker === marker;
  if (marker === 'end') {
    let first = at;
    for (let item = tree.prev(first); same(item); item = tree.prev(first)) {
      first = item;
    }

    return tree.prev(first);
  }

  let last = at;
  for (let item = tree.next(last); same(item); item = tree.next(last)) {
    last = item;
  }

  return last;
}

/**
 * A reply of `text` to the thread of comment `id`: it answers the thread's first comment, and
 * its range sits on that comment's, after the thread's earlier replies, as the engine anchors a
 * reply over its parent's range. Returns its id.
 */
export function replyToComment(writer: OrderedCommentWriter, id: string, text: string): string {
  const { doc } = writer;
  const first = threadRootOf(doc, id);
  const root = withParaId(writer, first);
  if (root !== first) {
    putComment(doc, root);
  }

  const { story, start, end, reference } = rangeOf(doc, root.id);
  const reply = newCommentId(writer);
  const markers = rangeMarkers(writer.ids, reply, start, end);
  insertItems(doc, story, besideParent(story, start, 'start'), [markers.start]);
  insertItems(doc, story, besideParent(story, end, 'end'), [markers.end]);
  insertItems(doc, story, besideParent(story, reference ?? end, 'reference'), [markers.reference]);
  putRange(doc, reply, story, markers);
  addComment(writer, reply, text, root.id);
  return reply;
}

/**
 * Resolve the thread of comment `id`: `note`, when given, is added as the writer's last reply,
 * then every comment of the thread is resolved, as the engine resolves a thread.
 */
export function resolveComment(writer: OrderedCommentWriter, id: string, note?: string): void {
  const { doc } = writer;
  if (note !== undefined) {
    replyToComment(writer, id, note);
  }

  for (const comment of threadOf(doc, threadRootOf(doc, id))) {
    putComment(doc, { ...withParaId(writer, comment), done: true });
  }
}
