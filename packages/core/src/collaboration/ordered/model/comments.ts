// The document's comments, as the ordered document holds them: the Word file's, and those the
// room's changes added.

import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
} from '../../../store/package/ooxml-tree.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { itemAt } from './document.ts';
import { OrderedRefusal } from './refusal.ts';
import type { OrderedPerson } from './person.ts';
import type { OrderedDocument, OrderedItem, OrderedItemAddress, OrderedStory } from './types.ts';

/**
 * One comment, with its thread and resolved state: a comment of the Word file, whose values are
 * the file's, kept as text and never run, or one a change added.
 *
 * @public
 */
export interface OrderedComment {
  /** The comment's `w:id`: its range markers and its reference in a story name it. */
  readonly id: string;
  /** `w:author`, empty when the file gives none. */
  readonly author: string;
  /** `w:initials`, or null. */
  readonly initials: string | null;
  /** `w:date`, or null. */
  readonly date: string | null;
  /** The `w14:paraId` of each paragraph of the body that has one, upper-case, in order. */
  readonly paraIds: readonly string[];
  /**
   * The last paragraph's `w14:paraId`, upper-case: the key of the comment's thread entry in
   * `commentsExtended` and of its durable id in `commentsIds`. Null when that paragraph has none.
   */
  readonly paraId: string | null;
  /** The `w:id` of the comment this one answers, by the editor's thread rule, or null. */
  readonly parent: string | null;
  /** Whether the thread entry marks the comment resolved (`w15:done`). */
  readonly done: boolean;
  /**
   * The `w:comment` element: as the file wrote it (its body is kept, never rebuilt), or as the
   * change that added the comment built it.
   */
  readonly element: OoxmlElement;
  /** The person the room stamped on the change that added it; null for a file's comment. */
  readonly person: OrderedPerson | null;
}

/**
 * Where a comment's range sits: its story and the addresses of its markers, the first of each
 * kind. A marker the story does not hold is null.
 *
 * @public
 */
export interface OrderedCommentRange {
  /** The story part, such as `/word/document.xml`. */
  readonly story: string;
  /** The marker of the range's start, `w:commentRangeStart`. */
  readonly start: OrderedItemAddress | null;
  /** The marker of the range's end, `w:commentRangeEnd`. */
  readonly end: OrderedItemAddress | null;
  /** The marker of the comment's reference: `w:commentReference`, or the run that holds it. */
  readonly reference: OrderedItemAddress | null;
}

/** The kinds of marker a comment's range has. */
export type OrderedCommentMarker = 'start' | 'end' | 'reference';

const MARKERS: Readonly<Record<string, OrderedCommentMarker>> = {
  commentRangeStart: 'start',
  commentRangeEnd: 'end',
  commentReference: 'reference',
};

/** A node's comment id and marker kind, when it is a comment's marker. */
export function commentMarkerOf(
  node: OoxmlNode | null
): { readonly id: string; readonly marker: OrderedCommentMarker } | null {
  if (node === null || node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) {
    return null;
  }

  const marker = MARKERS[node.localName];
  const id = attributeOf(node, 'id');
  return marker === undefined || id === undefined ? null : { id, marker };
}

/**
 * Note a loaded item in `doc.commentRanges` when it is a comment's marker: the first marker of
 * each kind, in the story of the comment's first marker.
 */
export function noteCommentMarker(doc: OrderedDocument, story: string, item: OrderedItem): void {
  const found = commentMarkerOf(item.node);
  if (found === null) {
    return;
  }

  const range = doc.commentRanges.get(found.id) ?? {
    story,
    start: null,
    end: null,
    reference: null,
  };
  if (range.story === story && range[found.marker] === null) {
    const address = { space: item.space, offset: item.offset };
    doc.commentRanges.set(found.id, { ...range, [found.marker]: address });
  }
}

/**
 * The story of comment `id`'s range and the items of its markers, or null when it has no range
 * in a story. A marker the story does not hold is null.
 */
export function commentRangeItems(
  doc: OrderedDocument,
  id: string
): {
  readonly story: OrderedStory;
  readonly start: OrderedItem | null;
  readonly end: OrderedItem | null;
  readonly reference: OrderedItem | null;
} | null {
  const range = doc.commentRanges.get(id);
  const story = range === undefined ? undefined : doc.stories.get(range.story);
  if (range === undefined || story === undefined) {
    return null;
  }

  const at = (address: OrderedItemAddress | null) =>
    address === null ? null : itemAt(story, address);
  return { story, start: at(range.start), end: at(range.end), reference: at(range.reference) };
}

/** The first comment of a comment's thread: the one Word's flat thread hangs from. */
export function threadRootOf(doc: OrderedDocument, id: string): OrderedComment {
  let comment = doc.comments.get(id);
  if (comment === undefined) {
    throw new OrderedRefusal('unknown-comment', id.slice(0, 64));
  }

  // The model's threads have no cycle; the guard keeps a defect from looping
  const seen = new Set<string>([comment.id]);
  let parent = comment.parent === null ? undefined : doc.comments.get(comment.parent);
  while (parent !== undefined && !seen.has(parent.id)) {
    seen.add(parent.id);
    comment = parent;
    parent = parent.parent === null ? undefined : doc.comments.get(parent.parent);
  }

  return comment;
}
