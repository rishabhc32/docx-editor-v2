// The main document's comments as the model's records. The editor's own readers give each
// comment's thread and resolved state, so a room lists the threads the editor alone lists.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { paraIdOf } from '../../../store/package/para-id.ts';
import {
  commentAnchorsOfStoryTransient,
  commentsOfPart,
  threadStateOfPart,
  type CommentAnchor,
  type CommentThreadState,
} from '../../../store/store/comment-reads.ts';
import {
  commentPartNameOf,
  commentsExtendedPartNameOf,
} from '../../../store/store/comment-writes.ts';
import { commentItemsOf } from '../../../store/store/review-reads.ts';
import type { ReviewCommentItem } from '../../../store/store/review-items.ts';
import type { OrderedDocument } from '../model/types.ts';
import { attributeOf, isWml } from '../ooxml-nodes.ts';
import { chargeItem, type LoadBudget } from './state.ts';

/**
 * Record the comments of the main document's comments part in `doc.comments`, each `w:comment`
 * of the part's root by its `w:id` (the first one when the file repeats an id), with no person. Each comment's nodes count against the load's item limit. A thread
 * entry the editor ignores is ignored here.
 */
export function loadComments(
  doc: OrderedDocument,
  pkg: OoxmlPackage,
  stories: Iterable<OoxmlPart>,
  budget: LoadBudget
): void {
  const main = pkg.mainDocumentPart;
  const part = pkg.parts.get(commentPartNameOf(pkg, main));
  if (part === undefined) {
    return;
  }

  const elements = commentElementsOf(part);
  const items = threadItemsOf(pkg, part, stories);
  for (const [id, element] of elements) {
    chargeNodes(budget, element);
    const item = items.get(id);
    const { paraIds, paraId } = paraIdsOf(element);
    const parent = item?.parentId;
    doc.comments.set(id, {
      id,
      author: attributeOf(element, 'author') ?? '',
      initials: attributeOf(element, 'initials') ?? null,
      date: attributeOf(element, 'date') ?? null,
      paraIds,
      paraId,
      // A parent the model does not hold (a comment nested away from the part's root) is none
      parent: parent !== undefined && elements.has(parent) ? parent : null,
      done: item?.resolved ?? false,
      element,
      person: null,
    });
  }
}

/** The `w:comment` elements of the part's root by `w:id`, the first of each id. */
function commentElementsOf(part: OoxmlPart): Map<string, OoxmlElement> {
  const elements = new Map<string, OoxmlElement>();
  for (const child of part.root.children) {
    if (child.kind === 'textValue' || !isWml(child, 'comment')) {
      continue;
    }

    const id = attributeOf(child, 'id');
    if (id !== undefined && !elements.has(id)) {
      elements.set(id, child);
    }
  }

  return elements;
}

/** The editor's comment cards by id (the first of each id): threads and resolved state. */
function threadItemsOf(
  pkg: OoxmlPackage,
  part: OoxmlPart,
  stories: Iterable<OoxmlPart>
): Map<string, ReviewCommentItem> {
  const anchors: CommentAnchor[] = [];
  for (const story of stories) {
    for (const anchor of commentAnchorsOfStoryTransient(story)) {
      anchors.push(anchor);
    }
  }

  const extended = pkg.parts.get(commentsExtendedPartNameOf(pkg, pkg.mainDocumentPart));
  const states =
    extended === undefined ? new Map<string, CommentThreadState>() : threadStateOfPart(extended);
  const items = new Map<string, ReviewCommentItem>();
  for (const item of commentItemsOf(commentsOfPart(part), anchors, states)) {
    if (!items.has(item.id)) {
      items.set(item.id, item);
    }
  }

  return items;
}

/**
 * The upper-case `w14:paraId` of each body paragraph that has one, in order, and the last
 * paragraph's (null when it has none).
 */
function paraIdsOf(element: OoxmlElement): {
  readonly paraIds: string[];
  readonly paraId: string | null;
} {
  const paraIds: string[] = [];
  let paraId: string | null = null;
  for (const child of element.children) {
    if (child.kind !== 'paragraph') {
      continue;
    }

    paraId = paraIdOf(child)?.toUpperCase() ?? null;
    if (paraId !== null) {
      paraIds.push(paraId);
    }
  }

  return { paraIds, paraId };
}

/** Count each node of `element` against the load's item limit. */
function chargeNodes(budget: LoadBudget, element: OoxmlElement): void {
  const stack: OoxmlNode[] = [element];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    chargeItem(budget);
    if (node.kind !== 'textValue') {
      for (const child of node.children) {
        stack.push(child);
      }
    }
  }
}
