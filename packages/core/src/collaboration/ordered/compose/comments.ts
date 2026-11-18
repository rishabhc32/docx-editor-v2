// The main document's comment parts as the model gives them: `comments.xml` with each comment as
// the model holds it, and the thread entries of `commentsExtended.xml`. Parts the model leaves as
// the file had them stay the file's own objects. `commentsIds` and `commentsExtensible` keep the
// file's entries; the engine writes none for a new comment.

import {
  attribute,
  COMMENTS_TYPE,
  W15_NAMESPACE_URI,
} from '../../../store/package/comment-lifecycle-scan.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { withPart } from '../../../store/package/ooxml-package.ts';
import {
  readOoxmlPart,
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { withNewPart } from '../../../store/package/package-edit.ts';
import { threadStateOfPart, type CommentThreadState } from '../../../store/store/comment-reads.ts';
import {
  COMMENTS_EXTENDED_REL,
  COMMENTS_EXTENDED_TYPE,
  COMMENTS_REL,
  commentPartNameOf,
  commentsExtendedPartNameOf,
  resolvedEntry,
  withMainDocumentRelationship,
  withResolvedDone,
} from '../../../store/store/comment-writes.ts';
import type { OrderedComment } from '../model/comments.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { attributeOf, isWml, withChildren, withDeclaredPrefixes } from '../ooxml-nodes.ts';

/** The comment parts' roots as the model gives them; null where the file's part stays. */
export interface OrderedCommentRoots {
  readonly comments: OoxmlElement | null;
  readonly extended: OoxmlElement | null;
}

const KEEP_FILE: OrderedCommentRoots = Object.freeze({ comments: null, extended: null });

/** A part's root parsed from a literal, with ids under `name`. */
function emptyRoot(name: string, xml: string, contentType: string): OoxmlElement {
  const read = readOoxmlPart(xml, { name, contentType });
  if (!read.ok) {
    throw new OrderedRefusal('engine-refused', `a comment part template: ${read.reason}`);
  }

  return read.part.root;
}

/** A `w15:commentEx` entry's upper-case `w15:paraId`, or null for any other node. */
function entryParaId(node: OoxmlNode): string | null {
  if (
    node.kind === 'textValue' ||
    node.namespaceUri !== W15_NAMESPACE_URI ||
    node.localName !== 'commentEx'
  ) {
    return null;
  }

  return attribute(node, W15_NAMESPACE_URI, 'paraId')?.toUpperCase() ?? null;
}

/**
 * `comments.xml`'s root: the file's children with each comment as the model holds it, then the
 * new comments.
 */
function commentsRoot(
  doc: OrderedDocument,
  file: OoxmlElement | null,
  name: string
): { readonly root: OoxmlElement | null; readonly inFile: ReadonlySet<string> } {
  const inFile = new Set<string>();
  const children: OoxmlNode[] = [];
  let changed = false;
  for (const child of file?.children ?? []) {
    const id =
      child.kind === 'textValue' || !isWml(child, 'comment') ? undefined : attributeOf(child, 'id');
    const comment = id === undefined || inFile.has(id) ? undefined : doc.comments.get(id);
    if (id !== undefined) {
      inFile.add(id);
    }

    children.push(comment?.element ?? child);
    changed ||= comment !== undefined && comment.element !== child;
  }

  const added = [...doc.comments.values()]
    .filter((comment) => !inFile.has(comment.id))
    .map((comment) => comment.element);
  if (!changed && added.length === 0) {
    return { root: null, inFile };
  }

  const base =
    file ?? emptyRoot(name, `<w:comments xmlns:w="${WML_NAMESPACE_URI}"/>`, COMMENTS_TYPE);
  const replaced = children.filter((child, k) => child !== file?.children[k]);
  const root = withChildren(base, [...children, ...added]);
  return { root: withDeclaredPrefixes(root, [...replaced, ...added]), inFile };
}

/**
 * The comments that need a thread entry the file does not give: a comment the file does not
 * hold, a resolved one, and the comment each of those answers, so both ends of every thread are
 * written (Word drops a thread whose parent has no entry). So a save with no comment edit still
 * adds an entry for the parent of a file's resolved reply when the file gives that parent none.
 */
function entriesNeeded(doc: OrderedDocument, inFile: ReadonlySet<string>): Set<OrderedComment> {
  const needed = new Set<OrderedComment>();
  const need = (comment: OrderedComment) => {
    let at: OrderedComment | undefined = comment;
    while (at !== undefined && !needed.has(at)) {
      needed.add(at);
      at = at.parent === null ? undefined : doc.comments.get(at.parent);
    }
  };
  for (const comment of doc.comments.values()) {
    if (!inFile.has(comment.id) || comment.done) {
      need(comment);
    }
  }

  return needed;
}

/** The `w14:paraId` of the comment that `comment` answers; undefined when there is none. */
function parentParaId(doc: OrderedDocument, comment: OrderedComment): string | undefined {
  if (comment.parent === null) {
    return undefined;
  }

  return doc.comments.get(comment.parent)?.paraId ?? undefined;
}

/**
 * `commentsExtended.xml`'s root: the file's entries, each resolved as the model says, then the
 * new entries.
 */
function extendedRoot(
  doc: OrderedDocument,
  part: OoxmlPart | undefined,
  name: string,
  inFile: ReadonlySet<string>
): OoxmlElement | null {
  const file = part?.root ?? null;
  const byParaId = new Map<string, OrderedComment>();
  for (const comment of doc.comments.values()) {
    if (comment.paraId !== null && !byParaId.has(comment.paraId)) {
      byParaId.set(comment.paraId, comment);
    }
  }

  // The resolved state the file gives each comment, as the editor reads it
  const states =
    part === undefined ? new Map<string, CommentThreadState>() : threadStateOfPart(part);
  const entered = new Set<OrderedComment>();
  const children: OoxmlNode[] = [];
  let changed = false;
  for (const child of file?.children ?? []) {
    const paraId = entryParaId(child);
    const comment = paraId === null ? undefined : byParaId.get(paraId);
    if (child.kind === 'textValue' || paraId === null || comment === undefined) {
      children.push(child);
      continue;
    }

    entered.add(comment);
    if (comment.done === (states.get(paraId)?.done ?? false)) {
      children.push(child);
      continue;
    }

    children.push(withResolvedDone(child, parentParaId(doc, comment), comment.done));
    changed = true;
  }

  // In the model's order, so a thread's first comment comes before its replies
  const needed = entriesNeeded(doc, inFile);
  const added: OoxmlElement[] = [];
  for (const comment of doc.comments.values()) {
    if (needed.has(comment) && !entered.has(comment) && comment.paraId !== null) {
      const id = `${name}#ex-${comment.paraId}`;
      added.push(resolvedEntry(id, comment.paraId, parentParaId(doc, comment), comment.done));
    }
  }

  if (!changed && added.length === 0) {
    return null;
  }

  const base =
    file ??
    emptyRoot(name, `<w15:commentsEx xmlns:w15="${W15_NAMESPACE_URI}"/>`, COMMENTS_EXTENDED_TYPE);
  return withDeclaredPrefixes(withChildren(base, [...children, ...added]), added);
}

/**
 * The comment parts' roots as the model gives them, read against the file's parts in `pkg`:
 * null for a part that stays as the file has it (or does not have it).
 */
export function orderedCommentRoots(doc: OrderedDocument, pkg: OoxmlPackage): OrderedCommentRoots {
  if (doc.comments.size === 0) {
    return KEEP_FILE;
  }

  const main = pkg.mainDocumentPart;
  const commentsName = commentPartNameOf(pkg, main);
  const extendedName = commentsExtendedPartNameOf(pkg, main);
  const { root, inFile } = commentsRoot(
    doc,
    pkg.parts.get(commentsName)?.root ?? null,
    commentsName
  );
  const extended = extendedRoot(doc, pkg.parts.get(extendedName), extendedName, inFile);
  return { comments: root, extended };
}

/** A part with this root: the package's own, or a new part related from the main document. */
function withRoot(
  pkg: OoxmlPackage,
  name: string,
  root: OoxmlElement,
  contentType: string,
  relationship: string
): OoxmlPackage {
  const part = pkg.parts.get(name);
  if (part !== undefined) {
    return part.root === root ? pkg : withPart(pkg, { ...part, root });
  }

  return withMainDocumentRelationship(
    withNewPart(pkg, name, root, contentType),
    relationship,
    name
  );
}

/** The package with the comment parts' roots in place: made, with their relationships, when new. */
export function withOrderedCommentRoots(
  pkg: OoxmlPackage,
  roots: OrderedCommentRoots
): OoxmlPackage {
  const main = pkg.mainDocumentPart;
  let next = pkg;
  if (roots.comments !== null) {
    const name = commentPartNameOf(next, main);
    next = withRoot(next, name, roots.comments, COMMENTS_TYPE, COMMENTS_REL);
  }

  if (roots.extended !== null) {
    const name = commentsExtendedPartNameOf(next, main);
    next = withRoot(next, name, roots.extended, COMMENTS_EXTENDED_TYPE, COMMENTS_EXTENDED_REL);
  }

  return next;
}

/** The package with the comment parts as the model gives them. */
export function withOrderedComments(pkg: OoxmlPackage, doc: OrderedDocument): OoxmlPackage {
  return withOrderedCommentRoots(pkg, orderedCommentRoots(doc, pkg));
}
