// A composed package with no comment at all, for a file that keeps only the document and its
// tracked changes.

import { withPart, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { partNameKey } from '../../../store/package/opc-names.ts';
import { withoutPart } from '../../../store/package/package-edit.ts';
import {
  resolveRelationship,
  type RelationshipRecord,
} from '../../../store/package/relationships.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { isRunProperties, isWml, withChildren } from '../ooxml-nodes.ts';
import { OFFICE_RELATIONSHIPS } from './relationship-types.ts';

/** The comments part and the parts Word keeps a comment's thread and ids in. */
const COMMENT_RELATIONSHIPS: ReadonlySet<string> = new Set([
  `${OFFICE_RELATIONSHIPS}/comments`,
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  'http://schemas.microsoft.com/office/2016/relationships/commentsIds',
  'http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible',
]);

const MARKERS = ['commentRangeStart', 'commentRangeEnd', 'commentReference'];

const isMarker = (node: OoxmlNode): boolean => MARKERS.some((name) => isWml(node, name));

/** The parts that the relationships `related` accepts name, by part name key. */
function relatedParts(
  pkg: OoxmlPackage,
  related: (record: RelationshipRecord) => boolean
): Map<string, string> {
  const parts = new Map<string, string>();
  for (const records of pkg.relationships.values()) {
    for (const record of records) {
      const resolved = related(record) ? resolveRelationship(record) : null;
      if (resolved?.mode === 'Internal' && resolved.target.ok) {
        const part = resolved.target.partName;
        parts.set(partNameKey(part), part);
      }
    }
  }

  return parts;
}

/** The package without `part`, or a refusal: a half stripped file is never written. */
function withoutNamedPart(pkg: OoxmlPackage, part: string): OoxmlPackage {
  const removed = withoutPart(pkg, part);
  if (!removed.ok) {
    throw new OrderedRefusal('engine-refused', `the comments cannot leave ${part}`);
  }

  return removed.pkg;
}

/**
 * The node without comment markers, or null when it goes: a marker, or a run a marker leaves
 * holding only its properties, as the engine's comment delete drops one.
 */
function withoutMarkers(node: OoxmlNode): OoxmlNode | null {
  if (node.kind === 'textValue') {
    return node;
  }

  if (isMarker(node)) {
    return null;
  }

  let changed = false;
  const children: OoxmlNode[] = [];
  for (const child of node.children) {
    const kept = withoutMarkers(child);
    changed ||= kept !== child;
    if (kept !== null) {
      children.push(kept);
    }
  }

  if (!changed) {
    return node;
  }

  if (isWml(node, 'r') && children.every(isRunProperties)) {
    return null;
  }

  return withChildren(node, children);
}

/** Every XML part without comment markers. */
function withoutMarkersInParts(pkg: OoxmlPackage): OoxmlPackage {
  let next = pkg;
  for (const part of pkg.parts.values()) {
    const root = withoutMarkers(part.root);
    if (root !== null && root !== part.root && root.kind !== 'textValue') {
      next = withPart(next, { ...part, root });
    }
  }

  return next;
}

/**
 * `pkg` with no comment: no comment marker in any part (a story, a text box in its host part,
 * any other XML part), no comments, commentsExtended, commentsIds or commentsExtensible part of
 * any owner, with their relationships and content types, and no part that only removed parts
 * name, in turn (a picture in a comment, a chart's workbook). Everything else stays as it is:
 * tracked changes (an emptied one too), and the people part, which also names the authors of
 * tracked changes.
 *
 * @throws OrderedRefusal `engine-refused` when a part cannot leave (an owner whose
 *   relationships the package keeps as bytes).
 */
export function withoutComments(pkg: OoxmlPackage): OoxmlPackage {
  const removed = new Set<string>();
  let leaving = relatedParts(pkg, (record) => COMMENT_RELATIONSHIPS.has(record.type));
  let next = pkg;
  while (leaving.size > 0) {
    const owners = leaving;
    const theirs = relatedParts(next, (record) => owners.has(partNameKey(record.ownerPart)));
    for (const [key, part] of owners) {
      next = withoutNamedPart(next, part);
      removed.add(key);
    }

    // A part the removed parts named goes next when nothing else names it
    const named = relatedParts(next, () => true);
    leaving = new Map([...theirs].filter(([key]) => !removed.has(key) && !named.has(key)));
  }

  return withoutMarkersInParts(next);
}
