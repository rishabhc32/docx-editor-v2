import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import { isRangeMarkerKind } from '../package/ooxml-shared.ts';

/** Empty revision markers cannot contain preserved content. */
export function hasRevisionMarkerContent(node: OoxmlElement): boolean {
  return node.children.some((child) => child.kind !== 'textValue' || child.value.trim().length > 0);
}

/**
 * A child that marks a position rather than holding content: bookmark, comment-range and
 * move-range boundaries, permission-range boundaries, and proofing marks.
 */
export function isInertMarker(node: OoxmlNode): boolean {
  if (node.kind === 'textValue') return false;
  if (node.kind === 'bookmarkStart' || node.kind === 'bookmarkEnd') return true;
  if (isRangeMarkerKind(node.kind)) return true;
  return (
    node.kind === 'generic' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    (node.localName === 'proofErr' ||
      node.localName === 'permStart' ||
      node.localName === 'permEnd')
  );
}

/**
 * A `w:pPr` or `w:rPr` that held nothing but the revision marks this resolution drops.
 *
 * Proposing a paragraph-mark change writes those containers when the paragraph had none, so
 * resolving the proposal removes them again, and an accepted new paragraph writes no empty
 * `<w:pPr><w:rPr/></w:pPr>`. The rule reads the content, not who made the container, so it
 * holds the same in the session that made the mark, in a document reopened from saved bytes
 * and on every collaboration replica. A source `w:rPr` that held nothing before a stamp goes
 * with the stamp too: saved bytes cannot tell it from one the stamp made.
 */
export function heldOnlyResolvedMarks(node: OoxmlNode, dropped: ReadonlySet<string>): boolean {
  if (node.kind === 'textValue' || node.children.length === 0) return false;
  const markProperties =
    node.kind === 'runProperties' ||
    (node.kind === 'generic' &&
      node.namespaceUri === WML_NAMESPACE_URI &&
      node.localName === 'rPr');
  if (node.kind !== 'paragraphProperties' && !markProperties) return false;
  return node.children.every((child) =>
    markProperties ? dropped.has(child.id) : heldOnlyResolvedMarks(child, dropped)
  );
}

/**
 * A `w:trPr` that held nothing but the properties this resolution drops: the row's own
 * `w:ins` or `w:del`, an accepted `w:trPrChange`, or a height Word added without a record.
 *
 * The row marker is the row decision itself, so its container goes with it, and the row
 * digests like one that never had row properties. A `w:trPr` that arrives empty holds nothing
 * this resolution drops, so it stays.
 */
export function heldOnlyResolvedRowProperties(
  node: OoxmlNode,
  dropped: ReadonlySet<string>
): boolean {
  return (
    node.kind === 'generic' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    node.localName === 'trPr' &&
    node.children.length > 0 &&
    node.children.every((child) => dropped.has(child.id))
  );
}
