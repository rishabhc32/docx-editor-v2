import { WML_NAMESPACE_URI, type OoxmlElement, type OoxmlNode } from '../package/ooxml-tree.ts';
import { isMarkStampContainer } from './revision-marker-content.ts';
import { PARAGRAPH_VOCABULARY, propertyElement, schemaInsertIndex } from './tree-op-properties.ts';

/**
 * Clone a split's `w:pPr` with fresh ids, keeping the `mark` id family on the containers that a
 * paragraph-mark stamp created.
 *
 * A split leaves the ORIGINAL mark on the tail, so resolving a revision on that mark has to
 * remove those containers from the tail just as it would from the unsplit paragraph. With
 * ordinary ids, rejecting a proposed join after an Enter left an empty `<w:pPr><w:rPr/>`
 * that the paragraph never had.
 */
export function cloneSplitProperties(
  node: OoxmlNode,
  nextId: () => string,
  markId: () => string
): OoxmlNode {
  const id = splitIdFor(node, nextId, markId);
  if (node.kind === 'textValue') return { id, kind: 'textValue', value: node.value };
  return {
    ...node,
    id,
    children: node.children.map((child) => cloneSplitProperties(child, nextId, markId)),
  } as OoxmlNode;
}
/** A split copy's id: from the `mark` family for a container a paragraph-mark stamp created. */
function splitIdFor(node: OoxmlNode, nextId: () => string, markId: () => string): string {
  return isMarkStampContainer(node) ? markId() : nextId();
}

/**
 * The tail's `w:pPr`: a clone of the head's, with `w:pStyle` restated when the caller named
 * one.
 *
 * The clone is the default because every other paragraph property survives an Enter in Word
 * — centring, spacing, borders. A different following style also drops direct numbering,
 * allowing the new style to decide whether the following paragraph is a list item.
 * A tail left with nothing in its `w:pPr` drops the container rather than serializing an
 * empty one, so a paragraph whose only property was its style digests like one that never
 * had any.
 */
export function withTailStyle(
  pPr: OoxmlElement | undefined,
  tailStyleId: string | null | undefined,
  nextId: () => string,
  markId: () => string
): OoxmlNode | undefined {
  if (tailStyleId === undefined) return pPr ? cloneSplitProperties(pPr, nextId, markId) : undefined;
  const currentStyle = pPr?.children.find(
    (child) => child.kind !== 'textValue' && child.localName === 'pStyle'
  );
  const currentStyleId =
    currentStyle && currentStyle.kind !== 'textValue'
      ? (currentStyle.attributes.find((attribute) => attribute.localName === 'val')?.value ?? null)
      : null;
  const kept = (pPr?.children ?? []).filter(
    (child) =>
      child.kind === 'textValue' ||
      (child.localName !== 'pStyle' &&
        (child.localName !== 'numPr' || tailStyleId === currentStyleId))
  );
  const children = kept.map((child) => cloneSplitProperties(child, nextId, markId));
  if (tailStyleId !== null) {
    children.splice(
      schemaInsertIndex(children, PARAGRAPH_VOCABULARY.sequence, 'pStyle'),
      0,
      propertyElement({ localName: 'pStyle', attributes: { val: tailStyleId } }, nextId())
    );
  }
  if (children.length === 0) return undefined;
  if (pPr) {
    return { ...pPr, id: splitIdFor(pPr, nextId, markId), children } as unknown as OoxmlNode;
  }
  return {
    id: nextId(),
    kind: 'paragraphProperties',
    namespaceUri: WML_NAMESPACE_URI,
    localName: 'pPr',
    prefix: 'w',
    namespaceBindings: [],
    attributes: [],
    children,
  } as unknown as OoxmlNode;
}
