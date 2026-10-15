// The character style an unlink takes off a link's words, as Word's Remove Hyperlink does.

import { isInlineRunContainer, MAX_INLINE_CONTAINER_DEPTH } from '../package/ooxml-shared.ts';
import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';
import { runPropertiesNodeOf } from './tree-op-nodes.ts';

/**
 * Take a link's character style off a run: the `w:rStyle` naming `styleId` goes, and a `w:rPr`
 * it leaves empty goes too. Everything else stays, another character style included, because
 * only the link's own style belongs to the link. Runs inside revision wrappers are reached the
 * way `withCharacterStyle` reaches them.
 */
export function withoutCharacterStyle(node: OoxmlNode, styleId: string, depth = 0): OoxmlNode {
  if (depth >= MAX_INLINE_CONTAINER_DEPTH) return node;
  if (isInlineRunContainer(node)) {
    return {
      ...node,
      children: node.children.map((child) => withoutCharacterStyle(child, styleId, depth + 1)),
    } as OoxmlNode;
  }
  if (node.kind !== 'run') return node;
  const rPr = runPropertiesNodeOf(node);
  if (!rPr) return node;
  const kept = rPr.children.filter(
    (child) =>
      child.kind === 'textValue' ||
      child.localName !== 'rStyle' ||
      child.namespaceUri !== WML_NAMESPACE_URI ||
      !child.attributes.some(
        (attribute) =>
          attribute.localName === 'val' &&
          attribute.namespaceUri === WML_NAMESPACE_URI &&
          attribute.value === styleId
      )
  );
  if (kept.length === rPr.children.length) return node;
  const children =
    kept.length === 0
      ? node.children.filter((child) => child.id !== rPr.id)
      : node.children.map((child) => (child.id === rPr.id ? { ...rPr, children: kept } : child));
  return { ...node, children } as OoxmlNode;
}
