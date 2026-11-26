import { normalizedFingerprintAttributes } from '../package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';

/**
 * OOXML equality that ignores canonical node ids, which are intentionally fresh after copies, and
 * what a save changes without changing the meaning: the order of attributes, their prefixes, and a
 * text's `xml:space` where the text does not need it (as the oracle counts it).
 */
export function equivalentNodes(left: readonly OoxmlNode[], right: readonly OoxmlNode[]): boolean {
  if (left.length !== right.length) return false;
  return left.every((node, index) => equivalentNode(node, right[index]!));
}

function equivalentNode(left: OoxmlNode, right: OoxmlNode): boolean {
  if (left.kind !== right.kind) return false;
  if (left.kind === 'textValue' || right.kind === 'textValue') {
    return left.kind === 'textValue' && right.kind === 'textValue' && left.value === right.value;
  }
  if (left.namespaceUri !== right.namespaceUri || left.localName !== right.localName) {
    return false;
  }
  return sameAttributes(left, right) && equivalentNodes(left.children, right.children);
}

/** The same attributes by expanded name and value, in any order. */
function sameAttributes(left: OoxmlElement, right: OoxmlElement): boolean {
  const ours = normalizedFingerprintAttributes(left);
  const theirs = normalizedFingerprintAttributes(right);
  return (
    ours.length === theirs.length &&
    ours.every((a) =>
      theirs.some(
        (b) =>
          b.namespaceUri === a.namespaceUri && b.localName === a.localName && b.value === a.value
      )
    )
  );
}
