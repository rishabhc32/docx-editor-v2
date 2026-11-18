import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';

/**
 * An ActiveX control (`w:control`, ECMA-376 Part 1, §17.3.3.3): a program, which no copy
 * carries. An embedded object (`w:object` with `o:OLEObject`) travels with its parts as bytes
 * that nothing opens (`clipboard-fragment-objects.ts`), as Word copies it.
 */
export function isClipboardObject(node: OoxmlNode): boolean {
  if (node.kind === 'textValue' || node.namespaceUri !== WML_NAMESPACE_URI) return false;
  if (node.localName === 'control') return true;
  return (
    (node.localName === 'object' || node.localName === 'pict') &&
    node.children.some(
      (child) =>
        child.kind !== 'textValue' &&
        child.namespaceUri === WML_NAMESPACE_URI &&
        child.localName === 'control'
    )
  );
}

/** An ActiveX control cannot travel in clipboard fragments. Refuse the whole transfer. */
export function containsClipboardObject(nodes: readonly OoxmlNode[]): boolean {
  const pending = [...nodes];
  while (pending.length) {
    const node = pending.pop()!;
    if (node.kind === 'textValue') continue;
    if (isClipboardObject(node)) return true;
    for (const child of node.children) pending.push(child);
  }
  return false;
}

/** Ink, and alternate content around a drawing, beside the drawings themselves. */
const DROPPED_WITH_PART: ReadonlySet<string> = new Set([
  'http://schemas.microsoft.com/office/word/2010/wordml contentPart',
  'http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent',
]);

/**
 * Content a copy drops whole when it names a part that does not travel: a drawing, a VML
 * picture, an embedded object, ink (`w14:contentPart`), or the alternate content around one, so
 * nothing pasted names a relationship the target does not have.
 */
export function dropsWithItsPart(node: OoxmlNode): boolean {
  if (node.kind === 'textValue') return false;
  if (node.kind === 'drawing') return true;
  if (node.namespaceUri === WML_NAMESPACE_URI) {
    return node.localName === 'pict' || node.localName === 'object';
  }
  return DROPPED_WITH_PART.has(`${node.namespaceUri} ${node.localName}`);
}
