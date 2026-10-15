import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';

/**
 * A link a range write can divide: each child is a text run, or a tracked insertion or deletion
 * of text runs, which divides like the link around it. A suggested link change strikes the
 * link's words and then lets go of only the copy after them. Another author's proposed link
 * holds its words in their insertion, so that strike lands inside the insertion.
 *
 * The rule for a ranged link op only: `isTextHyperlink` keeps its narrower meaning for its other
 * callers.
 */
export function isDivisibleHyperlink(node: OoxmlNode): boolean {
  return node.kind === 'hyperlink' && node.children.every(isDivisibleText);
}

function isDivisibleText(node: OoxmlNode): boolean {
  if (node.kind === 'revisionInsert') {
    return node.children.every((child) => isTextRun(child, 't') || isStruckText(child));
  }
  return isStruckText(node) || isTextRun(node, 't');
}

/** A tracked deletion of text runs. */
function isStruckText(node: OoxmlNode): boolean {
  return node.kind === 'revisionDelete' && node.children.every((run) => isTextRun(run, 'delText'));
}

/** A run of properties and text: `w:t`, or `w:delText` inside a tracked deletion. */
export function isTextRun(run: OoxmlNode, text: 't' | 'delText'): boolean {
  return (
    run.kind === 'run' &&
    run.children.every(
      (child) =>
        child.namespaceUri === WML_NAMESPACE_URI &&
        (child.localName === 'rPr' || child.localName === text)
    )
  );
}
