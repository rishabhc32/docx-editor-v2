import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';

/**
 * A link a range write can divide: each child is a text run, or a tracked insertion or deletion
 * of text runs, which divides like the link around it. A suggested link change strikes the
 * link's words and then lets go of only the copy after them.
 *
 * The rule for a ranged link op only: `isTextHyperlink` keeps its narrower meaning for its other
 * callers.
 */
export function isDivisibleHyperlink(node: OoxmlNode): boolean {
  return node.kind === 'hyperlink' && node.children.every(isDivisibleText);
}

function isDivisibleText(node: OoxmlNode): boolean {
  if (node.kind === 'revisionInsert') return node.children.every((run) => isTextRun(run, 't'));
  if (node.kind === 'revisionDelete') {
    return node.children.every((run) => isTextRun(run, 'delText'));
  }
  return isTextRun(node, 't');
}

/** A run of properties and text: `w:t`, or `w:delText` inside a tracked deletion. */
function isTextRun(run: OoxmlNode, text: 't' | 'delText'): boolean {
  return (
    run.kind === 'run' &&
    run.children.every(
      (child) =>
        child.namespaceUri === WML_NAMESPACE_URI &&
        (child.localName === 'rPr' || child.localName === text)
    )
  );
}
