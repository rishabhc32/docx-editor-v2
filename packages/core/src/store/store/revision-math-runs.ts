// Word's tracked changes inside an equation: a `w:ins` or a `w:del` around a math run's content.

import { parentNodeOf } from '../package/ooxml-edit.ts';
import { OFFICE_MATH_NAMESPACE_URI } from '../package/omml-equation.ts';
import {
  WML_NAMESPACE_URI,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import type { Resolution } from './revision-rebuild-helpers.ts';

const isMath = (node: OoxmlNode | null | undefined, localName: string): boolean =>
  node !== null &&
  node !== undefined &&
  node.kind !== 'textValue' &&
  node.namespaceUri === OFFICE_MATH_NAMESPACE_URI &&
  node.localName === localName;

const isMark = (node: OoxmlNode): boolean =>
  node.kind !== 'textValue' &&
  node.namespaceUri === WML_NAMESPACE_URI &&
  (node.localName === 'ins' || node.localName === 'del');

/**
 * Whether a node is a `w:ins` or a `w:del` around a math run's content ([MS-OI29500]
 * 22.1.2.87(e); Word 2016 files `RP013`, `RP014`). Word 16 writes a whole equation deleted with
 * tracking on as one `w:del` inside each `m:r`, the `m:oMath` kept (Word 16 for Mac, 2026-10-29).
 */
export const isMathRunRevision = (node: OoxmlNode, parent: OoxmlNode | null): boolean =>
  isMark(node) && isMath(parent, 'r');

/** A math run's own properties, which stay when its content goes. */
const isRunProperties = (node: OoxmlNode): boolean =>
  isMath(node, 'rPr') ||
  (node.kind !== 'textValue' &&
    node.namespaceUri === WML_NAMESPACE_URI &&
    node.localName === 'rPr');

/** Whether a node is a `w:ins` or a `w:del` holding a math run's text (`m:t`), read alone. */
export const holdsMathRunContent = (node: OoxmlNode): boolean =>
  isMark(node) && node.kind !== 'textValue' && node.children.some((child) => isMath(child, 't'));

/**
 * How a mark around a math run's content resolves, as a content wrapper does; null for any other
 * node. The engine reads such a mark as a generic element, so its name says the direction; a
 * generic `w:ins` or `w:del` with no math text in it (a numbering record's) is not one.
 * Accepting a deletion or rejecting an insertion removes the run's content, and the run, the
 * equation and the math paragraph that decision emptied go with it (`emptiedByDecision`).
 */
export function mathRunResolution(
  node: OoxmlElement,
  action: 'accept' | 'reject'
): Resolution | null {
  if (!holdsMathRunContent(node)) return null;
  if (node.localName === 'ins') return action === 'accept' ? 'unwrap' : 'remove';
  return action === 'accept' ? 'remove' : 'restore';
}

/** Whether a math subtree holds a math run. */
function holdsMathRun(node: OoxmlNode): boolean {
  const stack: OoxmlNode[] = [node];
  for (let next = stack.pop(); next !== undefined; next = stack.pop()) {
    if (next.kind === 'textValue') continue;
    if (isMath(next, 'r')) return true;
    for (const child of next.children) {
      stack.push(child);
    }
  }

  return false;
}

/**
 * Whether a decision emptied a container (`hollow` is the container rebuilt, `before` as it
 * was), so it goes: a simple field or a link with no runs, a math run with only its properties,
 * an equation with no math run, a math paragraph with no equation.
 */
export function emptiedByDecision(before: OoxmlNode, hollow: OoxmlElement | null): boolean {
  if (before.kind === 'textValue' || hollow === null) return false;
  if (before.kind === 'fldSimple' || before.kind === 'hyperlink') {
    return hollow.children.length === 0;
  }

  if (isMath(before, 'r')) {
    return (
      before.children.some((child) => !isRunProperties(child)) &&
      hollow.children.every(isRunProperties)
    );
  }

  if (isMath(before, 'oMath')) return holdsMathRun(before) && !holdsMathRun(hollow);
  if (isMath(before, 'oMathPara')) {
    return (
      before.children.some((child) => isMath(child, 'oMath')) &&
      !hollow.children.some((child) => isMath(child, 'oMath'))
    );
  }

  return false;
}

/**
 * The review group of a mark around a math run's content: its equation (the outermost
 * `m:oMath` or `m:oMathPara`), so one equation's marks of one direction, author and date are one
 * card, as one deletion of the whole equation reads. Null for any other site.
 */
export function mathRunGroup(
  part: OoxmlPart,
  node: OoxmlNode,
  parent: OoxmlNode | null
): string | null {
  if (!isMathRunRevision(node, parent) || parent === null || node.kind === 'textValue') return null;
  let equation: OoxmlElement | null = null;
  for (let at = parentNodeOf(part, parent.id); at !== null; at = parentNodeOf(part, at.id)) {
    if (isMath(at, 'oMath') || isMath(at, 'oMathPara')) equation = at;
  }

  return equation === null ? null : `math\u0000${equation.id}\u0000${node.localName}`;
}

/** The review kind of a mark around a math run's content; null for any other site. */
export function mathRunKind(node: OoxmlNode, parent: OoxmlNode | null): 'insert' | 'delete' | null {
  if (!isMathRunRevision(node, parent) || node.kind === 'textValue') return null;
  return node.localName === 'ins' ? 'insert' : 'delete';
}
