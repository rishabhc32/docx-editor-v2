// Comparisons for the ordered model's tests: the repository's fidelity oracle
// (`canonicalOoxmlFingerprint`), with a readable first difference when two parts disagree.

import { expect } from 'bun:test';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
  validateOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';

/** The first place two strings differ, with some context on each side. */
export function firstDifference(left: string, right: string, context = 160): string | null {
  if (left === right) return null;
  let at = 0;
  while (at < left.length && left[at] === right[at]) at += 1;
  const from = Math.max(0, at - context);
  return `got:  …${left.slice(from, at + context)}…\nwant: …${right.slice(from, at + context)}…`;
}

/** Fail with the first difference of the serialized parts when the fingerprints differ. */
export function expectSameFingerprint(actual: OoxmlPart, expected: OoxmlPart, label = ''): void {
  if (canonicalOoxmlFingerprint(actual) === canonicalOoxmlFingerprint(expected)) return;
  const difference = firstDifference(serializeOoxmlPart(actual), serializeOoxmlPart(expected));
  expect(`${label}\n${difference ?? 'the fingerprints differ, the XML does not'}`).toBe(label);
}

/** Fail when a composed part breaks the tree's invariants. */
export function expectValid(part: OoxmlPart, label = ''): void {
  const result = validateOoxmlPart(part);
  expect({ label, issues: result.ok ? [] : result.issues.slice(0, 3) }).toEqual({
    label,
    issues: [],
  });
}

/** Every node id in a subtree, for duplicate checks. */
export function nodeIds(node: OoxmlNode, into: string[] = []): string[] {
  into.push(node.id);
  if (node.kind !== 'textValue') for (const child of node.children) nodeIds(child, into);
  return into;
}
