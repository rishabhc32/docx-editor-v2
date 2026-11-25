// Comparisons for the ordered model's tests: the repository's fidelity oracle
// (`canonicalOoxmlFingerprint`), with a readable first difference when two parts disagree.

import { expect } from 'bun:test';
import { isRelationshipReference } from '../../../store/package/canonical-media-references.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { relationshipsOf } from '../../../store/package/package-edit.ts';
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

/**
 * Each relationship id a part of a read package names that the part's own relationships do not
 * hold, as `part id`. Ids are per part (OPC): Word may refuse a file that names a missing one.
 */
export function unresolvedRelationships(pkg: OoxmlPackage): string[] {
  const out: string[] = [];
  for (const part of pkg.parts.values()) {
    const held = new Set(relationshipsOf(pkg, part.name).map((record) => record.id));
    const visit = (node: OoxmlNode): void => {
      if (node.kind === 'textValue') {
        return;
      }

      for (const attribute of node.attributes) {
        if (isRelationshipReference(attribute) && !held.has(attribute.value)) {
          out.push(`${part.name} ${attribute.value}`);
        }
      }

      node.children.forEach(visit);
    };
    visit(part.root);
  }

  return out;
}
