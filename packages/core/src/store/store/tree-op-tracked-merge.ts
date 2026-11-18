// Folding of adjacent tracked-change wrappers and text nodes that the tracked edits write.

import type { OoxmlElement, OoxmlNode } from '../package/ooxml-tree.ts';
import { equivalentNodes } from './ooxml-node-equality.ts';
import { wmlAttributeValue } from './tree-op-table-shared.ts';
import { build } from './tree-op-tracked-builders.ts';

/** A `w:t`, or the `w:delText` the same characters become once struck. */
export function textNode(mint: () => string, value: string, deleted: boolean): OoxmlNode {
  const valueId = mint();
  return build(
    mint(),
    deleted ? 'deletedText' : 'text',
    deleted ? 'delText' : 't',
    [],
    [{ id: valueId, kind: 'textValue', value } as OoxmlNode]
  );
}

export function isRunProperties(node: OoxmlNode): boolean {
  return node.kind === 'runProperties';
}

export function coalesced(mint: () => string, nodes: readonly OoxmlNode[]): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const node of nodes) {
    const previous = out[out.length - 1];
    const mergeable =
      previous !== undefined &&
      previous.kind === node.kind &&
      (node.kind === 'text' || node.kind === 'deletedText');
    if (!mergeable) {
      out.push(node);
      continue;
    }

    const left = childrenOf(previous).find((child) => child.kind === 'textValue');
    const right = childrenOf(node).find((child) => child.kind === 'textValue');
    const value =
      (left && left.kind === 'textValue' ? left.value : '') +
      (right && right.kind === 'textValue' ? right.value : '');
    out[out.length - 1] = textNode(mint, value, node.kind === 'deletedText');
  }

  return out;
}

/** A node's children, or none for a text value: the union's only childless member. */
export function childrenOf(node: OoxmlNode): readonly OoxmlNode[] {
  return node.kind === 'textValue' ? [] : node.children;
}

/**
 * Fold adjacent revision wrappers that are the same revision into one.
 *
 * Striking a character at a time leaves `<w:del id=0/><w:del id=0/><w:del id=0/>`: one
 * decision by every reader that groups on the id, but three elements where Word writes one,
 * and three more on the next keystroke. Same id, same author, same date, side by side: one
 * wrapper holding all their runs.
 */
export function mergedRevisions(mint: () => string, nodes: readonly OoxmlNode[]): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  for (const node of nodes) {
    const previous = out[out.length - 1];
    if (
      previous !== undefined &&
      previous.kind !== 'textValue' &&
      node.kind !== 'textValue' &&
      previous.kind === node.kind &&
      (node.kind === 'revisionInsert' || node.kind === 'revisionDelete') &&
      sameRevision(previous, node)
    ) {
      out[out.length - 1] = {
        ...previous,
        children: mergedRevisionRuns(mint, previous.children, node.children),
      } as OoxmlNode;
      continue;
    }

    out.push(node);
  }

  return out;
}

/** Merge the boundary runs when repeated deletion split one formatted run per keypress. */
function mergedRevisionRuns(
  mint: () => string,
  left: readonly OoxmlNode[],
  right: readonly OoxmlNode[]
): OoxmlNode[] {
  const previous = left[left.length - 1];
  const next = right[0];
  if (
    previous?.kind !== 'run' ||
    next?.kind !== 'run' ||
    !mergeableTextRun(previous) ||
    !mergeableTextRun(next)
  ) {
    return [...left, ...right];
  }

  const previousProperties = previous.children.filter(isRunProperties);
  const nextProperties = next.children.filter(isRunProperties);
  if (!equivalentNodes(previousProperties, nextProperties)) return [...left, ...right];
  const content = coalesced(mint, [
    ...previous.children.filter((child) => !isRunProperties(child)),
    ...next.children.filter((child) => !isRunProperties(child)),
  ]);
  const merged = { ...previous, children: [...previousProperties, ...content] } as OoxmlNode;
  return [...left.slice(0, -1), merged, ...right.slice(1)];
}

/** Only plain text runs are safe to collapse; drawings, breaks and generic atoms stay separate. */
function mergeableTextRun(run: OoxmlElement): boolean {
  const content = run.children.filter((child) => !isRunProperties(child));
  return (
    content.length > 0 &&
    content.every((child) => child.kind === 'text' || child.kind === 'deletedText')
  );
}

/** Two wrappers are the same revision when their `CT_TrackChange` triple agrees. */
function sameRevision(a: OoxmlElement, b: OoxmlElement): boolean {
  const id = wmlAttributeValue(a, 'id');
  return (
    id !== undefined &&
    id === wmlAttributeValue(b, 'id') &&
    wmlAttributeValue(a, 'author') === wmlAttributeValue(b, 'author') &&
    wmlAttributeValue(a, 'date') === wmlAttributeValue(b, 'date')
  );
}
