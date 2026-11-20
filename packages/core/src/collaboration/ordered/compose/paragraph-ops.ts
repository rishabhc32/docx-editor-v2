// A paragraph property press, run by the engine's own op on a scratch part that holds only the
// paragraph being drawn. A decided press runs with no tracking; a pending one with tracking, so
// the engine writes its own `w:pPrChange` (one record, the original under the latest author).

import { findNode } from '../../../store/package/ooxml-edit.ts';
import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import {
  directParagraphProperties,
  mergedProperties,
} from '../../../store/store/direct-properties.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import { cloneWithNewIds } from '../../../store/store/tree-op-nodes.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type {
  OoxmlProperty,
  RevisionAttributionInput,
  TreeDocOp,
} from '../../../store/store/tree-op-types.ts';
import type { OrderedParagraphOp } from '../model/types.ts';
import { childWhere, isParagraphProperties, isWml, withChildren } from '../ooxml-nodes.ts';

/** `part` with these blocks in place of its own: a document's body children, or the root's. */
export function withStoryBlocks(part: OoxmlPart, blocks: readonly OoxmlNode[]): OoxmlPart {
  const root = part.root;
  if (root.localName !== 'document') return { ...part, root: withChildren(root, blocks) };
  return {
    ...part,
    root: withChildren(
      root,
      root.children.map((child) =>
        child.kind !== 'textValue' && isWml(child, 'body') ? withChildren(child, blocks) : child
      )
    ),
  };
}

/** Properties without those a press takes off. */
export function withoutRemoved(
  properties: readonly OoxmlProperty[],
  remove: readonly string[] = []
): OoxmlProperty[] {
  return properties.filter((property) => !remove.includes(property.localName));
}

/** The engine's op for a press on the paragraph `paragraphId` of `scratch`. */
function treeOpFor(
  scratch: OoxmlPart,
  paragraphId: string,
  op: OrderedParagraphOp,
  revision: RevisionAttributionInput | undefined
): TreeDocOp {
  const tracked = revision === undefined ? {} : { revision };
  switch (op.op) {
    case 'setParagraphProperties':
      // A press writes the whole set: what the paragraph has, with the press merged in and
      // what it takes off left out, as the editor's paragraph command writes it
      return {
        op: 'setParagraphProperties',
        paragraphId,
        properties: withoutRemoved(
          mergedProperties(directParagraphProperties(scratch, paragraphId), op.properties),
          op.remove
        ),
        ...tracked,
      };
    case 'setListLevel':
      return { op: 'setListLevel', paragraphId, level: op.level, ...tracked };
    case 'setListNumbering':
      return {
        op: 'setListNumbering',
        paragraphId,
        numId: op.numId,
        ...(op.level === undefined ? {} : { level: op.level }),
        ...tracked,
      };
    case 'setParagraphTabStops':
      return {
        op: 'setParagraphTabStops',
        paragraphId,
        stops: op.stops,
        ...(op.inForcePositionsTwips === undefined
          ? {}
          : { inForcePositionsTwips: op.inForcePositionsTwips }),
        ...tracked,
      };
  }
}

/** A tracked press: the suggestion's attribution and its record's `w:id`. */
interface TrackedPress {
  readonly revision: RevisionAttributionInput;
  readonly revisionId: string;
}

/** The engine's result of one press on a scratch part that holds only this paragraph. */
function runOnScratch(
  part: OoxmlPart,
  pPr: OoxmlElement | null,
  op: OrderedParagraphOp,
  mint: () => string,
  tracked: TrackedPress | null
): { readonly result: ReturnType<typeof applyTreeOp>; readonly paragraphId: string } {
  const paragraph = build(mint(), 'paragraph', 'p', [], pPr ? [pPr] : []);
  const scratch = withStoryBlocks(part, [paragraph]);
  const result = applyTreeOp(
    scratch,
    treeOpFor(scratch, paragraph.id, op, tracked?.revision),
    tracked
      ? { revisionIds: () => tracked.revisionId, deferValidation: true }
      : { deferValidation: true }
  );
  return { result, paragraphId: paragraph.id };
}

/**
 * A paragraph's `w:pPr` after one press, or `pPr` itself when the press does not apply (a
 * list level once the list is gone). `tracked` gives the suggestion's attribution and the
 * record's `w:id`.
 */
export function withParagraphOp(
  part: OoxmlPart,
  pPr: OoxmlElement | null,
  op: OrderedParagraphOp,
  mint: () => string,
  tracked: TrackedPress | null
): OoxmlElement | null {
  const { result, paragraphId } = runOnScratch(part, pPr, op, mint, tracked);
  if (!result.ok) return pPr;
  const written = findNode(result.part, paragraphId);
  const after =
    written && written.kind !== 'textValue'
      ? childWhere(written, isParagraphProperties)
      : undefined;
  if (!after) return null;

  // New ids: the scratch part's minter does not know the ids the story already holds
  const copy = cloneWithNewIds(after, mint);
  return copy.kind === 'textValue' ? null : copy;
}

/**
 * Why the engine refuses this press on the paragraph, or null when it runs: a check for a caller
 * that refuses what the engine refuses.
 */
export function paragraphOpRefusal(
  part: OoxmlPart,
  pPr: OoxmlElement | null,
  op: OrderedParagraphOp,
  mint: () => string
): string | null {
  const { result } = runOnScratch(part, pPr, op, mint, null);
  return result.ok ? null : result.reason;
}

/**
 * Whether the engine's op for a press by `revision`'s author leaves the paragraph as it shows
 * (`pPr`, with any change record): the paragraph has every value the press sets, and the engine
 * writes no record for it. Word records none either.
 */
export function paragraphOpKeeps(
  part: OoxmlPart,
  pPr: OoxmlElement | null,
  op: OrderedParagraphOp,
  mint: () => string,
  revision: RevisionAttributionInput
): boolean {
  const after = withParagraphOp(part, pPr, op, mint, { revision, revisionId: mint() });
  return after === null || pPr === null ? after === pPr : ooxmlTreesEqual(after, pPr);
}
