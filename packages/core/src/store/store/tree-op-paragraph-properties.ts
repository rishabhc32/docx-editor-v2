// Paragraph property ops: `w:pPr` and the paragraph mark's `w:pPr/w:rPr`. Each write is merged
// over what its container already holds (`mergedPropertyChildren`) and, in suggesting mode,
// recorded so Reject restores what it replaced (`withPropertyChangeRecord`).

import type { OoxmlNode, OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import {
  insertChildren,
  removeNode,
  replaceChildren,
  replaceNode,
  type EditOptions,
} from '../package/ooxml-edit.ts';
import type {
  OoxmlProperty,
  RevisionAttributionInput,
  TreeOpEffect,
  TreeOpResult,
} from './tree-op-types.ts';
import { ownProposedParagraph, withPropertyChangeRecord } from './tree-op-tracked-properties.ts';
import { build } from './tree-op-tracked.ts';
import { nextRevisionId } from './tree-op-revision-ids.ts';
import { TEXT_DEPS, fromEdit, namedChild, ok, paragraphPropertiesNodeOf } from './tree-op-nodes.ts';
import {
  PARAGRAPH_VOCABULARY,
  RUN_VOCABULARY,
  UNCHANGED_PROPERTIES_EFFECT,
  keepsPropertyChildren,
  mergedPropertyChildren,
} from './tree-op-properties.ts';

/**
 * Write a paragraph's own properties (`w:pPr`, 17.3.1.26): what the op names is rewritten, and
 * what the op vocabulary cannot name (`w:sectPr`, the mark, `w:pBdr`) stays as authored.
 */
export function applySetParagraphProperties(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  properties: readonly OoxmlProperty[],
  options: EditOptions | undefined,
  nextId: () => string,
  revision?: RevisionAttributionInput
): TreeOpResult {
  const existing = paragraphPropertiesNodeOf(paragraph);
  const prior = existing?.children ?? [];
  let children: readonly OoxmlNode[] = mergedPropertyChildren(
    prior,
    properties,
    PARAGRAPH_VOCABULARY,
    nextId
  );

  // A tracked paragraph-property change records what `w:pPr` held in a `w:pPrChange`
  // (§17.13.5.29), so Reject restores the alignment, indents and style it replaced.
  //
  // Skipped on a paragraph whose mark THIS author proposed adding, the same rule the run
  // and mark writes apply: rejecting that `w:ins` runs the paragraph into the next one and
  // takes its properties with it, so a record of what they used to be decides nothing.
  const markProperties = namedChild(existing, 'rPr')?.children ?? [];
  if (revision && !ownProposedParagraph(part, paragraph.id, markProperties, revision.author)) {
    children = withPropertyChangeRecord({
      container: 'paragraphProperties',
      prior,
      next: children,
      revision: revision,
      mint: nextId,
      // The transaction's minter when it lent one, else a lazy walk of this part.
      nextRevisionId: () => options?.revisionIds?.() ?? nextRevisionId(part)(),
    });
  }

  // A write that changes no value leaves the paragraph exactly as it was: Word makes no change
  // for such a press, and rewriting the container in schema order would change the saved
  // package while the paragraph stayed the same.
  if (keepsPropertyChildren(prior, children)) return ok(part, UNCHANGED_PROPERTIES_EFFECT);
  const effect: TreeOpEffect = {
    dirty: [paragraph.id],
    created: [],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    impact: 'paragraph-local',
  };
  if (children.length === 0) {
    // Nothing left to hold: the container goes rather than staying as an empty `w:pPr`,
    // so a cleared paragraph digests identically to one that never had any.
    return existing ? fromEdit(removeNode(part, existing.id, options), effect) : ok(part, effect);
  }

  if (existing) return fromEdit(replaceChildren(part, existing.id, children, options), effect);
  const pPr = build(nextId(), 'paragraphProperties', 'pPr', [], children);

  // `w:pPr` must be the paragraph's FIRST child per the schema.
  return fromEdit(insertChildren(part, paragraph.id, 0, [pPr], options), effect);
}

/**
 * `CT_PPr` child order (17.3.1.26). `w:rPr` sits after every base property and before
 * `w:sectPr`, so a mark written at the end of `w:pPr` would land on the wrong side of a
 * section mark and Word's reader rejects it.
 */
const AFTER_PARAGRAPH_MARK = new Set(['sectPr', 'pPrChange']);

/**
 * Write the run properties of the paragraph MARK (`w:pPr/w:rPr`, 17.3.1.29).
 *
 * Surgical on `w:rPr`, like `setListNumbering` is on `w:numPr`: the mark is one nested
 * child of `w:pPr` and everything else there (the style, the numbering, the borders)
 * has to survive. It cannot go through `setParagraphProperties` at all, because that takes
 * a FLAT property bag and the mark's own run properties are children.
 *
 * An empty `properties` removes the element rather than leaving an empty `w:rPr`, so a
 * cleared mark digests identically to one that never had any, unless a tracked format
 * change put a `w:rPrChange` there, which is the one child that keeps the container alive
 * with nothing else in it.
 */
export function applySetParagraphMarkProperties(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  properties: readonly OoxmlProperty[],
  options: EditOptions | undefined,
  nextId: () => string,
  revision?: RevisionAttributionInput
): TreeOpResult {
  const effect: TreeOpEffect = {
    dirty: [paragraph.id],
    created: [],
    deleted: [],
    dependencyKeys: TEXT_DEPS,
    // The mark is what a list marker inherits its face from, so a change here re-resolves
    // the marker and can change its measured width.
    impact: 'paragraph-local',
  };
  const pPr = paragraphPropertiesNodeOf(paragraph);
  const existing = namedChild(pPr, 'rPr');

  // The mark is a run property container like any other: what the op names is rewritten,
  // and what it cannot name (a `w:rStyle` character style, `w:lang`, a revision record)
  // stays where it was authored.
  const prior = existing?.children ?? [];
  let children: readonly OoxmlNode[] = mergedPropertyChildren(
    prior,
    properties,
    RUN_VOCABULARY,
    nextId
  );

  // The mark's tracked format change is `w:pPr/w:rPr/w:rPrChange` (§17.13.5.31): the same
  // element a run writes, over `CT_ParaRPrOriginal` rather than `CT_RPrOriginal`. It is what
  // keeps a list marker's face reversible when a whole-paragraph format is proposed.
  //
  // Skipped on a mark THIS author proposed adding, the run rule one level up: rejecting that
  // `w:ins` runs the paragraph into the next one and takes the mark's properties with it, so
  // a record of what they used to be decides nothing. The mark never existed for anyone else.
  if (revision && !ownProposedParagraph(part, paragraph.id, prior, revision.author)) {
    children = withPropertyChangeRecord({
      container: 'runProperties',
      prior,
      next: children,
      revision,
      mint: nextId,
      // The transaction's minter when it lent one, else a lazy walk of this part.
      nextRevisionId: () => options?.revisionIds?.() ?? nextRevisionId(part)(),
    });
  }

  // A mark that already says all of this keeps its `w:rPr` as authored.
  if (keepsPropertyChildren(prior, children)) return ok(part, UNCHANGED_PROPERTIES_EFFECT);
  if (children.length === 0) {
    if (!existing) return ok(part, effect);
    return fromEdit(removeNode(part, existing.id, options), effect);
  }

  // THE TYPED KIND, and an EXISTING container rewritten in place: the same rule
  // `applySetRunProperties` states for a run's own `w:rPr`, and the one the `w:pPr` branch
  // below states for the paragraph. Layout finds the mark's revisions and its format record by
  // `kind === 'runProperties'` with no name fallback.
  const rPr = existing
    ? ({ ...existing, children } as OoxmlNode)
    : build(nextId(), 'runProperties', 'rPr', [], children);
  if (existing) return fromEdit(replaceNode(part, existing.id, rPr, options), effect);
  if (pPr) {
    const before = pPr.children.findIndex(
      (child) => child.kind !== 'textValue' && AFTER_PARAGRAPH_MARK.has(child.localName)
    );
    const index = before === -1 ? pPr.children.length : before;
    return fromEdit(insertChildren(part, pPr.id, index, [rPr], options), effect);
  }

  // The TYPED kind: every reader finds paragraph properties by
  // `kind === 'paragraphProperties'`, so a generic `w:pPr` is invisible to all of them.
  const created = build(nextId(), 'paragraphProperties', 'pPr', [], [rPr]);

  // `w:pPr` must be the paragraph's FIRST child per the schema.
  return fromEdit(insertChildren(part, paragraph.id, 0, [created], options), effect);
}
