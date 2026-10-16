// Tracked TAB STOPS — `w:tabs` recorded in `w:pPrChange` like every other paragraph property.
//
// `w:tabs` is a `CT_PPrBase` child, so Word records a tab-stop change the way it records an
// indent: the paragraph keeps its new `w:tabs`, and a `w:pPrChange` holds the `w:pPr` as it
// was. The untracked write stays the one in `tree-op-section.ts`; this only adds the record,
// through the same helper the list ops use, so one dialog OK is one card whatever it changed.

import type { EditOptions } from '../package/ooxml-edit.ts';
import type { OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { invalidRevisionAttribution } from './tree-op-revision-attribution.ts';
import { applySetParagraphTabStops, trackListProperties } from './tree-op-section.ts';
import type { TreeDocOp, TreeOpResult } from './tree-op-types.ts';

/** `setParagraphTabStops`, recorded as a `w:pPrChange` when the op carries a revision. */
export function applyParagraphTabStopsOp(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  op: Extract<TreeDocOp, { op: 'setParagraphTabStops' }>,
  options: EditOptions | undefined,
  nextId: () => string
): TreeOpResult {
  if (op.revision !== undefined && invalidRevisionAttribution(op.revision)) {
    return { ok: false, reason: 'invalid-property-value' };
  }
  const written = applySetParagraphTabStops(
    part,
    paragraph,
    op.stops,
    op.inForcePositionsTwips,
    options,
    nextId
  );
  return trackListProperties(part, paragraph, written, op.revision, options);
}
