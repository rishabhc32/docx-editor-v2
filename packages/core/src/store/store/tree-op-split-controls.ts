// Enter inside a content control that holds one line, or one paragraph.

import {
  contentControlPropertiesOf,
  type ContentControlProperties,
} from '../package/content-control-nodes.ts';
import type { OoxmlParagraphNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { propertiesHoldLineBreaks } from './tree-op-inline-elements.ts';
import { contentControlAncestorsOf } from './tree-op-nodes.ts';
import { segmentsOf } from './tree-op-segments.ts';
import type { TreeOpRejection } from './tree-op-types.ts';

type DividedProperties = Pick<ContentControlProperties, 'type' | 'multiLine'>;

/**
 * Whether Enter may divide an inline control of these properties: one that holds line breaks,
 * but not rich text, typed or untyped, which Word leaves whole. A multi-line plain text control
 * and the container kinds (a group, a building block) still split; Word is not checked there.
 *
 * Word refuses Enter in a plain text control without `w:multiLine`, and in a list, a date, a
 * check box or a picture control: their value is one line. Word also ignores Enter inside an
 * inline rich text control, typed or untyped (Word 16.114, 2026-10-30, samples
 * `~$i2-split.docx` and `~$i2n-split.docx`: no new paragraph and no revision, and typing at the
 * same caret works): an inline control lives in one paragraph.
 */
export function enterDividesControl(properties: DividedProperties): boolean {
  if (properties.type === 'richText' || properties.type === 'untyped') return false;
  return propertiesHoldLineBreaks(properties);
}

/**
 * `single-line-control` when a split at one of `offsets` divides a control that `divides` says
 * the split may not: by default one that holds one line, and for Enter `enterDividesControl`.
 * A split divides an inline control when units of the control are on both sides of it; a split
 * at the control's edge goes beside it. A paste of several paragraphs keeps the line rule only
 * (`propertiesHoldLineBreaks`): Word's paste there is not checked.
 */
export function splitDividesOneLineControl(
  part: OoxmlPart,
  paragraph: OoxmlParagraphNode,
  offsets: readonly number[],
  divides: (properties: DividedProperties) => boolean = propertiesHoldLineBreaks
): TreeOpRejection | null {
  const segments = segmentsOf(paragraph);

  // Block controls around the paragraph hold paragraphs: a split stays inside them
  const around = new Set(contentControlAncestorsOf(part, paragraph.id).map((c) => c.id));
  for (const offset of offsets) {
    const before = segments.find((s) => s.start < offset && offset <= s.end);
    const after = segments.find((s) => s.start <= offset && offset < s.end);
    if (!before || !after) continue;

    // An atom (a simple field) has no run of its own: its node stands for it
    const within = (segment: typeof before) =>
      contentControlAncestorsOf(part, segment.runId || segment.node.id);
    const holdsAfter = new Set(within(after).map((c) => c.id));
    for (const control of within(before)) {
      if (around.has(control.id) || !holdsAfter.has(control.id)) continue;
      if (!divides(contentControlPropertiesOf(control))) return 'single-line-control';
    }
  }

  return null;
}
