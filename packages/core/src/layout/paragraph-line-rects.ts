// Where whole paragraphs paint, one band per line, in every story.
//
// `paragraphRangeRects` bands the characters of a range. A change that stands for its whole
// paragraph, such as a paragraph formatting change, has no characters of its own, so this
// walk bands each line the paragraph holds: the line's text column at glyph height, and on
// the first line the list number too. Story offsets, frame clipping, and rotated cells come
// from the same helpers, so a line band and a character band on one line agree.

import { forEachStory, type KeyedParagraphRect } from './paragraph-range-rects.ts';
import { lineSegments } from './line-segments.ts';
import { clipParagraphBox } from './paragraph-frame-clip.ts';
import { paragraphFragmentsOfBlocks } from './semantic-records.ts';
import type { LayoutBox, SemanticLayout } from './semantic-records.ts';
import { glyphBand } from './selection-rects.ts';
import { bottomToTopRectInLayout } from './table-cell-text-direction.ts';

/** The horizontal extent of two boxes together. */
function spanOf(a: LayoutBox, b: LayoutBox): { readonly x: number; readonly width: number } {
  const x = Math.min(a.x, b.x);
  return { x, width: Math.max(a.x + a.width, b.x + b.width) - x };
}

/**
 * One band per line of each paragraph in `byParagraph`, once for every key it lists, over
 * the named pages and every story. Coordinates are those of {@link paragraphRangeRects}.
 *
 * A line belongs to each paragraph it holds a segment of, so the join line of two merged
 * paragraphs bands for both. Pass a layout with text box stories projected in to cover text
 * boxes.
 */
export function paragraphLineRects(
  layout: SemanticLayout,
  byParagraph: ReadonlyMap<string, readonly number[]>,
  pages?: ReadonlySet<number>
): KeyedParagraphRect[] {
  const rects: KeyedParagraphRect[] = [];
  if (byParagraph.size === 0) return rects;
  for (const page of layout.pages) {
    if (pages && !pages.has(page.index)) continue;
    forEachStory(page, (blocks, offset) => {
      for (const fragment of paragraphFragmentsOfBlocks(blocks)) {
        for (const [index, line] of fragment.lines.entries()) {
          const segments = lineSegments(line);
          for (const [at, { paragraphId }] of segments.entries()) {
            const keys = byParagraph.get(paragraphId);
            if (!keys) continue;
            // A paragraph that a line holds twice in visual order bands once.
            if (segments.findIndex((other) => other.paragraphId === paragraphId) < at) continue;
            // Layout publishes the marker on the first fragment only, beside its first line.
            const column =
              index === 0 && fragment.marker && paragraphId === fragment.paragraphId
                ? spanOf(line.box, fragment.marker.box)
                : line.box;
            const clipped = clipParagraphBox(
              { pageIndex: page.index, x: column.x, width: column.width, ...glyphBand(line) },
              fragment.clipToBox ? fragment.box : undefined
            );
            if (!clipped) continue;
            const placed = offset
              ? { ...clipped, x: clipped.x + offset.x, y: clipped.y + offset.y }
              : bottomToTopRectInLayout(layout, paragraphId, clipped);
            for (const key of keys) rects.push({ ...placed, key });
          }
        }
      }
    });
  }
  return rects;
}
