// Where whole paragraphs paint, one band per line, in every story. Story offsets, frame
// clipping and rotated cells come from the helpers `paragraphRangeRects` uses, so a line band
// and a character band on one line agree.

import { forEachStory, type KeyedParagraphRect } from './paragraph-range-rects.ts';
import { lineSegments } from './line-segments.ts';
import { clipParagraphBox } from './paragraph-frame-clip.ts';
import { paragraphFragmentsOfBlocks } from './semantic-records.ts';
import type {
  LayoutBox,
  LineRecord,
  ParagraphFragmentRecord,
  SemanticLayout,
} from './semantic-records.ts';
import { glyphBand } from './selection-rects.ts';
import { bottomToTopRectInLayout } from './table-cell-text-direction.ts';

/** The horizontal extent of two boxes together. */
function spanOf(a: LayoutBox, b: LayoutBox): { readonly x: number; readonly width: number } {
  const x = Math.min(a.x, b.x);
  return { x, width: Math.max(a.x + a.width, b.x + b.width) - x };
}

/**
 * The band of `line` for `paragraphId`: the line's text column at glyph height, with the list
 * number on the first line, placed in the story. Null when the fragment's frame clips it away.
 */
function lineBand(
  layout: SemanticLayout,
  pageIndex: number,
  fragment: ParagraphFragmentRecord,
  line: LineRecord,
  paragraphId: string,
  offset: { readonly x: number; readonly y: number } | null
): (LayoutBox & { readonly pageIndex: number }) | null {
  // Layout publishes the marker on the first fragment only, beside its first line.
  const column =
    line === fragment.lines[0] && fragment.marker && paragraphId === fragment.paragraphId
      ? spanOf(line.box, fragment.marker.box)
      : line.box;
  const clipped = clipParagraphBox(
    { pageIndex, x: column.x, width: column.width, ...glyphBand(line) },
    fragment.clipToBox ? fragment.box : undefined
  );
  if (!clipped) return null;
  return offset
    ? { ...clipped, x: clipped.x + offset.x, y: clipped.y + offset.y }
    : bottomToTopRectInLayout(layout, paragraphId, clipped);
}

/**
 * One band per line of each paragraph in `byParagraph`, once for every key it lists, over
 * the named pages and every story. Coordinates are those of {@link paragraphRangeRects}.
 * A change that stands for its whole paragraph, such as a paragraph formatting change, has no
 * characters of its own, so it bands each line the paragraph holds.
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
        for (const line of fragment.lines) {
          const segments = lineSegments(line);
          for (const [at, { paragraphId }] of segments.entries()) {
            const keys = byParagraph.get(paragraphId);
            if (!keys) continue;

            // A paragraph that a line holds twice in visual order bands once.
            if (segments.findIndex((other) => other.paragraphId === paragraphId) < at) continue;
            const placed = lineBand(layout, page.index, fragment, line, paragraphId, offset);
            if (!placed) continue;
            for (const key of keys) rects.push({ ...placed, key });
          }
        }
      }
    });
  }

  return rects;
}
