// The links and controls of the ordered view: the segments a view character's containers name,
// and the close item that ends a container.

import { isHidden, itemAt } from './model/document.ts';
import type { OrderedDocument, OrderedItem, OrderedStory } from './model/types.ts';

/** The segments of a view character's containers (`h:<space>`, `c:<space>`), outermost first. */
export function segmentsOf(containers: string): string[] {
  return containers === '' ? [] : containers.split('>');
}

/** The innermost segment of a view character's containers, or '' outside every container. */
export function innermostSegment(containers: string): string {
  return segmentsOf(containers).at(-1) ?? '';
}

/** Whether a segment names a control (`c:`), not a link (`h:`). */
export const isControlSegment = (segment: string): boolean => segment.startsWith('c:');

/** Whether a segment names a link (`h:`). */
export const isLinkSegment = (segment: string): boolean => segment.startsWith('h:');

/** The open item of the container a segment names, or null. */
export function openOfSegment(story: OrderedStory, segment: string): OrderedItem | null {
  return itemAt(story, { space: segment.slice(2), offset: 0 });
}

/** How many outer segments two container lists share. */
export function sharedDepth(a: readonly string[], b: readonly string[]): number {
  let depth = 0;
  while (depth < a.length && a[depth] === b[depth]) {
    depth += 1;
  }

  return depth;
}

/**
 * Whether an item closes the container `open` starts. Enter inside a link or a control ends it
 * with a close item of its own (its space is not the container's) and starts the rest with an
 * open item that pairs with the container's close; when that Enter does not show (rejected or
 * taken back), its close item ends nothing and the container is one again.
 */
export function closesContainer(
  item: OrderedItem,
  open: OrderedItem,
  doc: OrderedDocument | undefined
): boolean {
  if (item.kind !== 'close' || item.pair !== open.pair) return false;
  return item.space === item.pair || doc === undefined || !isHidden(doc, item);
}
