// A region of a story: the top-level blocks a decision's check draws (`ComposeOptions.region`).
// The engine plans a decision from the blocks its items are in, a paragraph join reads the block
// after one, and a part with a move has every item planned with its dependencies, so a block that
// holds a move draws too. The rest of the story is passed over: the check costs what its cards
// hold, not what the document holds.

import type { OrderedItem, OrderedStory } from '../model/types.ts';

export class RegionWalk {
  /** The last top-level block that wrote a node held the region: the next one draws too. */
  #after = false;

  constructor(private readonly items: ReadonlySet<OrderedItem>) {}

  /** Whether any of these items is one of the region's, or a move's. */
  holds(items: Iterable<OrderedItem>): boolean {
    for (const item of items) {
      if (this.items.has(item) || item.move !== null) {
        return true;
      }
    }

    return false;
  }

  /** Whether the block from `open` to its close holds the region: a look over its items. */
  holdsBlock(story: OrderedStory, open: OrderedItem): boolean {
    for (let at: OrderedItem | null = open; at; at = story.tree.next(at)) {
      if (this.holds([at])) {
        return true;
      }

      if (at !== open && at.kind === 'close' && at.pair === open.pair) {
        return false;
      }
    }

    return false;
  }

  /** Whether a top-level block draws: it holds the region, or it follows one that does. */
  draws(own: boolean): boolean {
    return own || this.#after;
  }

  /** A top-level block wrote a node: the next block draws if this one held the region. */
  wrote(own: boolean): void {
    this.#after = own;
  }
}
