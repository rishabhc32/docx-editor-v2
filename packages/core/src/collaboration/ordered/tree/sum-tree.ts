// A B+ tree of items with summaries, after Zed's `SumTree`.

/** How to sum item summaries. `add` mutates `into`. */
export interface SummaryOps<T, S> {
  zero(): S;
  add(into: S, other: Readonly<S>): void;
  of(item: T): S;
}

/** An item the tree can hold. The tree alone writes `leaf`. */
export interface SumTreeItem<T, S> {
  leaf: SumTreeLeaf<T, S> | undefined;
}

/** Most items in one leaf. */
export const LEAF_MAX = 64;

/** Most children of one inner node. */
export const INNER_MAX = 32;

/** The fill a bulk build and a split aim for, so the next inserts do not split again at once. */
const LEAF_FILL = (LEAF_MAX * 3) >> 2;
const INNER_FILL = (INNER_MAX * 3) >> 2;

/** A leaf of a {@link SumTree}. Read-only outside the tree. */
export class SumTreeLeaf<T, S> {
  readonly kind = 'leaf' as const;
  readonly tree: object;
  items: T[] = [];
  sum: S;
  parent: SumTreeInner<T, S> | null = null;
  prev: SumTreeLeaf<T, S> | null = null;
  next: SumTreeLeaf<T, S> | null = null;

  constructor(tree: object, sum: S) {
    this.tree = tree;
    this.sum = sum;
  }
}

/** An inner node of a {@link SumTree}. Read-only outside the tree. */
export class SumTreeInner<T, S> {
  readonly kind = 'inner' as const;
  children: SumTreeNode<T, S>[] = [];
  sum: S;
  parent: SumTreeInner<T, S> | null = null;

  constructor(sum: S) {
    this.sum = sum;
  }
}

/** A node of a {@link SumTree}: a leaf or an inner node. */
export type SumTreeNode<T, S> = SumTreeLeaf<T, S> | SumTreeInner<T, S>;

/** Where a seek stopped: the item whose span holds the target, and the sum before it. */
export interface SumTreeSeek<T, S> {
  readonly item: T;
  readonly before: S;
}

/** Sizes of a tree, for memory figures. */
export interface SumTreeStats {
  readonly depth: number;
  readonly leaves: number;
  readonly inners: number;
  readonly items: number;
}

/**
 * A summed B+ tree of items in order. Leaves hold the items in order. Every node keeps the sum
 * of its subtree's item summaries, so a seek by a summary measure, the sum before an item, and a
 * summary update all take O(log n). Items are mutable objects that the tree knows by identity:
 * each item keeps a pointer to the leaf that holds it, which the tree alone writes.
 *
 * The tree does not merge underfull nodes. Removal takes back an insertion or a split, or drops
 * items a replace re-cut, so the tree never shrinks far below the size it reached.
 */
export class SumTree<T extends SumTreeItem<T, S>, S> {
  private root: SumTreeNode<T, S>;
  private count = 0;
  private readonly ops: SummaryOps<T, S>;

  constructor(ops: SummaryOps<T, S>, items: readonly T[] = []) {
    this.ops = ops;
    this.root = this.build(items);
  }

  /** The sum of every item's summary. */
  get summary(): Readonly<S> {
    return this.root.sum;
  }

  /** The number of items. */
  get size(): number {
    return this.count;
  }

  // ---- reading ----

  /** Whether `item` is in this tree. */
  has(item: T): boolean {
    const leaf = item.leaf;
    return leaf !== undefined && leaf.tree === this && leaf.items.includes(item);
  }

  first(): T | null {
    let node = this.root;
    while (node.kind === 'inner') node = this.childAt(node, 0);
    for (let leaf: SumTreeLeaf<T, S> | null = node; leaf; leaf = leaf.next) {
      const item = leaf.items[0];
      if (item !== undefined) return item;
    }

    return null;
  }

  last(): T | null {
    let node = this.root;
    while (node.kind === 'inner') node = this.childAt(node, node.children.length - 1);
    for (let leaf: SumTreeLeaf<T, S> | null = node; leaf; leaf = leaf.prev) {
      const item = leaf.items[leaf.items.length - 1];
      if (item !== undefined) return item;
    }

    return null;
  }

  next(item: T): T | null {
    const leaf = this.leafOf(item);
    const after = leaf.items[leaf.items.indexOf(item) + 1];
    if (after !== undefined) return after;
    for (let other = leaf.next; other; other = other.next) {
      const first = other.items[0];
      if (first !== undefined) return first;
    }

    return null;
  }

  prev(item: T): T | null {
    const leaf = this.leafOf(item);
    const at = leaf.items.indexOf(item);
    if (at > 0) return leaf.items[at - 1] ?? null;
    for (let other = leaf.prev; other; other = other.prev) {
      const last = other.items[other.items.length - 1];
      if (last !== undefined) return last;
    }

    return null;
  }

  /** The items in order, from `from` (included) or from the first one. */
  *items(from?: T | null): Generator<T> {
    let item = from === undefined ? this.first() : from;
    if (item !== null) this.leafOf(item);
    while (item !== null) {
      yield item;
      item = this.next(item);
    }
  }

  /** The items in reverse order, from `from` (included) or from the last one. */
  *itemsBackward(from?: T | null): Generator<T> {
    let item = from === undefined ? this.last() : from;
    if (item !== null) this.leafOf(item);
    while (item !== null) {
      yield item;
      item = this.prev(item);
    }
  }

  toArray(): T[] {
    return [...this.items()];
  }

  /** The sum of the summaries of every item before `item`. */
  before(item: T): S {
    const sum = this.ops.zero();
    const leaf = this.leafOf(item);
    for (const other of leaf.items) {
      if (other === item) break;
      this.ops.add(sum, this.ops.of(other));
    }

    let child: SumTreeNode<T, S> = leaf;
    for (let node = leaf.parent; node; child = node, node = node.parent) {
      for (const sibling of node.children) {
        if (sibling === child) break;
        this.ops.add(sum, sibling.sum);
      }
    }

    return sum;
  }

  /**
   * The first item at which the running total of `measure` passes `target`: the item whose span
   * holds `target`, and the sum before it. Null when the total never passes `target`.
   *
   * `measure` must be additive (`measure(a + b) = measure(a) + measure(b)`) and never negative,
   * as a count field of the summary is.
   */
  seek(measure: (summary: Readonly<S>) => number, target: number): SumTreeSeek<T, S> | null {
    if (!(measure(this.root.sum) > target)) return null;
    const before = this.ops.zero();
    let total = 0;
    let node = this.root;
    while (node.kind === 'inner') {
      let chosen: SumTreeNode<T, S> | null = null;
      for (const child of node.children) {
        const width = measure(child.sum);
        if (total + width > target) {
          chosen = child;
          break;
        }

        total += width;
        this.ops.add(before, child.sum);
      }

      if (chosen === null) return null;
      node = chosen;
    }

    for (const item of node.items) {
      const summary = this.ops.of(item);
      const width = measure(summary);
      if (total + width > target) return { item, before };
      total += width;
      this.ops.add(before, summary);
    }

    return null;
  }

  // ---- writing ----

  /** The item's summary changed: the sums up the tree are computed again. */
  update(item: T): void {
    this.resumUp(this.leafOf(item));
  }

  /** Put `items` right after `after`, or at the very start when `after` is null. */
  insertAfter(after: T | null, items: readonly T[]): void {
    if (items.length === 0) return;
    for (const item of items) {
      if (item.leaf !== undefined) throw new RangeError('The item is already in a tree.');
    }

    let leaf: SumTreeLeaf<T, S>;
    let at: number;
    if (after !== null) {
      leaf = this.leafOf(after);
      at = leaf.items.indexOf(after) + 1;
    } else {
      let node = this.root;
      while (node.kind === 'inner') node = this.childAt(node, 0);
      leaf = node;
      at = 0;
    }

    leaf.items = [...leaf.items.slice(0, at), ...items, ...leaf.items.slice(at)];
    for (const item of items) item.leaf = leaf;
    this.count += items.length;
    this.settle(leaf);
  }

  /** Replace one item by others in its place (a fragment split in pieces). */
  replace(item: T, pieces: readonly T[]): void {
    if (pieces.length === 0) {
      this.remove(item);
      return;
    }

    for (const piece of pieces) {
      if (piece.leaf !== undefined) throw new RangeError('The piece is already in a tree.');
    }

    const leaf = this.leafOf(item);
    const at = leaf.items.indexOf(item);
    leaf.items = [...leaf.items.slice(0, at), ...pieces, ...leaf.items.slice(at + 1)];
    item.leaf = undefined;
    for (const piece of pieces) piece.leaf = leaf;
    this.count += pieces.length - 1;
    this.settle(leaf);
  }

  /** Take an item out (to undo its insertion or a split). */
  remove(item: T): void {
    const leaf = this.leafOf(item);
    leaf.items.splice(leaf.items.indexOf(item), 1);
    item.leaf = undefined;
    this.count -= 1;
    if (leaf.items.length > 0 || (leaf.prev === null && leaf.next === null)) {
      this.resumUp(leaf);
      return;
    }

    // An emptied leaf leaves the chain and its parent; an emptied parent goes too
    if (leaf.prev) leaf.prev.next = leaf.next;
    if (leaf.next) leaf.next.prev = leaf.prev;
    leaf.prev = null;
    leaf.next = null;
    let child: SumTreeNode<T, S> = leaf;
    let parent = leaf.parent;
    while (parent !== null) {
      parent.children.splice(parent.children.indexOf(child), 1);
      child.parent = null;
      if (parent.children.length > 0) break;
      child = parent;
      parent = parent.parent;
    }

    if (parent !== null) this.resumUp(parent);

    // A root with one inner child gives way to that child
    while (this.root.kind === 'inner' && this.root.children.length === 1) {
      const only = this.childAt(this.root, 0);
      only.parent = null;
      this.root = only;
    }
  }

  // ---- checks ----

  /** Depth and node counts, for memory figures. */
  stats(): SumTreeStats {
    let depth = 1;
    for (let node = this.root; node.kind === 'inner'; node = this.childAt(node, 0)) depth += 1;
    let leaves = 0;
    let inners = 0;
    const stack: SumTreeNode<T, S>[] = [this.root];
    for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
      if (node.kind === 'leaf') leaves += 1;
      else {
        inners += 1;
        for (const child of node.children) stack.push(child);
      }
    }

    return { depth, leaves, inners, items: this.count };
  }

  /**
   * Every broken invariant, as one line each; empty when the tree is sound. `same` compares two
   * summaries (the tree cannot, as a summary is any type). The invariants:
   *
   * - every node's `sum` is the sum of its children's sums, or of its items' summaries;
   * - every item's `leaf` is the leaf that holds it, and every node's `parent` is its parent;
   * - the leaves form one chain in document order through `prev` and `next`;
   * - no leaf holds more than {@link LEAF_MAX} items and no inner node more than
   *   {@link INNER_MAX} children;
   * - only a root leaf is empty: removal unlinks an emptied leaf and every inner node it empties.
   */
  checkInvariants(same: (a: Readonly<S>, b: Readonly<S>) => boolean): string[] {
    const problems: string[] = [];
    const leaves: SumTreeLeaf<T, S>[] = [];
    let items = 0;
    const visit = (node: SumTreeNode<T, S>, depth: number): number => {
      const sum = this.ops.zero();
      let height = 1;
      if (node.kind === 'leaf') {
        leaves.push(node);
        items += node.items.length;
        if (node.tree !== this) problems.push('a leaf of another tree');
        if (node.items.length > LEAF_MAX) problems.push(`a leaf of ${node.items.length} items`);
        if (node.items.length === 0 && node !== this.root) problems.push('an empty leaf');
        for (const item of node.items) {
          if (item.leaf !== node) problems.push('an item with a wrong leaf');
          this.ops.add(sum, this.ops.of(item));
        }
      } else {
        if (node.children.length === 0) problems.push('an empty inner node');
        if (node.children.length > INNER_MAX)
          problems.push(`an inner node of ${node.children.length} children`);
        const heights = new Set<number>();
        for (const child of node.children) {
          if (child.parent !== node) problems.push('a node with a wrong parent');
          const below = visit(child, depth + 1);
          heights.add(below);
          height = Math.max(height, 1 + below);
          this.ops.add(sum, child.sum);
        }

        if (heights.size > 1) problems.push('leaves at different depths');
      }

      if (!same(sum, node.sum)) problems.push(`a wrong sum at depth ${depth}`);
      return height;
    };
    if (this.root.parent !== null) problems.push('a root with a parent');
    visit(this.root, 0);
    if (items !== this.count) problems.push(`a count of ${this.count} for ${items} items`);
    for (let k = 0; k < leaves.length; k += 1) {
      const leaf = leaves[k];
      if (leaf === undefined) continue;
      if (leaf.prev !== (leaves[k - 1] ?? null)) problems.push(`a wrong prev link at leaf ${k}`);
      if (leaf.next !== (leaves[k + 1] ?? null)) problems.push(`a wrong next link at leaf ${k}`);
    }

    return problems;
  }

  // ---- internals ----

  private leafOf(item: T): SumTreeLeaf<T, S> {
    const leaf = item.leaf;
    if (leaf === undefined || leaf.tree !== this)
      throw new RangeError('The item is not in this tree.');
    return leaf;
  }

  private childAt(node: SumTreeInner<T, S>, at: number): SumTreeNode<T, S> {
    const child = node.children[at];
    if (child === undefined) throw new RangeError('An inner node without children.');
    return child;
  }

  private build(items: readonly T[]): SumTreeNode<T, S> {
    const leaves: SumTreeLeaf<T, S>[] = [];
    for (let at = 0; at < items.length || leaves.length === 0; at += LEAF_FILL) {
      const leaf = new SumTreeLeaf<T, S>(this, this.ops.zero());
      leaf.items = items.slice(at, at + LEAF_FILL);
      for (const item of leaf.items) {
        if (item.leaf !== undefined) throw new RangeError('The item is already in a tree.');
        item.leaf = leaf;
      }

      this.resum(leaf);
      const last = leaves[leaves.length - 1];
      if (last) {
        last.next = leaf;
        leaf.prev = last;
      }

      leaves.push(leaf);
      if (items.length === 0) break;
    }

    this.count = items.length;
    let level: SumTreeNode<T, S>[] = leaves;
    while (level.length > 1) {
      const up: SumTreeNode<T, S>[] = [];
      for (let at = 0; at < level.length; at += INNER_FILL) {
        const inner = new SumTreeInner<T, S>(this.ops.zero());
        inner.children = level.slice(at, at + INNER_FILL);
        for (const child of inner.children) child.parent = inner;
        this.resum(inner);
        up.push(inner);
      }

      level = up;
    }

    const root = level[0];
    if (root === undefined) throw new RangeError('A tree without a root.');
    return root;
  }

  private resum(node: SumTreeNode<T, S>): void {
    const sum = this.ops.zero();
    if (node.kind === 'leaf') for (const item of node.items) this.ops.add(sum, this.ops.of(item));
    else for (const child of node.children) this.ops.add(sum, child.sum);
    node.sum = sum;
  }

  private resumUp(node: SumTreeNode<T, S> | null): void {
    for (let at = node; at; at = at.parent) this.resum(at);
  }

  /** A node that may hold too many entries: split it in even pieces, and its parent after it. */
  private settle(node: SumTreeNode<T, S>): void {
    const max = node.kind === 'leaf' ? LEAF_MAX : INNER_MAX;
    const length = node.kind === 'leaf' ? node.items.length : node.children.length;
    if (length <= max) {
      this.resumUp(node);
      return;
    }

    const fill = node.kind === 'leaf' ? LEAF_FILL : INNER_FILL;
    const pieces = Math.ceil(length / fill);
    const size = Math.ceil(length / pieces);
    const siblings: SumTreeNode<T, S>[] = [];
    if (node.kind === 'leaf') {
      let last = node;
      const after = node.next;
      for (let at = size; at < length; at += size) {
        const leaf = new SumTreeLeaf<T, S>(this, this.ops.zero());
        leaf.items = node.items.slice(at, at + size);
        for (const item of leaf.items) item.leaf = leaf;
        last.next = leaf;
        leaf.prev = last;
        last = leaf;
        siblings.push(leaf);
      }

      last.next = after;
      if (after) after.prev = last;
      node.items = node.items.slice(0, size);
    } else {
      for (let at = size; at < length; at += size) {
        const inner = new SumTreeInner<T, S>(this.ops.zero());
        inner.children = node.children.slice(at, at + size);
        for (const child of inner.children) child.parent = inner;
        siblings.push(inner);
      }

      node.children = node.children.slice(0, size);
    }

    this.resum(node);
    for (const sibling of siblings) this.resum(sibling);
    let parent = node.parent;
    if (parent === null) {
      parent = new SumTreeInner<T, S>(this.ops.zero());
      parent.children = [node];
      node.parent = parent;
      this.root = parent;
    }

    const at = parent.children.indexOf(node) + 1;
    parent.children = [...parent.children.slice(0, at), ...siblings, ...parent.children.slice(at)];
    for (const sibling of siblings) sibling.parent = parent;
    this.settle(parent);
  }
}
