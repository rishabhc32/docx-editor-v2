import { describe, expect, test } from 'bun:test';
import { INNER_MAX, LEAF_MAX, SumTree, type SumTreeItem } from '../tree/sum-tree.ts';
import { seededRandom } from './random.ts';

interface Summary {
  count: number;
  weight: number;
}

interface Box extends SumTreeItem<Box, Summary> {
  readonly id: number;
  weight: number;
}

const OPS = {
  zero: (): Summary => ({ count: 0, weight: 0 }),
  add: (into: Summary, other: Readonly<Summary>) => {
    into.count += other.count;
    into.weight += other.weight;
  },
  of: (box: Box): Summary => ({ count: 1, weight: box.weight }),
};

const same = (a: Readonly<Summary>, b: Readonly<Summary>) =>
  a.count === b.count && a.weight === b.weight;

let nextId = 0;
const box = (weight: number): Box => ({ id: (nextId += 1), weight, leaf: undefined });
const ids = (boxes: Iterable<Box>) => [...boxes].map((b) => b.id);

/** The first box at which the running weight passes `target`, by a plain scan. */
function scanSeek(boxes: readonly Box[], target: number): { box: Box; before: number } | null {
  let total = 0;
  for (const b of boxes) {
    if (total + b.weight > target) return { box: b, before: total };
    total += b.weight;
  }
  return null;
}

function expectSound(tree: SumTree<Box, Summary>): void {
  expect(tree.checkInvariants(same)).toEqual([]);
}

describe('SumTree', () => {
  test('an empty tree has one empty leaf and no items', () => {
    const tree = new SumTree<Box, Summary>(OPS);
    expect(tree.first()).toBeNull();
    expect(tree.last()).toBeNull();
    expect(tree.size).toBe(0);
    expect(tree.summary).toEqual({ count: 0, weight: 0 });
    expect(tree.seek((s) => s.weight, 0)).toBeNull();
    expectSound(tree);
  });

  test('a bulk build keeps order, sums and links', () => {
    const boxes = Array.from({ length: 5000 }, (_, k) => box(k % 7));
    const tree = new SumTree<Box, Summary>(OPS, boxes);
    expect(ids(tree.items())).toEqual(ids(boxes));
    expect(ids(tree.itemsBackward())).toEqual(ids(boxes).reverse());
    expect(tree.summary.count).toBe(5000);
    expect(tree.stats().depth).toBeGreaterThan(2);
    expectSound(tree);
  });

  test('a large insertion in one leaf splits it into leaves that all fit', () => {
    const tree = new SumTree<Box, Summary>(OPS, [box(1), box(1)]);
    const many = Array.from({ length: LEAF_MAX * INNER_MAX * 3 }, () => box(2));
    tree.insertAfter(tree.first(), many);
    expect(tree.size).toBe(many.length + 2);
    expectSound(tree);
  });

  test('an item can be in one tree only, and a foreign item is refused', () => {
    const a = new SumTree<Box, Summary>(OPS, [box(1)]);
    const b = new SumTree<Box, Summary>(OPS, [box(1)]);
    const first = a.first();
    expect(first).not.toBeNull();
    if (first === null) return;
    expect(() => b.next(first)).toThrow(RangeError);
    expect(() => b.insertAfter(null, [first])).toThrow(RangeError);
    expect(b.has(first)).toBe(false);
    expect(a.has(first)).toBe(true);
  });

  test('removing every item leaves a sound empty tree', () => {
    const boxes = Array.from({ length: 700 }, () => box(1));
    const tree = new SumTree<Box, Summary>(OPS, boxes);
    for (const b of boxes) tree.remove(b);
    expect(tree.size).toBe(0);
    expect(tree.first()).toBeNull();
    expectSound(tree);
    tree.insertAfter(null, [box(3)]);
    expect(tree.summary).toEqual({ count: 1, weight: 3 });
    expectSound(tree);
  });

  test('random edits agree with a plain array (seed 20261028)', () => {
    const random = seededRandom(20261028);
    const int = (n: number) => Math.floor(random() * n);
    const model: Box[] = Array.from({ length: 200 }, () => box(int(5)));
    const tree = new SumTree<Box, Summary>(OPS, model);
    for (let step = 0; step < 4000; step += 1) {
      const choice = random();
      if (choice < 0.35 || model.length === 0) {
        // Insert a batch, sometimes a large one, after a random item or at the start
        const batch = Array.from({ length: random() < 0.05 ? 300 : 1 + int(8) }, () => box(int(5)));
        const at = int(model.length + 1);
        tree.insertAfter(at === 0 ? null : (model[at - 1] ?? null), batch);
        model.splice(at, 0, ...batch);
      } else if (choice < 0.55) {
        // Split one item in pieces
        const at = int(model.length);
        const target = model[at];
        if (target === undefined) continue;
        const pieces = Array.from({ length: 1 + int(3) }, () => box(int(5)));
        tree.replace(target, pieces);
        model.splice(at, 1, ...pieces);
      } else if (choice < 0.8) {
        const at = int(model.length);
        const target = model[at];
        if (target === undefined) continue;
        tree.remove(target);
        model.splice(at, 1);
      } else {
        const target = model[int(model.length)];
        if (target === undefined) continue;
        target.weight = int(5);
        tree.update(target);
      }
      if (step % 50 !== 0) continue;
      expect(ids(tree.items())).toEqual(ids(model));
      expect(tree.size).toBe(model.length);
      const weight = model.reduce((n, b) => n + b.weight, 0);
      expect(tree.summary).toEqual({ count: model.length, weight });
      for (let probe = 0; probe < 20; probe += 1) {
        const target = int(weight + 2);
        const want = scanSeek(model, target);
        const got = tree.seek((s) => s.weight, target);
        expect(got === null ? null : { id: got.item.id, before: got.before.weight }).toEqual(
          want === null ? null : { id: want.box.id, before: want.before }
        );
      }
      const at = int(model.length);
      const picked = model[at];
      if (picked) {
        const before = model.slice(0, at).reduce((n, b) => n + b.weight, 0);
        expect(tree.before(picked)).toEqual({ count: at, weight: before });
        expect(tree.prev(picked)?.id ?? null).toBe(model[at - 1]?.id ?? null);
        expect(tree.next(picked)?.id ?? null).toBe(model[at + 1]?.id ?? null);
      }
      expectSound(tree);
    }
    expect(ids(tree.items())).toEqual(ids(model));
    expectSound(tree);
  });
});
