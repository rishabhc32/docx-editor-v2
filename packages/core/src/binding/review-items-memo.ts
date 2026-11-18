// The review queue of a few recent document states, so returning to one does not rebuild it.

import type { OoxmlPart, ReviewItem } from '@docx-editor.dev/core/store';
import type { CustomNodePayloadRead } from '../store/store/custom-node-writes.ts';

/** Everything the queue derivation reads, by identity. */
export interface ReviewItemsInputs {
  readonly root: object;
  readonly commentsPart: OoxmlPart | undefined;
  readonly commentsExtendedPart: OoxmlPart | undefined;
  readonly stylesPart: OoxmlPart | undefined;
  readonly furnitureParts: readonly OoxmlPart[];
  readonly customNodePayloads: ReadonlyMap<string, CustomNodePayloadRead>;
}

/** States kept, at most. Undo walks back through the states an edit just left. */
const CAPACITY = 8;

/** Payload reads are rebuilt per call, so their fields are compared, not the objects. */
function samePayloads(
  left: ReadonlyMap<string, CustomNodePayloadRead>,
  right: ReadonlyMap<string, CustomNodePayloadRead>
): boolean {
  if (left.size !== right.size) return false;
  for (const [key, value] of left) {
    const other = right.get(key);
    if (
      !other ||
      other.nodeId !== value.nodeId ||
      other.label !== value.label ||
      other.data !== value.data
    )
      return false;
  }
  return true;
}

function sameInputs(left: ReviewItemsInputs, right: ReviewItemsInputs): boolean {
  return (
    left.root === right.root &&
    left.commentsPart === right.commentsPart &&
    left.commentsExtendedPart === right.commentsExtendedPart &&
    left.stylesPart === right.stylesPart &&
    left.furnitureParts.length === right.furnitureParts.length &&
    left.furnitureParts.every((part, index) => part === right.furnitureParts[index]) &&
    samePayloads(left.customNodePayloads, right.customNodePayloads)
  );
}

/**
 * Undo restores an earlier tree by reference, but the package around it is new, so the
 * revision-keyed cache misses and the whole queue was derived again: a whole-document walk
 * per undo step on a document with thousands of changes. The derivation is a function of
 * these inputs alone, so equal inputs give back the queue derived for them before.
 *
 * An entry lives only while something else holds its tree (the document, the undo history):
 * a tree nothing holds cannot come back, so its queue is dead weight. In a room every remote
 * change drops the undo history, and the memo was then the only holder of its older trees and
 * of their whole queues.
 */
export function createReviewItemsMemo(): {
  get(inputs: ReviewItemsInputs): readonly ReviewItem[] | undefined;
  set(inputs: ReviewItemsInputs, items: readonly ReviewItem[]): void;
} {
  // Keyed weakly on the tree, so an entry goes with its tree. The entry's own reference to
  // the tree does not keep it: a WeakMap value does not hold its key.
  const byRoot = new WeakMap<
    object,
    { readonly inputs: ReviewItemsInputs; readonly items: readonly ReviewItem[] }
  >();
  // Most recent first, for the CAPACITY bound alone, so it holds the trees weakly too.
  let recent: WeakRef<object>[] = [];
  const touch = (root: object): void => {
    const next = [new WeakRef(root)];
    for (const ref of recent) {
      const other = ref.deref();
      if (other === undefined || other === root) continue;
      if (next.length < CAPACITY) next.push(ref);
      else byRoot.delete(other);
    }
    recent = next;
  };
  return {
    get(inputs) {
      const entry = byRoot.get(inputs.root);
      if (!entry || !sameInputs(entry.inputs, inputs)) return undefined;
      touch(inputs.root);
      return entry.items;
    },
    set(inputs, items) {
      byRoot.set(inputs.root, { inputs, items });
      touch(inputs.root);
    },
  };
}
