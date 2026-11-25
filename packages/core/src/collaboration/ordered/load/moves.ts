// A file's moves: the `w:moveFrom` and `w:moveTo` wrappers are labels of their suggestions, and
// the range markers around them stay as the file had them (a range can span paragraphs). A
// range marker shows while its move is pending and goes with any decision of it, as the engine
// removes a decided move's markers.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { REVISION_ROLE } from '../model/revision-ids.ts';
import { attributeOf } from '../ooxml-nodes.ts';
import { moveName, pushZero, type Labels, type LoadState, type MoveRange } from './state.ts';

/** A move range marker's local name: its half (`From`, `To`) and its edge (`Start`, `End`). */
export const MOVE_RANGE = /^move(From|To)Range(Start|End)$/;

/** Labels that keep an item while `suggestion` is pending, and hide it once it is decided. */
const whilePending = (suggestion: string, outer: Labels): Labels => ({
  by: suggestion,
  dels: [suggestion],
  move: null,
  within: outer.within,
});

/**
 * Load a move range marker, inline or between blocks. Returns false when the node is not one.
 * `more` gives the item's other fields (a block marker, a field slot).
 */
export function loadMoveRangeMarker(
  state: LoadState,
  node: OoxmlElement,
  labels: Labels,
  more: { readonly block?: boolean; readonly fieldSlot?: { field: string; before: number } } = {}
): boolean {
  const match = MOVE_RANGE.exec(node.localName);
  if (!match) return false;
  const range = match[1] === 'From' ? state.moveFrom : state.moveTo;
  if (match[2] === 'Start') {
    const name = attributeOf(node, 'name');
    if (name !== undefined) {
      range.name = name;
      state.doc.movesWithRanges.add(name);
    }

    range.start = node;

    // Its suggestion is the first wrapper's, read after it
    range.startIndex = pushZero(state, node, labels, { moveRange: true, ...more });
    return true;
  }

  // A range inside another ends before it: its end names its own start
  const suggestion = range.byId.get(attributeOf(node, 'id') ?? '') ?? range.last;
  pushZero(state, node, suggestion === null ? labels : whilePending(suggestion, labels), {
    moveRange: true,
    ...more,
  });

  // The range ended: a later wrapper that no range holds is no part of this move
  range.name = null;
  return true;
}

/** The suggestion of a move wrapper; the range start before it takes the same suggestion. */
export function movedBy(
  state: LoadState,
  node: OoxmlElement,
  range: MoveRange,
  role: string
): string {
  const suggestion = state.revisions.suggestionFor(node, role, moveName(range, node));
  if (range.start !== null) {
    const rangeRole =
      range === state.moveFrom ? REVISION_ROLE.moveFromRange : REVISION_ROLE.moveToRange;
    state.revisions.keepId(suggestion, rangeRole, range.start);
    const id = attributeOf(range.start, 'id');
    if (id !== undefined) {
      range.byId.set(id, suggestion);
    }

    range.start = null;
  }

  if (range.startIndex !== null) {
    const startItem = state.items[range.startIndex];
    if (startItem) {
      state.items[range.startIndex] = { ...startItem, by: suggestion, dels: [suggestion] };
    }

    range.startIndex = null;
  }

  range.last = suggestion;
  return suggestion;
}
