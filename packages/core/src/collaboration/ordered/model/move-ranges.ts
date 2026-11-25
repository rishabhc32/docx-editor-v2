// A pending move's range markers. Each half of every move has its words between a range start
// and a range end item: the file's own markers, markers a load gives a file's move without them,
// or markers the move's change made. Compose draws them all as a file's, so an edit at the edge of
// moved words goes inside or outside the range the same way live, after a fold and in the editor
// alone.

import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { wmlAttribute } from '../ooxml-nodes.ts';
import { REVISION_ROLE, revisionMarkOf } from './revision-ids.ts';
import type { OrderedDocument } from './types.ts';
import { recordUndo } from './undo-record.ts';

/** A move's half, as its markers name it (`w:moveFromRangeStart`, `w:moveToRangeEnd`). */
export type MoveHalf = 'From' | 'To';

/** Labels that keep a range marker while its move is pending, and hide it once it is decided. */
export const whilePending = (suggestion: string) =>
  ({ by: suggestion, dels: [suggestion], move: null }) as const;

/**
 * A move half's range start and end, as Word writes them: the start with the move's `w:id`, name
 * and mark, the end with the same `w:id`. `nodeIds` gives the two elements their node ids. The
 * document notes the move as one with range markers (an edit's undo record takes that back).
 */
export function moveRangeNodes(
  doc: OrderedDocument,
  half: MoveHalf,
  name: string,
  suggestion: string,
  nodeIds: readonly [string, string]
): readonly [OoxmlElement, OoxmlElement] {
  const role = half === 'From' ? REVISION_ROLE.moveFromRange : REVISION_ROLE.moveToRange;

  // `revisionMarkOf` puts the `w:id` first: the check is for the type only
  const [id, ...mark] = revisionMarkOf(doc, suggestion, role);
  const ids = id === undefined ? [] : [id];
  if (!doc.movesWithRanges.has(name)) {
    doc.movesWithRanges.add(name);
    recordUndo(doc, () => doc.movesWithRanges.delete(name));
  }

  const start = `move${half}RangeStart` as const;
  const end = `move${half}RangeEnd` as const;
  return [
    build(nodeIds[0], start, start, [...ids, wmlAttribute('name', name), ...mark], []),
    build(nodeIds[1], end, end, ids, []),
  ];
}
