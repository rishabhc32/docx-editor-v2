// Take back a table formatting suggestion that another author's change to the same row beat.
//
// Suggesting refuses a table formatting change to a row that holds another author's pending
// change, because Word answers a row's records as one decision. In a shared room two people
// can still each change one row before either sees the other's record. Every replica gives the
// row the same owner (`compareRowRecordOwners`): the record with the earlier `w:date`, then the
// lesser author, then the lesser revision id.
//
// After a remote update, the editor whose latest table formatting gesture lost takes it back
// with collaborative undo, which reverts this author's changes only. It does so while the
// gesture is still the last step of the local history, so it never undoes a later edit; after
// one, the room keeps showing the row as its owner's one decision. The undone step leaves the
// redo history too, so Redo cannot bring the losing change back, and the user reads why it went.

import type { TreeDocxSessionView } from '@docx-editor.dev/core/binding';
import {
  findNode,
  ORIGIN_IDS,
  type OoxmlElement,
  type TreeModelChange,
} from '@docx-editor.dev/core/store';
import type { EditorCollaborationSession } from '../collaboration/index.ts';
import { tableChildren } from '../store/store/revision-table-children.ts';
import {
  compareRowRecordOwners,
  recordsTableFormatting,
  rowPropertyRecords,
  rowRecordOwnerOf,
  type RowRecordOwner,
} from '../store/store/table-row-owner.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import { undoHistoryPosition } from '../store/store/undo-history-position.ts';
import { partOfNodeId } from './surface-scope.ts';

/** Keyed in `@docx-editor.dev/i18n`'s refusal table; the two must stay the same sentence. */
export const ROW_TAKEN_REASON =
  'another person suggested a change to this row first; your table formatting was removed';

/** The latest local table formatting gesture, while it can still be taken back. */
interface Gesture {
  readonly author: string;
  /** The last step of the local history right after the gesture: the gesture's own step. */
  readonly anchor: object;
  /** Each row the gesture recorded, with this author's earliest record in it. */
  readonly rows: ReadonlyMap<string, RowRecordOwner>;
}

/** The session reads the watch takes, and the commits it hears of. */
type Session = Pick<TreeDocxSessionView, 'currentPackage' | 'part'> & {
  subscribe(onChange: (change: TreeModelChange) => void): () => void;
};

/** The row as the session holds it now, or undefined once it is gone. */
function rowAt(session: Session, rowId: string): OoxmlElement | undefined {
  const row = findNode(partOfNodeId(session, rowId) ?? session.part(), rowId);
  return row?.kind === 'tableRow' ? row : undefined;
}

/** This author's earliest record in the row, or undefined when it holds none. */
function ownRecord(row: OoxmlElement, author: string): RowRecordOwner | undefined {
  let earliest: RowRecordOwner | undefined;
  for (const record of rowPropertyRecords(row)) {
    const owner = rowRecordOwnerOf(record.attributes);
    if (owner.author !== author) continue;
    if (!earliest || compareRowRecordOwners(owner, earliest) < 0) earliest = owner;
  }
  return earliest;
}

/** Whether a row of the gesture now shows a record of another author that owns the row. */
function lostRow(session: Session, gesture: Gesture): boolean {
  for (const [rowId, own] of gesture.rows) {
    const row = rowAt(session, rowId);
    if (!row) continue;
    for (const record of rowPropertyRecords(row)) {
      const owner = rowRecordOwnerOf(record.attributes);
      if (owner.author !== gesture.author && compareRowRecordOwners(owner, own) < 0) return true;
    }
  }
  return false;
}

/**
 * Watch the local table formatting gestures of an editor in a shared room. `watch` is called
 * with a gesture's ops before it commits, and its result after the gesture landed.
 */
export function createTableRowRace(
  session: Session,
  collaboration: EditorCollaborationSession | null | undefined,
  publishRefusal: (reason: string) => void
): {
  readonly watch: (ops: readonly TreeDocOp[]) => () => void;
  readonly destroy: () => void;
} {
  const position = collaboration ? undoHistoryPosition(collaboration) : undefined;
  let gesture: Gesture | null = null;
  let checking = false;

  function check(): void {
    checking = false;
    const current = gesture;
    if (!current || !collaboration || !position) return;
    if (position.current() !== current.anchor) {
      gesture = null;
      return;
    }
    if (!lostRow(session, current)) return;
    gesture = null;
    if (!collaboration.undo()) return;
    position.split();
    publishRefusal(ROW_TAKEN_REASON);
  }

  // A remote update can change who owns a row of the gesture; any local commit, undo or redo
  // takes the gesture off the top of the local history, so it can no longer be taken back. The
  // check runs once the update is installed: an undo inside the install would not reach this
  // editor until the next one.
  const unsubscribe = position
    ? session.subscribe((change) => {
        if (gesture === null) return;
        if (change.origin !== ORIGIN_IDS.mutationRemote) gesture = null;
        else if (!checking) {
          checking = true;
          queueMicrotask(check);
        }
      })
    : () => {};

  return {
    watch(ops) {
      const formatting = ops.filter(recordsTableFormatting);
      if (!position || formatting.length === 0) return () => {};
      const author = formatting[0]!.revision.author;
      const before = new Map(
        formatting.map((op) => [op.tableId, partOfNodeId(session, op.tableId) ?? session.part()])
      );
      return () => {
        const rows = new Map<string, RowRecordOwner>();
        for (const [tableId, part] of before) {
          const table = findNode(partOfNodeId(session, tableId) ?? session.part(), tableId);
          if (table?.kind !== 'table') continue;
          for (const row of tableChildren(table, 'tableRow')) {
            if (findNode(part, row.id) === row) continue;
            const own = ownRecord(row, author);
            if (own) rows.set(row.id, own);
          }
        }
        const anchor = position.current();
        if (rows.size > 0 && anchor) gesture = { author, anchor, rows };
      };
    },
    destroy() {
      unsubscribe();
      gesture = null;
    },
  };
}
