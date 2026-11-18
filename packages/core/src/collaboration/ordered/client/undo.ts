// Undo and redo of the person's own steps, as the session runs them (tasks 6.3, 6.7 and 9.12).

import { isOrderedEditChange, type OrderedChange } from '../change/types.ts';
import type { OrderedDocument } from '../model/types.ts';
import type { OrderedHistory, OrderedHistoryCommit, OrderedHistoryStep } from './history.ts';
import { inverseOfDirect, madeBy } from './undo-direct.ts';
import type { OrderedDirty } from './publish.ts';
import { OrderedClientRefusal, type OrderedClientRefusalReport } from './refusal.ts';
import type { OrderedReplica } from './replica.ts';

/** What undo and redo use of the session. */
export interface OrderedHistoryWriter {
  readonly history: OrderedHistory;
  readonly replica: OrderedReplica;
  /** The change as the room's validator reads it, or a refusal. */
  validate(candidate: Record<string, unknown>): OrderedChange;
  /** The person's own pending suggestions among these. */
  ownPending(ids: readonly string[]): string[];
  /** What a write changed, for the next publish. */
  written(dirty: OrderedDirty): void;
  refused(report: OrderedClientRefusalReport): void;
}

/**
 * Undo the person's last step. False when there is none or it has no inverse. The step's own
 * pending suggestions are taken back in one `retract`; a step whose suggestions someone decided
 * has no inverse and is refused by name. A step of direct edits (the organiser's) is taken back
 * by its inverse, read from the list (`undo-direct.ts`).
 */
export function undoOwnStep(writer: OrderedHistoryWriter): boolean {
  const { history, replica } = writer;
  const step = history.popUndo();
  if (!step) {
    return false;
  }

  if (step.direct) {
    const made = applyInverse(writer, step, (back) => history.restoreUndo(back));
    if (made === null) {
      return false;
    }

    // A step of tracked and direct changes is made again as it was, under new ids
    const mixed = step.changes.some((c) => isOrderedEditChange(c) && c.mode === 'suggest');

    // A step nothing of which shows any more is spent: there is nothing to redo either
    if (made.length > 0) {
      history.pushRedo(
        mixed
          ? { ...step, remake: true }
          : {
              ...step,
              suggestions: new Set(),
              changes: made,
              commits: [{ start: 0, selection: step.selection }],
            }
      );
    }

    return true;
  }

  const commit = ownEditCommit(replica.doc, step);
  if (commit) {
    return undoOwnEditCommit(writer, step, commit);
  }

  const own = writer.ownPending([...step.suggestions]);
  if (own.length === 0) {
    return refuseStep(writer, 'no-inverse', 'decided');
  }

  if (withdrewAll(writer, step, own)) {
    return remakeEarlier(writer, step);
  }

  try {
    const change = writer.validate({
      v: 1,
      id: replica.nextId(),
      kind: 'retract',
      suggestions: own,
    });
    writer.written(replica.applyLocal(change).dirty);
  } catch (error) {
    return restoreAndRefuse(writer, step, error);
  }

  history.pushRedo(step);
  return true;
}

/** Report why an undo or a redo did not run. Always false. */
function refuseStep(writer: OrderedHistoryWriter, reason: string, detail?: string): false {
  writer.refused({
    code: 'ordered-refused',
    reason,
    ...(detail === undefined ? {} : { detail }),
  });
  return false;
}

/**
 * Put an undo's step back on the undo stack and report the refusal. Any other error is a defect,
 * thrown again after the step is back.
 */
function restoreAndRefuse(
  writer: OrderedHistoryWriter,
  step: OrderedHistoryStep,
  error: unknown
): false {
  writer.history.restoreUndo(step);
  if (!(error instanceof OrderedClientRefusal)) {
    throw error;
  }

  return refuseStep(writer, error.reason, error.detail);
}

/** The kinds whose change takes back the author's own pending items it reaches. */
const WITHDRAWS: ReadonlySet<string> = new Set(['delete', 'delete-column', 'delete-row', 'join']);

/**
 * The kinds whose edits of the author's own pending words the model takes back exactly (`undo`
 * of a `retract`): what they took back, added (typed text, a tab, a break, Enter, a note), or
 * formatted there.
 */
const OWN_EDITS: ReadonlySet<string> = new Set([
  ...WITHDRAWS,
  'insert',
  'insert-tab',
  'insert-break',
  'split',
  'insert-note',
  'format',
  'mark-format',
]);

/** Whether a change edited its author's earlier pending words (`ownEdited`, `added`). */
const editedOwn = (doc: OrderedDocument, change: OrderedChange): boolean =>
  doc.ownEdited.has(change.id) || doc.added.has(change.id);

/**
 * The step's last commit, when it edited the person's own earlier pending words (took part of
 * them back, added to them, formatted them) and the model can take that back exactly:
 * its undo takes back that commit alone, as the editor undoes it, and the step's earlier
 * commits stay to undo next.
 */
function ownEditCommit(
  doc: OrderedDocument,
  step: OrderedHistoryStep
): OrderedHistoryCommit | null {
  const last = step.commits.at(-1);
  if (!last) return null;
  const changes = step.changes.slice(last.start);
  if (!changes.every((change) => OWN_EDITS.has(change.kind))) return null;
  return changes.some((change) => editedOwn(doc, change)) ? last : null;
}

/**
 * Undo of a commit that edited the person's own pending words: one `retract` of the suggestions
 * it made (another person's words it struck, formatted) that also takes back exactly what it did
 * to the person's earlier suggestions (`undo`). The step's earlier commits go back on the undo
 * stack, and the commit waits on the redo stack.
 */
function undoOwnEditCommit(
  writer: OrderedHistoryWriter,
  step: OrderedHistoryStep,
  commit: OrderedHistoryCommit
): boolean {
  const { history, replica } = writer;
  const changes = step.changes.slice(commit.start);
  const made = writer.ownPending(
    changes.flatMap((change) => madeBy(replica.doc, change.id, false))
  );
  try {
    const change = writer.validate({
      v: 1,
      id: replica.nextId(),
      kind: 'retract',
      suggestions: made,
      undo: changes.filter((c) => editedOwn(replica.doc, c)).map((c) => c.id),
    });
    writer.written(replica.applyLocal(change).dirty);
  } catch (error) {
    return restoreAndRefuse(writer, step, error);
  }

  if (commit.start > 0) {
    history.restoreUndo({
      ...step,
      suggestions: new Set([...step.suggestions].filter((id) => !made.includes(id))),
      changes: step.changes.slice(0, commit.start),
      commits: step.commits.slice(0, -1),
    });
  }

  history.pushRedo({
    suggestions: new Set(made),
    changes,
    commits: [{ start: 0, selection: commit.selection }],
    direct: false,
    selection: commit.selection,
  });
  history.restoreSelection(commit.selection);
  return true;
}

/**
 * Whether the step's last change, a deletion or a reject's `retract`, took back all the step made
 * and left nothing to show again: none of its own suggestions labels an item that is not taken
 * back. The deletion of one's own new column withdraws its cells, as the engine does, and takes
 * the column's suggestion back whole with its lent widths, so neither a `retract` nor its `undo`
 * would change anything. The reject of one's own change inside another person's suggestion is
 * such a `retract`: its undo gives the change back, as the editor's undo of that reject does.
 */
function withdrewAll(
  writer: OrderedHistoryWriter,
  step: OrderedHistoryStep,
  own: readonly string[]
): boolean {
  const last = step.changes[step.changes.length - 1];
  if (step.changes.length < 2 || !last) {
    return false;
  }

  if (!WITHDRAWS.has(last.kind) && last.kind !== 'retract') {
    return false;
  }

  const { labelled } = writer.replica.doc;
  return own.every((id) => {
    for (const item of labelled.get(id) ?? []) {
      if (!item.retracted) {
        return false;
      }
    }

    return true;
  });
}

/**
 * Undo of a deletion that took back all its step made: the step's earlier changes made again
 * under new ids, as the editor's undo gives back what the deletion took away. They are one step
 * on the undo stack, and the deletion waits on the redo stack, renamed to what they make.
 */
function remakeEarlier(writer: OrderedHistoryWriter, step: OrderedHistoryStep): boolean {
  const { history } = writer;
  let made: Remade[];
  try {
    made = remade(writer, step.changes.slice(0, -1));
  } catch (error) {
    return restoreAndRefuse(writer, step, error);
  }

  recordRemade(history, made, step);
  history.pushRedo({
    suggestions: new Set(),
    changes: step.changes.slice(-1),
    commits: [{ start: 0, selection: step.selection }],
    direct: step.direct,
    selection: step.selection,
  });
  return true;
}

/** A change made again, and the person's own pending suggestions it made. */
interface Remade {
  readonly change: OrderedChange;
  readonly own: string[];
}

/**
 * Make changes again under new ids, as the person's own. On any error the changes made so far
 * are taken back and the error is thrown again.
 */
function remade(writer: OrderedHistoryWriter, changes: readonly OrderedChange[]): Remade[] {
  const { history, replica } = writer;
  const made: Remade[] = [];
  try {
    for (const original of changes) {
      const change = writer.validate(history.again(original, replica.nextId(), replica.doc));
      const { outcome, dirty } = replica.applyLocal(change);
      writer.written(dirty);
      made.push({ change, own: writer.ownPending(outcome.suggestions) });
    }
  } catch (error) {
    writer.written(replica.dropLast(made.length));
    throw error;
  }

  return made;
}

/** The changes made again as one step of the history, with the step's selection. */
function recordRemade(
  history: OrderedHistory,
  made: readonly Remade[],
  step: OrderedHistoryStep
): void {
  made.forEach(({ change, own }, k) =>
    history.record(change, own, step.direct, step.selection, {
      fromRedo: true,
      sameCommit: k > 0,
    })
  );
}

/**
 * Apply the inverse of a direct step as the person's own changes, or report why it has none
 * and put the step back. Returns the changes it made.
 */
function applyInverse(
  writer: OrderedHistoryWriter,
  step: OrderedHistoryStep,
  putBack: (step: OrderedHistoryStep) => void
): OrderedChange[] | null {
  const { replica } = writer;
  const made: OrderedChange[] = [];
  try {
    for (const body of inverseOfDirect(replica.doc, step.changes)) {
      const change = writer.validate({ v: 1, id: replica.nextId(), ...body });
      writer.written(replica.applyLocal(change).dirty);
      made.push(change);
    }
  } catch (error) {
    writer.written(replica.dropLast(made.length));
    putBack(step);
    if (!(error instanceof OrderedClientRefusal)) {
      throw error;
    }

    refuseStep(writer, error.reason, error.detail);
    return null;
  }

  return made;
}

/**
 * Redo the person's last undone step. False when there is none or it no longer applies. The
 * step's changes are made again under new ids; a change that no longer applies refuses the whole
 * redo and takes back what it made.
 */
export function redoOwnStep(writer: OrderedHistoryWriter): boolean {
  const { history } = writer;
  const step = history.popRedo();
  if (!step) {
    return false;
  }

  if (step.direct && step.remake !== true) {
    const again = applyInverse(writer, step, (back) => history.pushRedo(back));
    if (again === null) {
      return false;
    }

    if (again.length > 0) {
      history.recordStep(again, step.selection);
    }

    return true;
  }

  let made: Remade[];
  try {
    made = remade(writer, step.changes);
  } catch (error) {
    if (!(error instanceof OrderedClientRefusal)) {
      throw error;
    }

    return refuseStep(writer, 'redo-refused', error.reason);
  }

  recordRemade(history, made, step);
  return true;
}
