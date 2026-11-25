import type { CollaborationLocalSelection } from '../../index.ts';
import type { HistoryGroup } from '../../../store/store/history-group.ts';
import { isOrderedCommentChange, type OrderedChange } from '../change/types.ts';
import { orderedListNumber } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';

/** One editor commit of a step: where its changes start, and the selection before it. */
export interface OrderedHistoryCommit {
  readonly start: number;
  readonly selection: CollaborationLocalSelection | null;
}

/** One undo step. */
export interface OrderedHistoryStep {
  /** The person's suggestions the step made or extended. */
  readonly suggestions: Set<string>;
  /** The changes, in the order they were made. */
  readonly changes: OrderedChange[];
  /**
   * The editor commits the changes came in, in order. The undo of a step whose last commit took
   * back the person's own items takes back that commit alone (`undo.ts`).
   */
  readonly commits: OrderedHistoryCommit[];
  /** The step holds direct edits (the organiser's): its inverse is read from the list. */
  direct: boolean;
  /** The selection before the step's first change. */
  readonly selection: CollaborationLocalSelection | null;
  /**
   * An undone step of tracked and direct changes: redo makes its changes again under new ids,
   * as a step of suggestions is redone, since a take-back has no inverse to read.
   */
  readonly remake?: boolean;
  /** The editor's undo group of the step's last commit: a later commit of it joins the step. */
  group?: HistoryGroup;
}

/** Most steps kept; the oldest go first. */
const MAX_STEPS = 500;

/**
 * A space or suggestion a change minted: units, marks, containers, markers, resources,
 * suggestions, and node ids (a new cell's, a new row's), also as the token space of a container
 * that holds one (`t:<file>:n:<change>.<k>`).
 */
const MINTED = /^((?:t:\d{1,9}:)?)([ikcxpSn]):(.+)\.(\d{1,9})$/;

/**
 * What a change minted under an old id, under the new id; any other string as it is. A redo
 * renames so the later changes of a step name the units its earlier changes make again.
 */
function renamedId(value: string, ids: ReadonlyMap<string, string>): string {
  const match = MINTED.exec(value);
  const next = match?.[3] === undefined ? undefined : ids.get(match[3]);
  return match && next !== undefined ? `${match[1]}${match[2]}:${next}.${match[4]}` : value;
}

/**
 * The numbers of the lists a renamed change made, to the numbers they take again: a list's
 * number is read from its key, so it follows the key (a long paste's later changes name it).
 */
function renamedLists(doc: OrderedDocument, ids: ReadonlyMap<string, string>): Map<string, string> {
  const out = new Map<string, string>();
  for (const resource of doc.resources.added.values()) {
    if (resource.kind !== 'list') {
      continue;
    }

    const key = renamedId(resource.key, ids);
    if (key !== resource.key) {
      out.set(resource.numId, orderedListNumber(key, 'num'));
    }
  }

  return out;
}

/** Rename what a change minted under an old id to a new id, in every string of a value. */
function renamed(
  value: unknown,
  ids: ReadonlyMap<string, string>,
  lists: ReadonlyMap<string, string>
): unknown {
  if (typeof value === 'string') {
    return lists.get(value) ?? renamedId(value, ids);
  }

  if (Array.isArray(value)) {
    return value.map((entry) => renamed(entry, ids, lists));
  }

  if (typeof value === 'object' && value !== null) {
    const out: Record<string, unknown> = {};
    for (const [key, entry] of Object.entries(value)) {
      out[key] = renamed(entry, ids, lists);
    }

    return out;
  }

  return value;
}

/**
 * A new change's step: the older steps that share a suggestion with it, in order, and its own
 * suggestions. The selection is the first older step's that has one.
 */
function mergedStep(
  older: readonly OrderedHistoryStep[],
  suggestions: readonly string[],
  direct: boolean,
  selection: CollaborationLocalSelection | null
): OrderedHistoryStep {
  const merged: OrderedHistoryStep = {
    suggestions: new Set(suggestions),
    changes: [],
    commits: [],
    direct,
    selection: older.find((step) => step.selection !== null)?.selection ?? selection,
  };
  for (const step of older) {
    for (const id of step.suggestions) {
      merged.suggestions.add(id);
    }

    const offset = merged.changes.length;
    merged.changes.push(...step.changes);
    merged.commits.push(...step.commits.map((c) => ({ ...c, start: c.start + offset })));
    merged.direct ||= step.direct;
  }

  return merged;
}

/**
 * Undo and redo of the person's own changes (design D7). A step is one or more of the person's
 * suggestions and the changes that made them. Undo takes a step's suggestions back with a
 * `retract`, so it never touches another person's text. Direct steps undo by their inverse, read
 * from the list (`undo-direct.ts`). The editor's undo group decides which commits are one step, a
 * typing run among them (`editor/typing-history.ts`), for direct changes as for suggestions.
 */
/** How `OrderedHistory.record` takes a change. */
export interface OrderedRecordOptions {
  /** A redo made the change: the redo stack stays. */
  readonly fromRedo?: boolean;
  /** The change is a later change of the step's last editor commit. */
  readonly sameCommit?: boolean;
}

export class OrderedHistory {
  readonly #undo: OrderedHistoryStep[] = [];
  readonly #redo: OrderedHistoryStep[] = [];
  /** Old change id to the id its redo took: later steps name what the earlier redo minted. */
  readonly #ids = new Map<string, string>();
  #restored: CollaborationLocalSelection | null = null;

  get canUndo(): boolean {
    return this.#undo.length > 0;
  }

  get canRedo(): boolean {
    return this.#redo.length > 0;
  }

  /** The selection the last undo restored. */
  get restored(): CollaborationLocalSelection | null {
    return this.#restored;
  }

  /**
   * A change of the person's own, and the suggestions it made or extended. A step that already
   * holds one of them takes it and goes to the top; a new change clears redo.
   */
  record(
    change: OrderedChange,
    suggestions: readonly string[],
    direct: boolean,
    selection: CollaborationLocalSelection | null,
    { fromRedo = false, sameCommit = false }: OrderedRecordOptions = {}
  ): void {
    if (!fromRedo) {
      this.#redo.length = 0;
    }

    const top = this.#undo[this.#undo.length - 1];

    // The changes of one editor commit are one step, as the editor undoes them (a table's rows,
    // the lines of a paste), even when each made a suggestion of its own (direct edits)
    if (sameCommit && top) {
      top.changes.push(change);
      for (const id of suggestions) {
        top.suggestions.add(id);
      }

      top.direct ||= direct;
      return;
    }

    const merged = mergedStep(this.#takeSteps(suggestions), suggestions, direct, selection);
    merged.commits.push({ start: merged.changes.length, selection });
    merged.changes.push(change);
    this.#undo.push(merged);
    if (this.#undo.length > MAX_STEPS) {
      this.#undo.shift();
    }
  }

  /** Take off the undo stack the steps that hold one of these suggestions, oldest first. */
  #takeSteps(suggestions: readonly string[]): OrderedHistoryStep[] {
    const taken: OrderedHistoryStep[] = [];
    for (let k = this.#undo.length - 1; k >= 0; k -= 1) {
      const step = this.#undo[k];
      if (!step || !suggestions.some((id) => step.suggestions.has(id))) {
        continue;
      }

      this.#undo.splice(k, 1);
      taken.unshift(step);
    }

    return taken;
  }

  /**
   * The changes of one editor commit, as the editor's undo takes it: one step. Each direct
   * change makes a suggestion of its own, so a direct commit of several changes (an
   * `authorTable` with values, Enter that names the tail's style) is one new step. A commit
   * that also made suggestions (a paste's lines, a tracked insertion with a direct format)
   * joins its changes to its first one's step, with their suggestions (`record`).
   */
  recordCommit(
    made: readonly { readonly change: OrderedChange; readonly own: readonly string[] }[],
    direct: boolean,
    selection: CollaborationLocalSelection | null,
    group?: HistoryGroup
  ): void {
    // Comments stay out of undo, as Word and Google Docs keep them
    const entries = made.filter(({ change }) => !isOrderedCommentChange(change));
    if (entries.length === 0) {
      return;
    }

    const top = this.#undo[this.#undo.length - 1];

    // A later commit of the same editor command (a table of contents' page pass after its
    // rows) joins the command's step, as the editor's own history joins it
    const joins = group !== undefined && top?.group === group;
    if (!joins && direct && entries.length > 1 && entries.every(({ own }) => own.length === 0)) {
      this.#redo.length = 0;
      this.recordStep(
        entries.map(({ change }) => change),
        selection
      );
    } else {
      entries.forEach(({ change, own }, k) =>
        this.record(change, own, direct, selection, { sameCommit: joins || k > 0 })
      );
    }

    const step = this.#undo[this.#undo.length - 1];
    if (step && entries.length > 0) step.group = group;
  }

  /** A whole step of direct changes: a direct commit of several changes, or a redo. */
  recordStep(
    changes: readonly OrderedChange[],
    selection: CollaborationLocalSelection | null
  ): void {
    this.#undo.push({
      suggestions: new Set(),
      changes: [...changes],
      commits: [{ start: 0, selection }],
      direct: true,
      selection,
    });
    if (this.#undo.length > MAX_STEPS) this.#undo.shift();
  }

  /** The step undo takes, off the stack. */
  popUndo(): OrderedHistoryStep | null {
    const step = this.#undo.pop() ?? null;
    this.#restored = step?.selection ?? null;
    return step;
  }

  /** An undone step, for redo. */
  pushRedo(step: OrderedHistoryStep): void {
    this.#redo.push(step);
    if (this.#redo.length > MAX_STEPS) this.#redo.shift();
  }

  popRedo(): OrderedHistoryStep | null {
    return this.#redo.pop() ?? null;
  }

  /** Forget the undone steps: what they would make again no longer follows the room. */
  clearRedo(): void {
    this.#redo.length = 0;
  }

  /** The selection the last undo restores, when it took back the step's last commit alone. */
  restoreSelection(selection: CollaborationLocalSelection | null): void {
    this.#restored = selection;
  }

  /** Put a step back where it was taken from, when its undo or redo was refused. */
  restoreUndo(step: OrderedHistoryStep): void {
    this.#undo.push(step);
  }

  /**
   * A step's change made again under `id`: the addresses, suggestions and list numbers its
   * step's earlier changes minted in `doc` carry their new ids. The result is untrusted until the
   * validator reads it.
   */
  again(change: OrderedChange, id: string, doc: OrderedDocument): Record<string, unknown> {
    const out = renamed(change, this.#ids, renamedLists(doc, this.#ids));

    // A step redone twice: what named the first redo's ids names the second's
    for (const [old, now] of this.#ids) if (now === change.id) this.#ids.set(old, id);
    this.#ids.set(change.id, id);
    if (typeof out !== 'object' || out === null || Array.isArray(out)) return {};
    return { ...out, id };
  }

  /** Forget everything: a new Word file, or a reload. */
  clear(): void {
    this.#undo.length = 0;
    this.#redo.length = 0;
    this.#ids.clear();
    this.#restored = null;
  }
}
