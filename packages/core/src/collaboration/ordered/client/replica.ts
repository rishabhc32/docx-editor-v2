// The client's copy of the room (design D7, D10): the list, the person's pending changes each
// with its undo record, and the room's changes since the last save, also with their records.

import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedChangeId, orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange, OrderedStampedChange } from '../change/types.ts';
import type { OrderedChangeOutcome } from '../change/apply.ts';
import type { OrderedPerson } from '../model/person.ts';
import { errorName, OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord, type OrderedUndoRecord } from '../model/undo-record.ts';
import { orderedStampDate } from '../room/room.ts';
import { noteDirty } from './dirty.ts';
import { OrderedDirty } from './publish.ts';
import { OrderedClientRefusal, type OrderedClientRefusalReport } from './refusal.ts';
import { OrderedRoomClock } from './room-clock.ts';

/** A change of this person the room has not ordered yet. */
interface Pending {
  readonly change: OrderedChange;
  /** Its undo record while it is applied here; null while it does not apply. */
  record: OrderedUndoRecord | null;
  sent: boolean;
  /** What it wrote when it was last applied. */
  dirty: OrderedDirty;
  /** The local time it was made, and its date until the room stamps it: kept on every apply. */
  readonly made: number;
  readonly date: string;
}

/** A change the room ordered after its last save. */
interface Confirmed {
  readonly change: OrderedStampedChange;
  readonly record: OrderedUndoRecord;
  readonly own: boolean;
}

/** The room's change does not apply on this copy: the copies differ. Terminal. */
export class OrderedReplicaDiverged extends Error {
  readonly n: number;

  constructor(n: number, detail: string) {
    super(`the room's change ${n} does not apply here: ${detail}`);
    this.name = 'OrderedReplicaDiverged';
    this.n = n;
  }
}

/** What a cycle did: what to publish, and the pending changes it refused. */
export interface OrderedCycle {
  readonly dirty: OrderedDirty;
  readonly refused: readonly OrderedClientRefusalReport[];
}

export interface OrderedReplicaOptions {
  readonly doc: OrderedDocument;
  /** The person as the room stamps them (`by`). Own suggestions are found by their ID. */
  readonly person: OrderedPerson;
  readonly namespace: string;
  /** Milliseconds since 1970, for the date of a change until the room stamps it. */
  readonly now: () => number;
  /**
   * Called when a pending change throws something other than `OrderedRefusal` (a defect) as it
   * is applied again. The change was taken back and refused with `apply-failed`; the host logs
   * the error here.
   */
  readonly onApplyError?: (error: unknown, change: OrderedChange) => void;
  /** The room's relaxation of content-control locks, from `welcome`. */
  readonly contentLockPolicy?: ContentLockPolicy;
}

export class OrderedReplica {
  readonly doc: OrderedDocument;
  readonly person: OrderedPerson;
  readonly namespace: string;
  readonly #clock: OrderedRoomClock;
  readonly #onApplyError: ((error: unknown, change: OrderedChange) => void) | undefined;
  readonly #contentLockPolicy: ContentLockPolicy | undefined;
  #counter = 0;
  #pending: Pending[] = [];
  /** The room's changes after its last save, with their records: kept until the room saves them. */
  #tail: Confirmed[] = [];
  /** The last order number applied: without its pending changes, the copy is the room at it. */
  confirmed = 0;
  /** The last order number the room's storage holds, as far as this copy heard. */
  saved = 0;

  constructor(options: OrderedReplicaOptions) {
    this.doc = options.doc;
    this.person = options.person;
    this.namespace = options.namespace;
    this.#clock = new OrderedRoomClock(options.now);
    this.#onApplyError = options.onApplyError;
    this.#contentLockPolicy = options.contentLockPolicy;
  }

  /**
   * A new change id in the person's namespace, past every counter of it the copy holds: a host may
   * keep a namespace across a reload, and the room refuses a counter it already applied (D7).
   */
  nextId(): string {
    this.#counter = Math.max(this.#counter, this.doc.lastCounters.get(this.namespace) ?? 0) + 1;
    return orderedChangeId(this.namespace, this.#counter);
  }

  get pendingCount(): number {
    return this.#pending.length;
  }

  /** The pending changes, in order. */
  pending(): readonly OrderedChange[] {
    return this.#pending.map((p) => p.change);
  }

  /**
   * The pending changes not sent yet, marked sent, in order up to the first that `ready` holds
   * back (a change whose media the room does not store yet): the order never changes.
   */
  takeUnsent(ready: (change: OrderedChange) => boolean = () => true): readonly OrderedChange[] {
    const out: OrderedChange[] = [];
    for (const p of this.#pending)
      if (!p.sent) {
        if (!ready(p.change)) break;
        p.sent = true;
        out.push(p.change);
      }

    return out;
  }

  /** After the connection dropped: nothing counts as sent, the next connection sends all. */
  markAllUnsent(): void {
    for (const p of this.#pending) p.sent = false;
  }

  /** Note the paragraphs a change touches before and after `write` applies or takes it back. */
  #noted<T>(change: OrderedChange, dirty: OrderedDirty, write: () => T): T {
    noteDirty(this.doc, change, dirty);
    const value = write();
    noteDirty(this.doc, change, dirty);
    return value;
  }

  #apply(change: OrderedChange, person: OrderedPerson, date: string): OrderedUndoRecorded {
    return withUndoRecord(this.doc, () =>
      applyOrderedChange(this.doc, change, {
        person,
        date,
        ids: orderedIdMinter(change.id),
        ...(this.#contentLockPolicy === undefined
          ? {}
          : { contentLockPolicy: this.#contentLockPolicy }),
      })
    );
  }

  /**
   * Apply a validated change of this person here and queue it, unsent. Refused by name when the
   * list refuses it; nothing changed then.
   */
  applyLocal(change: OrderedChange): {
    readonly outcome: OrderedChangeOutcome;
    readonly dirty: OrderedDirty;
  } {
    const dirty = new OrderedDirty();

    // The room's clock as its stamps tell it: the room's date is then most often this one
    const made = this.#clock.local();
    const date = orderedStampDate(this.#clock.stampFor(made));
    const recorded = this.#noted(change, dirty, () => {
      try {
        return this.#apply(change, this.person, date);
      } catch (error) {
        if (error instanceof OrderedRefusal) {
          throw new OrderedClientRefusal('ordered-refused', error.reason, error.detail);
        }

        throw error;
      }
    });
    this.#pending.push({ change, record: recorded.record, sent: false, dirty, made, date });
    return { outcome: recorded.value, dirty };
  }

  /**
   * Take back the last `count` pending changes, which must be unsent (a gate batch that failed
   * halfway, or one the editor did not commit). Returns what they had written.
   */
  dropLast(count: number): OrderedDirty {
    const dirty = new OrderedDirty();
    for (let k = 0; k < count; k += 1) {
      const last = this.#pending[this.#pending.length - 1];
      if (!last || last.sent) throw new Error('Only unsent pending changes can be dropped.');
      if (last.record) takeBack(this.doc, last.record);
      dirty.add(last.dirty);
      this.#pending.pop();
    }

    return dirty;
  }

  #takeBackPending(dirty: OrderedDirty): void {
    for (let k = this.#pending.length - 1; k >= 0; k -= 1) {
      const p = this.#pending[k];
      if (!p) continue;
      dirty.add(p.dirty);
      if (p.record) takeBack(this.doc, p.record);
      p.record = null;
    }
  }

  /**
   * Apply the pending changes again. One that no longer applies is refused: one not sent yet
   * goes, one sent stays until the room orders or refuses it (the room applies it in its own
   * order). One that throws a defect is refused the same way, by name, so it cannot end the
   * session at each join.
   */
  #applyPending(dirty: OrderedDirty, refused: OrderedClientRefusalReport[]): void {
    const kept: Pending[] = [];
    for (const p of this.#pending) {
      const written = new OrderedDirty();
      noteDirty(this.doc, p.change, written);
      try {
        p.record = this.#apply(p.change, this.person, p.date).record;
        noteDirty(this.doc, p.change, written);
        p.dirty = written;
        dirty.add(written);
      } catch (error) {
        p.record = null;
        refused.push(this.#pendingRefusal(p.change, error));
      }

      if (p.record !== null || p.sent) kept.push(p);
    }

    this.#pending = kept;
  }

  /**
   * Why a pending change does not apply again. A throw other than a refusal is a defect: the
   * change is refused all the same (its record took its writes back), so one change cannot end
   * the session at each join; the host logs the error, and the report names only its type.
   */
  #pendingRefusal(change: OrderedChange, error: unknown): OrderedClientRefusalReport {
    if (error instanceof OrderedRefusal)
      return {
        code: 'ordered-refused',
        reason: 'pending-refused',
        detail: error.reason,
        change: change.id,
      };
    try {
      this.#onApplyError?.(error, change);
    } catch {
      // A failing host callback must not stop the session
    }

    return {
      code: 'ordered-refused',
      reason: 'apply-failed',
      detail: errorName(error),
      change: change.id,
    };
  }

  /**
   * Take the pending changes back (last first), run `between` (the room's changes, in order),
   * and apply the pending ones again.
   */
  #cycle(between: (dirty: OrderedDirty) => void): OrderedCycle {
    const dirty = new OrderedDirty();
    const refused: OrderedClientRefusalReport[] = [];
    this.#takeBackPending(dirty);
    try {
      between(dirty);
    } finally {
      this.#applyPending(dirty, refused);
    }

    return { dirty, refused };
  }

  /**
   * The room's changes, in order from `confirmed + 1`. Throws `OrderedReplicaDiverged` when one
   * does not apply (the copy is then the room at the order number before it, without pending).
   */
  receive(changes: readonly OrderedStampedChange[]): OrderedCycle {
    return this.#cycle((dirty) => {
      for (const change of changes) {
        if (change.n <= this.confirmed) continue;
        if (change.n !== this.confirmed + 1)
          throw new OrderedReplicaDiverged(change.n, `a gap after ${this.confirmed}`);
        const record = this.#noted(change, dirty, () => {
          try {
            return this.#apply(change, change.by, change.at).record;
          } catch (error) {
            if (error instanceof OrderedRefusal) {
              throw new OrderedReplicaDiverged(change.n, error.reason);
            }

            throw error;
          }
        });
        const mine = this.#pending.find((p) => p.change.id === change.id);
        const own = mine !== undefined;
        if (mine) {
          this.#clock.observe(mine.made, change.at);
          this.#pending = this.#pending.filter((p) => p !== mine);
        }

        this.#tail.push({ change, record, own: own || this.#ownId(change.id) });
        this.confirmed = change.n;
      }
    });
  }

  /** The room refused these pending changes: they go, and the others are applied again. */
  refuse(ids: ReadonlySet<string>): OrderedCycle {
    return this.#cycle(() => {
      this.#pending = this.#pending.filter((p) => !ids.has(p.change.id));
    });
  }

  /** The room's storage holds every change up to `n`: their records are not needed again. */
  markSaved(n: number): void {
    if (n <= this.saved) return;
    this.saved = Math.min(n, this.confirmed);
    this.#tail = this.#tail.filter((entry) => entry.change.n > this.saved);
  }

  /**
   * Go back to the last save, for a new connection: the room may have lost what came after it
   * (a restart). The person's own changes after it become pending again, ahead of the pending
   * ones; the room answers those it still has by id.
   */
  rollbackToSaved(): OrderedDirty {
    const dirty = new OrderedDirty();
    this.#takeBackPending(dirty);
    const own: Pending[] = [];
    for (let k = this.#tail.length - 1; k >= 0; k -= 1) {
      const entry = this.#tail[k];
      if (!entry) continue;
      this.#noted(entry.change, dirty, () => takeBack(this.doc, entry.record));
      if (entry.own)
        own.unshift({
          change: withoutStamp(entry.change),
          record: null,
          sent: false,
          dirty: new OrderedDirty(),
          made: this.#clock.local(),
          // The room's own stamp: a room that still holds the change answers with it
          date: entry.change.at,
        });
    }

    this.#tail = [];
    this.confirmed = this.saved;
    this.#pending = [...own, ...this.#pending];
    this.markAllUnsent();
    return dirty;
  }

  /** Apply the pending changes again after `rollbackToSaved`, before the catch-up arrives. */
  reapplyPending(): OrderedCycle {
    const dirty = new OrderedDirty();
    const refused: OrderedClientRefusalReport[] = [];
    this.#applyPending(dirty, refused);
    return { dirty, refused };
  }

  #ownId(id: string): boolean {
    return id.startsWith(`${this.namespace}:`);
  }
}

type OrderedUndoRecorded = {
  readonly value: OrderedChangeOutcome;
  readonly record: OrderedUndoRecord;
};

/** A stamped change as its sender sent it: without the room's `n`, `by` and `at`. */
function withoutStamp(change: OrderedStampedChange): OrderedChange {
  const { n: _n, by: _by, at: _at, ...rest } = change;
  return rest;
}
