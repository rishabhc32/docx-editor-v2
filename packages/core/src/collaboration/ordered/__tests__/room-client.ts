// A client of the room for tests: it applies its own changes at once, sends them, and on the
// room's changes takes its pending ones back, applies the room's, and applies its pending ones
// again, as the ordered client will (design D7).

import type { OrderedPerson } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord, type OrderedUndoRecord } from '../model/undo-record.ts';
import { orderedChangeId, orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { decodeOrderedServerMessage, encodeOrderedClientMessage } from '../protocol/codec.ts';
import type { OrderedActor } from '../room/policy.ts';
import { realApply } from './room-harness.ts';

interface Pending {
  readonly change: OrderedChange;
  record: OrderedUndoRecord | null;
  sent: boolean;
}

/** The fields a client fills in for a change: everything but `v` and `id`. */
export type ChangeBody = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

/**
 * A client of the room: it applies its own changes at once, sends them, and on the room's
 * changes takes its pending ones back, applies the room's, and applies its pending ones again.
 */
export class TestClient {
  readonly doc: OrderedDocument;
  readonly actor: OrderedActor;
  readonly received: number[] = [];
  readonly refusedIds: string[] = [];
  last = 0;
  #counter = 0;
  #pending: Pending[] = [];

  constructor(doc: OrderedDocument, who: OrderedActor) {
    this.doc = doc;
    this.actor = who;
  }

  get pending(): readonly OrderedChange[] {
    return this.#pending.map((p) => p.change);
  }

  /** Make a change and apply it here at once. Null when it does not apply here. */
  local(body: ChangeBody): OrderedChange | null {
    this.#counter += 1;
    const change = {
      v: 1,
      id: orderedChangeId(this.actor.namespace, this.#counter),
      ...body,
    } as OrderedChange;
    const record = this.#apply(change, this.actor.person, '2026-10-28T00:00:00Z');
    if (record === null) return null;
    this.#pending.push({ change, record, sent: false });
    return change;
  }

  /** The frame with the changes not sent yet, or null. */
  flush(): string | null {
    const unsent = this.#pending.filter((p) => !p.sent);
    if (unsent.length === 0) return null;
    for (const p of unsent) p.sent = true;
    return encodeOrderedClientMessage({
      v: 3,
      type: 'changes',
      changes: unsent.map((p) => p.change),
    });
  }

  /** Take a frame from the room. */
  receive(frame: string): void {
    const decoded = decodeOrderedServerMessage(frame);
    if (!decoded.ok)
      throw new Error(
        `the client cannot read a frame: ${decoded.error.code} ${decoded.error.path}`
      );
    const message = decoded.value;
    if (message.type === 'changes') {
      this.#takeBackPending();
      for (const change of message.changes) {
        if (change.n !== this.last + 1) throw new Error(`a gap before ${change.n}`);
        if (this.#apply(change, change.by, change.at) === null)
          throw new Error(`the room's change ${change.n} does not apply here`);
        this.last = change.n;
        this.received.push(change.n);
        this.#pending = this.#pending.filter((p) => p.change.id !== change.id);
      }
      this.#applyPending();
    } else if (message.type === 'refuse') {
      const refused = new Set(message.refusals.map((r) => r.id));
      for (const r of message.refusals) if (r.id !== null) this.refusedIds.push(r.id);
      this.#takeBackPending();
      this.#pending = this.#pending.filter((p) => !refused.has(p.change.id));
      this.#applyPending();
    }
  }

  #apply(change: OrderedChange, person: OrderedPerson, date: string): OrderedUndoRecord | null {
    const ctx = { person, date, ids: orderedIdMinter(change.id) };
    try {
      return withUndoRecord(this.doc, () => realApply(this.doc, change, ctx)).record;
    } catch (error) {
      if (error instanceof OrderedRefusal) return null;
      throw error;
    }
  }

  #takeBackPending(): void {
    for (let k = this.#pending.length - 1; k >= 0; k -= 1) {
      const p = this.#pending[k];
      if (p?.record) takeBack(this.doc, p.record);
      if (p) p.record = null;
    }
  }

  /** Apply pending changes again; one that no longer applies here is dropped. */
  #applyPending(): void {
    const kept: Pending[] = [];
    for (const p of this.#pending) {
      p.record = this.#apply(p.change, this.actor.person, '2026-10-28T00:00:00Z');
      if (p.record !== null || p.sent) kept.push(p);
    }
    this.#pending = kept;
  }
}
