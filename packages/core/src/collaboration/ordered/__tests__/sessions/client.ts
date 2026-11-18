// A client of a random session, as the ordered client works (design D7): it applies its own
// changes to its copy at once, each in an undo record, and sends them. On the room's changes it
// takes its pending changes back (last first), applies the room's changes in order, and applies
// its pending changes again. A pending change that no longer applies is refused and counted:
// one not sent yet is dropped; one sent stays until the room orders or refuses it.

import { applyOrderedChange } from '../../change/apply.ts';
import { orderedChangeId, orderedIdMinter } from '../../change/ids.ts';
import type { OrderedChange } from '../../change/types.ts';
import { validateOrderedChange } from '../../change/validate.ts';
import type { OrderedPerson } from '../../model/person.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { takeBack, withUndoRecord, type OrderedUndoRecord } from '../../model/undo-record.ts';
import { decodeOrderedServerMessage, encodeOrderedClientMessage } from '../../protocol/codec.ts';
import type { OrderedActor } from '../../room/policy.ts';
import type { ChangeBody } from './random-change.ts';

/** The date a client gives its own change until the room stamps it. */
const LOCAL_DATE = '2026-10-28T00:00:00Z';

interface Pending {
  readonly change: OrderedChange;
  record: OrderedUndoRecord | null;
  sent: boolean;
}

/** What a client counts. */
export interface ClientCounts {
  /** Changes the format refused before applying (never sent). */
  invalid: number;
  /** Changes the client's own copy refused before sending (never sent). */
  local: number;
  /** Pending changes that no longer applied after the room's changes. */
  reapply: number;
  /** Changes the room refused. */
  room: number;
}

export class SessionClient {
  readonly doc: OrderedDocument;
  readonly actor: OrderedActor;
  readonly counts: ClientCounts = { invalid: 0, local: 0, reapply: 0, room: 0 };
  /** The last order number applied: the copy without pending changes is the room at it. */
  last = 0;
  #counter = 0;
  #pending: Pending[] = [];
  /** Called with the copy at `last` (pending changes taken back), after each changes frame. */
  readonly #confirmed: (client: SessionClient) => void;

  constructor(
    doc: OrderedDocument,
    actor: OrderedActor,
    confirmed: (client: SessionClient) => void
  ) {
    this.doc = doc;
    this.actor = actor;
    this.#confirmed = confirmed;
  }

  get pending(): readonly OrderedChange[] {
    return this.#pending.map((p) => p.change);
  }

  /**
   * Make a change and apply it here at once. Null when the format or this copy refuses it: a
   * correct client checks a change with the room's own validator before it applies it.
   */
  local(body: ChangeBody): OrderedChange | null {
    this.#counter += 1;
    const checked = validateOrderedChange(
      { v: 1, id: orderedChangeId(this.actor.namespace, this.#counter), ...body },
      { namespace: this.actor.namespace, parts: this.doc.stories }
    );
    if (!checked.ok) {
      this.counts.invalid += 1;
      return null;
    }
    const change = checked.value;
    const record = this.#apply(change, this.actor.person, LOCAL_DATE);
    if (record === null) {
      this.counts.local += 1;
      return null;
    }
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
    if (!decoded.ok) throw new Error(`a frame the client cannot read: ${decoded.error.code}`);
    const message = decoded.value;
    if (message.type === 'changes') {
      this.#takeBackPending();
      for (const change of message.changes) {
        if (change.n !== this.last + 1) throw new Error(`a gap before ${change.n}`);
        if (this.#apply(change, change.by, change.at) === null)
          throw new Error(`the room's change ${change.n} (${change.kind}) does not apply here`);
        this.last = change.n;
        this.#pending = this.#pending.filter((p) => p.change.id !== change.id);
      }
      this.#confirmed(this);
      this.#applyPending();
    } else if (message.type === 'refuse') {
      const refused = new Set(message.refusals.map((r) => r.id));
      this.counts.room += message.refusals.length;
      this.#takeBackPending();
      this.#pending = this.#pending.filter((p) => !refused.has(p.change.id));
      this.#applyPending();
    }
  }

  #apply(change: OrderedChange, person: OrderedPerson, date: string): OrderedUndoRecord | null {
    const ctx = { person, date, ids: orderedIdMinter(change.id) };
    try {
      return withUndoRecord(this.doc, () => applyOrderedChange(this.doc, change, ctx)).record;
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

  #applyPending(): void {
    const kept: Pending[] = [];
    for (const p of this.#pending) {
      p.record = this.#apply(p.change, this.actor.person, LOCAL_DATE);
      if (p.record === null) this.counts.reapply += 1;
      if (p.record !== null || p.sent) kept.push(p);
    }
    this.#pending = kept;
  }
}
