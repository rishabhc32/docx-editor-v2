// The room core (design D9): one document, one order. It has no I/O: the host reads frames,
// calls `submit`, writes the accepted changes to its log, broadcasts them, and answers the
// sender with the acks and refusals. It runs in the same process as the host's connection
// handling, one `submit` at a time.

import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import { contentLockPolicyOf } from '../../../store/store/content-lock-policy.ts';
import {
  mintedChange,
  ORDERED_NAMESPACE_PATTERN,
  orderedIdMinter,
  parseOrderedChangeId,
  type OrderedIdMinter,
} from '../change/ids.ts';
import { orderedChangeTargets, type OrderedTargetErrorCode } from '../change/targets.ts';
import { orderedKindSpecOf } from '../change/kinds/registry.ts';
import {
  isOrderedEditChange,
  ORDERED_CHANGE_LIMITS,
  type OrderedChange,
  type OrderedStampedChange,
} from '../change/types.ts';
import { validateOrderedBatch, validateOrderedStampedChange } from '../change/validate.ts';
import type { OrderedChangeErrorCode } from '../change/read.ts';
import { isFileMediaHash } from '../model/file-media.ts';
import { checkedPerson, type OrderedPerson } from '../model/person.ts';
import { errorName, OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import type { OrderedChangeOutcome } from '../change/apply.ts';
import type {
  OrderedActor,
  OrderedPolicyCode,
  OrderedPolicyInput,
  OrderedPolicyReason,
  OrderedPolicyVerdict,
  OrderedRoomPolicy,
} from './policy.ts';

/** What `applyChange` gets besides the change: the stamp and the ids to mint. @public */
export interface OrderedApplyContext {
  /** The person the change is written under: their Word name is its `w:author`. */
  readonly person: OrderedPerson;
  /** The change's date (`w:date`). */
  readonly date: string;
  /** The ids the change mints, fresh for this application. */
  readonly ids: OrderedIdMinter;
  /** Told what the change did, once it holds: for the policy's `checkApplied`. */
  readonly onApplied?: (outcome: OrderedChangeOutcome) => void;
  /** The room's relaxation of content-control locks: every copy applies with the same one. */
  readonly contentLockPolicy?: ContentLockPolicy;
}

/**
 * Apply one change to a document, in place. It runs inside an undo record: it writes only
 * through the model's mutations, and throws `OrderedRefusal` to refuse a change it cannot
 * apply. The client and the room use the same function, so every copy makes the same document.
 * It passes the context on to `applyOrderedChange`, which tells the room what the change did; a
 * policy with `checkApplied` refuses a change whose apply told it nothing.
 *
 * @public
 */
export type OrderedApplyChange = (
  doc: OrderedDocument,
  change: OrderedChange,
  context: OrderedApplyContext
) => void;

/** How the room is set up. @public */
export interface OrderedRoomOptions {
  /** The document as loaded from the room's Word file. The room owns it from now on. */
  readonly document: OrderedDocument;
  readonly applyChange: OrderedApplyChange;
  readonly policy: OrderedRoomPolicy;
  /** The clock: milliseconds since 1970. Written as `w:date` to the second. */
  readonly now: () => number;
  /** The changes after the Word file, in order, to rebuild the room after a restart. */
  readonly log?: readonly OrderedStampedChange[];
  /**
   * The media the room's storage holds, by SHA-256. A change that names media the room does
   * not hold is refused (`missing-blob`). Without it, no change may name media but the file's
   * media a checkpoint names (`resources.fileMedia`), which the storage holds.
   */
  readonly blobs?: { has(sha256: string): boolean };
  /**
   * Called when `applyChange` throws something other than `OrderedRefusal` (a defect). The
   * change was taken back and refused with `apply-failed`; the host logs the error here.
   */
  readonly onApplyError?: (error: unknown, change: OrderedStampedChange) => void;
  /**
   * Let the organiser's direct formats pass a content lock, as the editor's option of the same
   * name lets an untracked restyle pass. The host sends it to every client in `welcome`.
   * Default: Word's lock, which refuses every format there.
   */
  readonly contentLockPolicy?: ContentLockPolicy;
}

/** An accepted or repeated change id and its order number. @public */
export interface OrderedAck {
  readonly id: string;
  readonly n: number;
}

/**
 * Why the room refused a change. Stable.
 *
 * - The validator's codes (`OrderedChangeErrorCode`): the batch is malformed.
 * - `stamp`: the sender's namespace belongs to another person (another person ID).
 * - `missing-blob`: the change names media the room's storage does not hold (design D14).
 * - The target codes (`OrderedTargetErrorCode`): an address, a story or a suggestion is not in
 *   the document, an address names the wrong kind of unit, or a range runs backwards. Also
 *   `other-file` for a change whose counter is at or under the one the room's file holds for its
 *   namespace: it was made before the room's last close (design D7).
 * - The policy codes (`OrderedPolicyCode`).
 * - `apply-refused`: the model refused the change; `reason` names why (`OrderedRefusalReason`).
 * - `apply-failed`: applying the change threw something else; it was taken back.
 *
 * @public
 */
export type OrderedRoomRefusalCode =
  | OrderedChangeErrorCode
  | OrderedTargetErrorCode
  | OrderedPolicyCode
  | 'stamp'
  | 'missing-blob'
  | 'apply-refused'
  | 'apply-failed';

/** A refused change. @public */
export interface OrderedRoomRefusal {
  /** The change's id; null when the batch is refused before any id could be read. */
  readonly id: string | null;
  readonly code: OrderedRoomRefusalCode;
  /** The stable cause: the policy's `OrderedPolicyReason`, or the model's refusal reason. */
  readonly reason?: string;
  /** A short note for logs. */
  readonly detail?: string;
}

/** What `submit` returns for the host to log, broadcast and answer. @public */
export interface OrderedSubmitResult {
  /** The changes the room applied, in order: log them, then send them to everyone. */
  readonly accepted: readonly OrderedStampedChange[];
  /** Order numbers for the sender: new changes and repeated ids, in batch order. */
  readonly acks: readonly OrderedAck[];
  /** Refused changes, in batch order. */
  readonly refused: readonly OrderedRoomRefusal[];
}

/**
 * Why a room could not be rebuilt from its log. Stable.
 *
 * - `log-invalid`: an entry is not a valid ordered change.
 * - `log-gap`: an order number is missing or out of order.
 * - `log-repeated-id`: one change id twice.
 * - `log-stamp`: one namespace under two person IDs.
 * - `log-apply`: an entry does not apply to the document, or it was made before the room's file
 *   (`other-file`).
 *
 * @public
 */
export type OrderedRoomOpenCode =
  | 'log-invalid'
  | 'log-gap'
  | 'log-repeated-id'
  | 'log-stamp'
  | 'log-apply';

/** A room that cannot open: the log does not rebuild the document. @public */
export class OrderedRoomOpenError extends Error {
  readonly code: OrderedRoomOpenCode;
  /** The order number of the entry, or its position when it has none. */
  readonly at: number;

  constructor(code: OrderedRoomOpenCode, at: number, detail: string) {
    super(`${code} at ${at}: ${detail}`);
    this.name = 'OrderedRoomOpenError';
    this.code = code;
    this.at = at;
  }
}

/** A clock reading as Word's `w:date`: `YYYY-MM-DDTHH:MM:SSZ`. @public */
export function orderedStampDate(ms: number): string {
  if (!Number.isFinite(ms)) throw new RangeError('The room clock gave no time.');
  return new Date(Math.floor(ms / 1000) * 1000).toISOString().replace('.000Z', 'Z');
}

const bounded = (text: string): string => (text.length > 512 ? `${text.slice(0, 512)}...` : text);

/** A validator error's note: its path, its detail, or both. */
function detailOf(path: string, detail: string | undefined): string {
  if (detail === undefined) {
    return path;
  }

  return path === '' ? detail : `${path}: ${detail}`;
}

type Outcome =
  | { readonly ok: true; readonly change: OrderedStampedChange }
  | { readonly ok: false; readonly refusal: OrderedRoomRefusal };

/**
 * One room's document and order. Neutral: no I/O, no timers.
 *
 * @public
 */
export class OrderedRoom {
  readonly #doc: OrderedDocument;
  readonly #apply: OrderedApplyChange;
  readonly #policy: OrderedRoomPolicy;
  readonly #now: () => number;
  readonly #onApplyError: ((error: unknown, change: OrderedStampedChange) => void) | undefined;
  readonly #blobs: { has(sha256: string): boolean } | undefined;
  readonly #contentLockPolicy: ContentLockPolicy | undefined;
  readonly #log: OrderedStampedChange[] = [];
  /** Change id to order number: answers a repeated id. Rebuilt from the log. */
  readonly #ids = new Map<string, number>();
  /**
   * Namespace to the ID of the person who owns it: from the document (a checkpoint names the
   * person of each namespace that typed), then rebuilt from the log.
   */
  readonly #namespaces = new Map<string, string>();
  /**
   * The highest counter of each namespace the room's file holds (`lastCounters` as it opened): a
   * change at or under it was made before the file.
   */
  readonly #fileCounters: ReadonlyMap<string, number>;

  constructor(options: OrderedRoomOptions) {
    this.#doc = options.document;
    this.#apply = options.applyChange;
    this.#policy = options.policy;
    this.#now = options.now;
    this.#onApplyError = options.onApplyError;
    this.#blobs = options.blobs;
    this.#contentLockPolicy = contentLockPolicyOf(options.contentLockPolicy);
    if (this.#doc.undo !== null) throw new Error('The document has an open undo record.');
    for (const [namespace, person] of this.#doc.namespaces) {
      this.#namespaces.set(namespace, person.id);
    }

    this.#fileCounters = new Map(this.#doc.lastCounters);
    for (const [k, entry] of (options.log ?? []).entries()) this.#replay(entry, k);
  }

  /** The last order number: 0 before the first change. */
  get order(): number {
    return this.#log.length;
  }

  /** The room's document. Read it; write it only through `submit`. */
  get document(): OrderedDocument {
    return this.#doc;
  }

  /** The order number of an applied change id, if any. */
  orderOf(id: string): number | undefined {
    return this.#ids.get(id);
  }

  /**
   * The ID of the person who owns a namespace, if any change from it was applied or the document
   * names its person.
   */
  namespaceOwner(namespace: string): string | undefined {
    return this.#namespaces.get(namespace);
  }

  /**
   * The changes after order number `after`, for a client that joins or comes back: all of
   * them for 0, none past the last. The host splits them into frames under the frame cap.
   */
  joinSnapshot(after: number): readonly OrderedStampedChange[] {
    if (!Number.isSafeInteger(after) || after < 0)
      throw new RangeError('A join needs an order number from 0.');
    return this.#log.slice(after);
  }

  /**
   * Check, order and apply a batch from one client. `batch` is untrusted: the decoded
   * `changes` of a message, as it came.
   *
   * Per change, in order:
   *  1. The batch is checked as a whole: its shape, caps, ids and namespace. One malformed change
   *     refuses the whole batch and nothing from it is applied. A correct client checks its own
   *     changes with the same validator before it sends them, so only a broken or hostile client
   *     sends one, and a partial batch from such a client would leave its copy in a state it
   *     cannot predict.
   *  2. An id the room already applied is answered with its order number and applied again never.
   *  3. The change's targets must exist in the document (`orderedChangeTargets`).
   *  4. The policy gives its verdict, reading only the change and its targets.
   *  5. The change is applied inside an undo record by the injected `applyChange`. A throw takes
   *     it back exactly and refuses that change by name. The policy may also judge what the
   *     change did (`checkApplied`); its refusal takes the change back the same way.
   *  6. The change gets the next order number, `by` from the login and `at` from the clock.
   */
  submit(actor: OrderedActor, batch: unknown): OrderedSubmitResult {
    const sender = checkedActor(actor);
    const valid = validateOrderedBatch(batch, {
      namespace: sender.namespace,
      parts: this.#doc.stories,
    });
    if (!valid.ok) {
      const { code, path, detail } = valid.error;
      return {
        accepted: [],
        acks: [],
        refused: readableIds(batch).map((id) => ({
          id,
          code,
          detail: bounded(detailOf(path, detail)),
        })),
      };
    }

    const owner = this.#namespaces.get(sender.namespace);
    if (owner !== undefined && owner !== sender.person.id) {
      const refused = valid.value.map((change) => ({
        id: change.id,
        code: 'stamp' as const,
        reason: 'stamp',
        detail: 'the namespace belongs to another person',
      }));
      return { accepted: [], acks: [], refused };
    }

    // One reading per batch, before anything applies: a clock that fails changes nothing
    const at = orderedStampDate(this.#now());
    const accepted: OrderedStampedChange[] = [];
    const acks: OrderedAck[] = [];
    const refused: OrderedRoomRefusal[] = [];

    // All or nothing per change, not per batch: a batch is a burst of keys, not one intention,
    // and refusing it whole for one change inside someone else's suggestion would throw away
    // valid typing. A refusal names the change, so the client takes back exactly that one. A
    // change that depends on a refused one fails on its own targets and is refused by name too.
    for (const change of valid.value) {
      const known = this.#ids.get(change.id);
      if (known !== undefined) {
        acks.push({ id: change.id, n: known });
        continue;
      }

      const outcome = this.#take(sender, change, at);
      if (!outcome.ok) {
        refused.push(outcome.refusal);
        continue;
      }

      accepted.push(outcome.change);
      acks.push({ id: change.id, n: outcome.change.n });
    }

    return { accepted, acks, refused };
  }

  #take(actor: OrderedActor, change: OrderedChange, at: string): Outcome {
    const late = this.#beforeFile(change);
    if (late !== null) {
      return { ok: false, refusal: late };
    }

    const targets = orderedChangeTargets(this.#doc, change);
    if (!targets.ok)
      return {
        ok: false,
        refusal: { id: change.id, code: targets.code, detail: bounded(targets.detail) },
      };

    // Media a change names must be in the room's storage before the change applies: what the
    // host stored, or the file's media a checkpoint names
    const missing = blobsOf(change).find(
      (hash) => this.#blobs?.has(hash) !== true && !isFileMediaHash(this.#doc, hash)
    );
    if (missing !== undefined)
      return { ok: false, refusal: { id: change.id, code: 'missing-blob', detail: missing } };
    const input: OrderedPolicyInput = {
      actor,
      change,
      touched: targets.touched,
      document: this.#doc,
    };
    const verdict = judged(() => this.#policy.check(input));
    if (!verdict.ok) return { ok: false, refusal: policyRefusal(change.id, verdict) };
    const stamped = {
      ...change,
      n: this.#log.length + 1,
      by: actor.person,
      at,
    } as OrderedStampedChange;
    const policy = this.#policy;
    const applied = this.#applyOne(
      stamped,
      policy.checkApplied && ((did) => judged(() => policy.checkApplied?.(input, did) ?? ALLOWED))
    );
    if (applied !== null) return { ok: false, refusal: applied };
    this.#record(stamped, actor.namespace);
    return { ok: true, change: stamped };
  }

  /**
   * Apply inside an undo record; null when it applied, else the refusal (nothing changed). With
   * `judge`, the policy's verdict on what the change did: a refusal takes the change back.
   */
  #applyOne(
    change: OrderedStampedChange,
    judge?: (outcome: OrderedChangeOutcome) => OrderedPolicyVerdict
  ): OrderedRoomRefusal | null {
    let outcome: OrderedChangeOutcome | null = null;
    const context: OrderedApplyContext = {
      person: change.by,
      date: change.at,
      ids: orderedIdMinter(change.id),
      onApplied: (did) => {
        outcome = did;
      },
      ...(this.#contentLockPolicy === undefined
        ? {}
        : { contentLockPolicy: this.#contentLockPolicy }),
    };
    try {
      const { record } = withUndoRecord(this.#doc, () => this.#apply(this.#doc, change, context));
      let verdict = ALLOWED;
      if (judge !== undefined) {
        // A host's apply that does not pass the context on reports nothing: refused, unjudged
        verdict = outcome === null ? UNREPORTED : judge(outcome);
      }

      if (verdict.ok) {
        return null;
      }

      takeBack(this.#doc, record);
      return policyRefusal(change.id, verdict);
    } catch (error) {
      if (error instanceof OrderedRefusal)
        return {
          id: change.id,
          code: 'apply-refused',
          reason: error.reason,
          ...(error.detail === undefined ? {} : { detail: bounded(error.detail) }),
        };
      try {
        this.#onApplyError?.(error, change);
      } catch {
        // A failing log hook must not stop the batch halfway: the changes before it are applied
      }

      return { id: change.id, code: 'apply-failed', detail: errorName(error) };
    }
  }

  /**
   * The refusal of a change made before the room's file, or of one that undoes a change made
   * before it (design D7), or null. An editor that opens the file has no undo of such an edit.
   */
  #beforeFile(change: OrderedChange): OrderedRoomRefusal | null {
    for (const made of [change.id, ...undoneBy(change)]) {
      const parsed = parseOrderedChangeId(made);
      const saved = parsed === null ? undefined : this.#fileCounters.get(parsed.namespace);
      if (parsed !== null && saved !== undefined && parsed.counter <= saved) {
        const detail = `the room's file holds ${parsed.namespace}:${saved}`;
        return { id: change.id, code: 'other-file', detail };
      }
    }

    return null;
  }

  #record(change: OrderedStampedChange, namespace: string): void {
    this.#log.push(change);
    this.#ids.set(change.id, change.n);
    this.#namespaces.set(namespace, change.by.id);
  }

  #replay(entry: OrderedStampedChange, k: number): void {
    const checked = validateOrderedStampedChange(entry, { parts: this.#doc.stories });
    if (!checked.ok) {
      throw new OrderedRoomOpenError(
        'log-invalid',
        k + 1,
        `${checked.error.code} ${checked.error.path}`
      );
    }

    const change = checked.value;
    if (change.n !== this.#log.length + 1) {
      throw new OrderedRoomOpenError('log-gap', change.n, `expected ${this.#log.length + 1}`);
    }

    if (this.#ids.has(change.id)) {
      throw new OrderedRoomOpenError('log-repeated-id', change.n, change.id);
    }

    const namespace = parseOrderedChangeId(change.id)?.namespace ?? '';
    const owner = this.#namespaces.get(namespace);
    if (owner !== undefined && owner !== change.by.id) {
      throw new OrderedRoomOpenError('log-stamp', change.n, namespace);
    }

    const refusal = this.#beforeFile(change) ?? this.#applyOne(change);
    if (refusal !== null) {
      throw new OrderedRoomOpenError(
        'log-apply',
        change.n,
        `${refusal.code} ${refusal.reason ?? ''}`
      );
    }

    this.#record(change, namespace);
  }
}

/** The earlier changes a change undoes: a revert's direct edits, a take-back's own edits. */
function undoneBy(change: OrderedChange): readonly string[] {
  if (change.kind === 'revert') {
    return change.suggestions.flatMap((id) => mintedChange(id, 'S') ?? []);
  }

  if (change.kind === 'retract') {
    return change.undo ?? [];
  }

  return [];
}

/** The media a change names, by SHA-256: a group's kind says so in its spec. */
function blobsOf(change: OrderedChange): readonly string[] {
  if (!isOrderedEditChange(change)) return [];
  return orderedKindSpecOf(change.kind)?.blobs?.(change) ?? [];
}

/** The host's actor under a valid namespace, with its person trimmed and checked. */
function checkedActor(actor: OrderedActor): OrderedActor {
  if (!ORDERED_NAMESPACE_PATTERN.test(actor.namespace)) {
    throw new RangeError('The actor has no valid namespace.');
  }

  return { ...actor, person: checkedPerson(actor.person) };
}

function policyRefusal(
  id: string,
  verdict: {
    readonly code: OrderedPolicyCode;
    readonly reason?: OrderedPolicyReason;
    readonly detail?: string;
  }
): OrderedRoomRefusal {
  return {
    id,
    code: verdict.code,
    ...(verdict.reason === undefined ? {} : { reason: verdict.reason }),
    ...(verdict.detail === undefined ? {} : { detail: bounded(verdict.detail) }),
  };
}

const ALLOWED: OrderedPolicyVerdict = Object.freeze({ ok: true });
const UNREPORTED: OrderedPolicyVerdict = Object.freeze({
  ok: false,
  code: 'policy-failed',
  detail: 'the apply did not report what the change did',
});

/** A policy's verdict; a policy that throws refuses the change (`policy-failed`). */
function judged(verdict: () => OrderedPolicyVerdict): OrderedPolicyVerdict {
  try {
    return verdict();
  } catch (error) {
    return { ok: false, code: 'policy-failed', detail: errorName(error) };
  }
}

/** The well-formed change ids a refused batch holds, or one null when there are none to read. */
function readableIds(batch: unknown): (string | null)[] {
  if (!Array.isArray(batch) || batch.length > ORDERED_CHANGE_LIMITS.batch) return [null];
  const ids: string[] = [];
  for (const item of batch) {
    if (typeof item !== 'object' || item === null) continue;
    const id: unknown = Object.getOwnPropertyDescriptor(item, 'id')?.value;
    if (typeof id === 'string' && parseOrderedChangeId(id) !== null && !ids.includes(id))
      ids.push(id);
  }

  return ids.length === 0 ? [null] : ids;
}
