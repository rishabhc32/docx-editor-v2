// The room's policy: a verdict for each change, before it is applied, and optionally one on what
// it did once applied. A policy reads only the change, the items it touches (`OrderedTouched`)
// and what the apply reports, never the whole document, so a check costs what the change
// covers. The room runs it once per change, in order, so a change sees the document after the
// changes before it.

import type { OrderedChange } from '../change/types.ts';
import type { OrderedChangeOutcome } from '../change/apply.ts';
import type { OrderedTouched } from '../change/targets.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedDocument } from '../model/types.ts';

/**
 * Who sends a change, as the host's login gives it. One person can have several clients, each
 * with its own namespace.
 *
 * @public
 */
export interface OrderedActor {
  /** The person as the login gives them: stamped on each change; owns their suggestions by ID. */
  readonly person: OrderedPerson;
  /** This client's change id namespace. One person owns it in a room. */
  readonly namespace: string;
  /** May only look: sends presence, no change. */
  readonly readOnly: boolean;
  /** Only adds suggestions: a party in a negotiation. */
  readonly suggests: boolean;
  /** May accept and reject suggestions: an organiser. */
  readonly decides: boolean;
  /**
   * May add and answer comments, and resolve them when it is an organiser (it decides, and does
   * not only suggest): the host gives it only while comments are open. A read-only person sends
   * no comment.
   */
  readonly comments: boolean;
}

/**
 * Why a policy refused a change. Stable.
 *
 * - `policy-read-only`: the person may only look.
 * - `policy-decide`: a decision from someone who may not decide.
 * - `policy-furniture`: a party's change in a header or a footer.
 * - `policy-nested`: a party's change inside another person's pending suggestion.
 * - `policy-direct`: a party's edit that is not a suggestion.
 * - `policy-not-own`: a change that adds to or takes back another person's suggestion.
 * - `policy-comment`: a comment the person may not write: comments are not open to them
 *   (`comments`), a resolve from someone who may not decide (`resolve`), a reply to or a resolve
 *   of a resolved thread (`resolved`), or a comment past the document's 5,000 (`limit`).
 * - `policy-failed`: the policy itself failed (a host lookup threw); the change is refused.
 *
 * @public
 */
export type OrderedPolicyCode =
  | 'policy-read-only'
  | 'policy-decide'
  | 'policy-furniture'
  | 'policy-nested'
  | 'policy-direct'
  | 'policy-not-own'
  | 'policy-comment'
  | 'policy-failed';

/**
 * The finer reason a policy gives a refusal, where it gives one, for a host that reports it in
 * its own refusal message (such as `not_a_suggestion:<reason>`): `furniture`, `nested`, `stamp`
 * (written under another person's id), `text` and `package` (an edit that is not a suggestion),
 * `others_suggestions`, and a comment's `comments`, `resolve`, `resolved` and `limit`.
 *
 * @public
 */
export type OrderedPolicyReason =
  | 'furniture'
  | 'nested'
  | 'stamp'
  | 'text'
  | 'package'
  | 'others_suggestions'
  | 'comments'
  | 'resolve'
  | 'resolved'
  | 'limit';

/** A policy's verdict on one change. @public */
export type OrderedPolicyVerdict =
  | { readonly ok: true }
  | {
      readonly ok: false;
      readonly code: OrderedPolicyCode;
      readonly reason?: OrderedPolicyReason;
      readonly detail?: string;
    };

/** What a policy reads. The document is for lookups by id; a policy writes nothing. @public */
export interface OrderedPolicyInput {
  readonly actor: OrderedActor;
  readonly change: OrderedChange;
  readonly touched: OrderedTouched;
  readonly document: OrderedDocument;
}

/** A per-change policy. @public */
export interface OrderedRoomPolicy {
  check(input: OrderedPolicyInput): OrderedPolicyVerdict;
  /**
   * A second verdict, on what the change did, once `check` allowed it and it applied: the room's
   * rule compares the document before and after a change. A refusal takes it back.
   */
  checkApplied?(input: OrderedPolicyInput, outcome: OrderedChangeOutcome): OrderedPolicyVerdict;
}

/**
 * A policy that allows every change from anyone who may write. For tests and trusted hosts.
 *
 * @public
 */
export const allowWritersPolicy: OrderedRoomPolicy = Object.freeze({
  check: ({ actor }: OrderedPolicyInput): OrderedPolicyVerdict =>
    actor.readOnly ? { ok: false, code: 'policy-read-only' } : { ok: true },
});
