// Why the ordered client refused a local edit, an undo or a pending change. The gate returns
// the code to the editor before the editor commits; the host hears the code, a stable reason
// and a short detail through `onRefused`.

/**
 * The ordered client's refusal codes. Each is also a `CollaborationFailureCode`, which the gate
 * returns.
 *
 * - `ordered-unsupported-op`: the edit holds an op the change format does not express.
 * - `ordered-opaque-write`: an editor write the session cannot see (a review write).
 * - `ordered-refused`: the list, the room or the client refused the edit; `reason` names why.
 *
 * @public
 */
export type OrderedClientRefusalCode =
  | 'ordered-unsupported-op'
  | 'ordered-opaque-write'
  | 'ordered-refused';

/**
 * What the client refused, for the host. `reason` is stable: an op name for
 * `ordered-unsupported-op`; for `ordered-refused` the model's `OrderedRefusalReason`, the
 * room's refusal code, or one of the client's own:
 *
 * - `out-of-step`: the editor names a paragraph or an offset the list does not show.
 * - `text-differs`: the editor's text of a paragraph is not the list's.
 * - `invalid-change`: the change format refuses the translated change (its code follows).
 * - `engine-refused`: the engine's own plan refuses the op, as the editor alone does (a table
 *   authoring plan's reason, a section index the engine does not count).
 * - `unknown-story`: the list has no main document story for the op.
 * - `media-too-large`: a picture is over the room's cap for its type (`detail` is the type). A
 *   host may show it as "too large".
 * - `invalid-media`: a picture's bytes are not an image the room stores, or the session has no
 *   media store.
 * - `pending-refused`: a pending change no longer applies after the room's changes.
 * - `apply-failed`: a pending change threw something other than a refusal (a defect) as it was
 *   applied again; `detail` is the error's type. The room's own refusal has the same name.
 * - `no-inverse`: undo of a direct edit, or of a suggestion someone decided.
 * - `redo-refused`: a step to redo no longer applies.
 * - `ungated-write`: the editor wrote without asking the gate; the list's version stays.
 * - `stale-decision`: the card changed since the organiser saw it.
 * - `not-in-room`: a review write reaches change records outside the stories that the room holds
 *   no suggestion for (a style's record the engine cannot decide both ways). They stay as they
 *   are; the cards the room holds are decided, and the batch is refused only when it reaches
 *   nothing else.
 * - `external-picture-link`: a pasted picture names its image only by an address. Word loads
 *   the address when it opens the file, so a room never writes one (design D18); a picture with
 *   an embedded copy keeps the copy and travels without the address.
 *
 * @public
 */
export interface OrderedClientRefusalReport {
  readonly code: OrderedClientRefusalCode;
  readonly reason: string;
  /** A short note for logs. Never holds document text. */
  readonly detail?: string;
  /** The change id, when one change was refused. */
  readonly change?: string;
}

/** A refusal thrown inside the client, caught where the editor or the host is answered. */
export class OrderedClientRefusal extends Error {
  readonly code: OrderedClientRefusalCode;
  readonly reason: string;
  readonly detail: string | undefined;

  constructor(code: OrderedClientRefusalCode, reason: string, detail?: string) {
    super(detail === undefined ? `${code}: ${reason}` : `${code}: ${reason}: ${detail}`);
    this.name = 'OrderedClientRefusal';
    this.code = code;
    this.reason = reason;
    this.detail = detail;
  }

  report(change?: string): OrderedClientRefusalReport {
    return {
      code: this.code,
      reason: this.reason,
      ...(this.detail === undefined ? {} : { detail: this.detail.slice(0, 512) }),
      ...(change === undefined ? {} : { change }),
    };
  }
}
