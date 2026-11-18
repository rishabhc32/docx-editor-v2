// Every op the editor can make, and what the ordered client makes of it (design D13, task 9.1).

import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { TreeDocOp, TreeDocOpKind } from '../../../../store/store/tree-op-types.ts';
import type {
  OrderedChangeKind,
  OrderedCommentChange,
  OrderedDecideChange,
  OrderedEditChange,
  OrderedRetractChange,
} from '../../change/types.ts';
import type { OrderedChangeMedia } from '../../change/resources.ts';
import type { OrderedPerson } from '../../model/person.ts';
import type { OrderedDocument } from '../../model/types.ts';
import type { OrderedPolicyCode, OrderedPolicyReason } from '../../room/policy.ts';
import type { ListParagraph } from '../paragraphs.ts';
import type { OrderedClientRefusalReport } from '../refusal.ts';

type Body<T> = T extends unknown ? Omit<T, 'v' | 'id' | 'into' | 'mode'> : never;

/** A change of content without the fields the session adds (`v`, `id`, `mode`, `into`). */
export type OrderedEditBody = Body<OrderedEditChange>;

/** A change without the fields the session adds (`v`, `id`), for each change of a union. */
export type WithoutIds<T> = T extends unknown ? Omit<T, 'v' | 'id'> : never;

/**
 * A decision, a take-back of the sender's own suggestion (a reject of it) or a new comment,
 * without the fields the session adds (`v`, `id`).
 */
export type OrderedReviewBody = WithoutIds<
  OrderedDecideChange | OrderedRetractChange | OrderedCommentChange
>;

/** What one translated op becomes, before the session adds its id. */
export type OrderedChangeBody = OrderedEditBody | OrderedReviewBody;

/** Whether a body is an edit, which the session gives the batch's mode and suggestion. */
export function isEditBody(body: OrderedChangeBody): body is OrderedEditBody {
  return body.kind !== 'decide' && body.kind !== 'retract' && body.kind !== 'comment';
}

/** The group of `tasks.md` that builds an op. `6.1` built the first ones. */
export type OrderedOpGroup = '6.1' | '9.3' | '9.4' | '9.5' | '9.6' | '9.7' | '9.8' | '9.9' | '9.10';

/**
 * How an op says whether it is tracked.
 *
 * - `attribution`: a suggestion when it carries a revision attribution (`revision`).
 * - `either`: the same tracked or not; the batch's other ops say.
 * - `direct`: always an edit with tracking off.
 * - `review`: no edit at all: a decision or a comment.
 */
export type OrderedOpTracking = 'attribution' | 'either' | 'direct' | 'review';

/**
 * The room's party rule that refuses an op for a party, as the host's own check also refuses it
 * (a party's change must be a suggestion that taking back gives the document
 * before it). Every edit also meets the rules for all edits: nothing in a header or a footer
 * (`policy-furniture`), nothing inside another person's suggestion (`policy-nested`).
 */
export interface OrderedPartyRule {
  readonly code: OrderedPolicyCode;
  /** The reason the host's check gives (`not_a_suggestion:<reason>`), where it gives one. */
  readonly reason?: OrderedPolicyReason;
  /** Why the rule applies: what the engine writes for it in suggesting mode. */
  readonly why: string;
}

/** No party rule: a party makes the op as a suggestion, under the rules for every edit. */
export const PARTY_SUGGESTS: OrderedPartyRule | null = null;

/** A decision is the organiser's alone. */
export const DECIDE_RULE: OrderedPartyRule = {
  code: 'policy-decide',
  why: 'a party never decides; the organiser does',
};

/** What a review write gives besides its ops: the text of the comment it adds. */
export interface OrderedReviewContext {
  readonly commentText?: string;
}

/** What a translator reads besides its op. */
export interface OrderedTranslateBatch {
  readonly doc: OrderedDocument;
  /** The person the room stamps the batch with: whose suggestions are their own. */
  readonly person: OrderedPerson;
  /** The document's default paragraph style id, for an Enter that names no style for the tail. */
  readonly defaultParagraphStyle: string | null;
  /** Every op of the batch: an op whose meaning needs the ops beside it reads them here. */
  readonly ops: readonly TreeDocOp[];
  /**
   * The list's paragraph for an editor node id. The first time a batch names it, the editor's
   * text must be the list's: offsets mean the same units only then.
   */
  paragraph(nodeId: string): ListParagraph;
  /** The package the editor shows, as the client last published it; null before the first. */
  shown(): OoxmlPackage | null;
  /**
   * The address of a link relationship the editor minted on its own package for the batch (a
   * new link's), by its owner part and id, or null.
   */
  hyperlinkTarget(part: string, relationshipId: string): string | null;
  /**
   * The package the ops apply to, as the editor gave it to the gate (a paste's merged styles,
   * numbering, links and notes), or null when the editor gave none.
   */
  gated(): OoxmlPackage | null;
  /** The review write the batch is, when the editor gave one (a comment's text). */
  readonly review: OrderedReviewContext | null;
  /**
   * The image a relationship of `part` names in the package the ops apply to (`gated`): a new
   * picture's bytes, held by the session under their hash, as a change names them. Null when
   * that package has no internal image there. The kind's spec names the hash in `blobs`; the
   * session sends the bytes before the change (design D14). With `heldMedia`, the entry points
   * for every change that names media.
   *
   * @throws OrderedClientRefusal when the bytes are not an image the room stores.
   */
  media(part: string, relationshipId: string): OrderedChangeMedia | null;
  /**
   * Bytes an op carries itself (a picture watermark's image), held by the session under their
   * hash as `media` holds an image of the package.
   *
   * @throws OrderedClientRefusal when the bytes are not an image the room stores.
   */
  heldMedia(bytes: Uint8Array, contentType: string): OrderedChangeMedia;
  /**
   * The media of a part of the package the ops apply to, by its bytes there: the file's media a
   * checkpoint names, by its reference (the room's storage holds it, so a copy needs no upload,
   * even before the copy reads it), or else the bytes held as `heldMedia` holds them. Null when
   * the part is neither.
   *
   * @throws OrderedClientRefusal when the bytes are not media of `contentType` the room stores.
   */
  partMedia(
    part: string,
    bytes: Uint8Array | undefined,
    contentType: string
  ): OrderedChangeMedia | null;
  /** The safe external address a hyperlink relationship of `part` names there, or null. */
  link(part: string, relationshipId: string): string | null;
  /**
   * Report what the batch leaves as it is, by name, without refusing the batch (the marks a
   * review write reaches that name no suggestion: a style's record the room does not hold, marks
   * in part of an equation). The editor's write there is taken back by the next publish.
   */
  leave(report: OrderedClientRefusalReport): void;
}

/**
 * One step of a batch: the change bodies of one op, read from the list as the steps before it
 * left it. Each body is applied before the next is read, so a step that yields its bodies one
 * at a time reads the list as its own earlier bodies left it.
 */
export type OrderedTranslateStep = () => Iterable<OrderedChangeBody>;

/**
 * An op's translator. It checks the op's options at once and throws `OrderedClientRefusal` for
 * one the format does not say, so a refused batch applies nothing; the step it returns reads
 * the list when it runs.
 */
export type OrderedOpTranslator = (
  op: TreeDocOp,
  batch: OrderedTranslateBatch
) => OrderedTranslateStep;

interface OrderedOpEntryBase {
  /** The group of `tasks.md` that builds it (or built it). */
  readonly group: OrderedOpGroup;
  /** For a party: the room's rule that refuses it, or null when the room has none for it. */
  readonly party: OrderedPartyRule | null;
  /** What the op does, and what the engine writes for it. */
  readonly note: string;
  /** How the op says it is tracked. Default `attribution`. */
  readonly tracking?: OrderedOpTracking;
}

/** An op the client translates. */
export interface OrderedTranslatedOp extends OrderedOpEntryBase {
  readonly status: 'translated';
  /** The change kinds it becomes. */
  readonly changes: readonly OrderedChangeKind[];
  readonly translate: OrderedOpTranslator;
}

/** An op a later group builds; until then the client refuses it by its name. */
export interface OrderedPlannedOp extends OrderedOpEntryBase {
  readonly status: 'planned';
  readonly group: Exclude<OrderedOpGroup, '6.1' | '9.3'>;
  /** The change kinds it will become: existing kinds, or the planned kind's name. */
  readonly changes: readonly string[];
}

/** One op's entry. */
export type OrderedOpEntry = OrderedTranslatedOp | OrderedPlannedOp;

/**
 * A group's ops. Each group of `tasks.md` keeps its ops in its own file under `client/ops/`: the
 * op's change kinds, the room's party rule that refuses it for a party, its state, and its
 * translator. The table in `index.ts` joins the files and is typed over every `TreeDocOp` kind,
 * so a new op does not compile until a group lists it, and the guard test fails until it is
 * translated.
 */
export type OrderedOpTable<K extends TreeDocOpKind> = { readonly [P in K]: OrderedOpEntry };

/**
 * An option that changes what an op means (tracking, a legacy form field, a field result, a
 * link's relationship, a new list definition). A planned option is refused by name: `refusal`
 * is the refusal's `<code>:<reason>`.
 */
export interface OrderedOpOption {
  readonly op: TreeDocOpKind;
  /** The option's name, as the guard and `design.md` list it. */
  readonly option: string;
  readonly status: 'translated' | 'planned';
  readonly group: OrderedOpGroup;
  /** `<code>:<reason>` of the refusal the client gives while it is planned. */
  readonly refusal?: string;
  readonly note: string;
}

/** Whether an op is of one kind. */
export function isOp<K extends TreeDocOpKind>(
  op: TreeDocOp,
  kind: K
): op is Extract<TreeDocOp, { readonly op: K }> {
  return op.op === kind;
}

/** A translator for one kind of op, which reads the op as its own kind. */
export function translator<K extends TreeDocOpKind>(
  kind: K,
  translate: (
    op: Extract<TreeDocOp, { readonly op: K }>,
    batch: OrderedTranslateBatch
  ) => OrderedTranslateStep
): OrderedOpTranslator {
  return (op, batch) => {
    if (!isOp(op, kind)) throw new TypeError(`A ${kind} translator got a ${op.op} op.`);
    return translate(op, batch);
  };
}
