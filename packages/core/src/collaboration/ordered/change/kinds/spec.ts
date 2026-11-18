// What a group's change kind says about itself (design, "How to add a group"). The format's
// first kinds are handled by name in the shared files: the validator, the targets, the apply,
// the suggestions-only policy, the publish and the caret. A kind a group adds (tasks 9.4 to
// 9.10) is one spec in the group's own file under `change/kinds/`, and every shared file asks
// the registry (`registry.ts`) for it. A group then adds a kind without editing a shared file.

import type { OrderedEditOutcome } from '../../edits/apply.ts';
import type { OrderedEditContext } from '../../edits/context.ts';
import type {
  OrderedContainerToken,
  OrderedDocument,
  OrderedItem,
  OrderedStory,
} from '../../model/types.ts';
import type { OrderedMaker } from '../../model/person.ts';
import type { OrderedPolicyVerdict } from '../../room/policy.ts';
import type { Fields } from '../read.ts';
import type { OrderedChangeAddress, OrderedChangePlace, OrderedEditChange } from '../types.ts';

/**
 * How a spec names what its change touches. Each call checks the address against the document
 * and adds what it finds; an address that does not resolve refuses the change by name before
 * the policy or the apply runs.
 */
export interface OrderedTargetReader {
  /** A caret the change types or inserts at. */
  caret(place: OrderedChangePlace): void;
  /** Every item of `[from, to)`, hidden ones too. */
  range(from: OrderedChangePlace, to: OrderedChangePlace): void;
  /**
   * The item whose first unit is at `address`, of `kind` (and one of `tokens` for a container),
   * as one the change edits in place. Returned for a spec that reads it further.
   */
  item(
    address: OrderedChangeAddress,
    kind: 'mark' | 'open' | 'atom' | 'zero',
    tokens?: readonly OrderedContainerToken[]
  ): OrderedItem;
}

/** What a spec's apply gets: the document, the change's story, and the edit's context. */
export interface OrderedKindApply {
  readonly doc: OrderedDocument;
  readonly story: OrderedStory;
  /**
   * The author, the date, the change's id minter, `direct` for an edit made with tracking off,
   * and `join` for the suggestion the change adds to (`into`).
   */
  readonly edit: OrderedEditContext;
}

/** What the suggestions-only policy reads of a kind (design D9). */
export interface OrderedKindPolicy {
  /**
   * The edit changes the text a reader sees (the host's reason `text` for a party's direct edit);
   * otherwise only properties or the package (`package`).
   */
  readonly text: boolean;
  /**
   * The edit sets properties: it is inside another person's suggestion also on units that
   * person struck (Word nests the record inside the deletion).
   */
  readonly properties: boolean;
  /**
   * The edit adds a container beside its target (a row or a column): it is inside another
   * person's suggestion only when the target's container is that person's insertion.
   */
  readonly beside?: boolean;
  /**
   * The edit strikes what it touches (a deletion, a new value or target): units another
   * person struck already stay as they are (policy 1), so they are not inside their suggestion.
   */
  readonly strikes?: boolean;
  /**
   * Whether the edit adds, changes or removes a header or footer part: the host's check refuses
   * a party's as `furniture` before any other rule (group 9.8).
   */
  readonly furniture?: (change: OrderedEditChange, doc: OrderedDocument) => boolean;
  /**
   * A party's verdict for a kind that names suggestions, not places (a change of their
   * attribution), read before the rules for places; null when those rules decide.
   */
  readonly party?: (
    change: OrderedEditChange,
    doc: OrderedDocument,
    maker: OrderedMaker
  ) => OrderedPolicyVerdict | null;
}

/**
 * One change kind a group adds. `C` is the kind's change type. Methods take the change as the
 * validator returned it, so a spec reads its own fields without checks.
 */
export interface OrderedKindSpec<C extends OrderedEditChange = OrderedEditChange> {
  readonly kind: C['kind'];
  /** The kind's own payload fields, past `v`, `id`, `kind`, `part`, `mode` and `into`. */
  readonly fields: readonly string[];
  /**
   * Read and bound the payload: a new object with the kind's own fields only, each through a
   * reader of `read.ts` or `values.ts` (or the group's own), which throws `Invalid` by name.
   */
  read(fields: Fields): Record<string, unknown>;
  /** Name what the change touches, through the reader. */
  touches(change: C, reader: OrderedTargetReader): void;
  /**
   * Apply the change inside the caller's undo record. It writes through the model's
   * mutations only and throws `OrderedRefusal` to refuse; the record then takes back every
   * write.
   */
  apply(change: C, input: OrderedKindApply): OrderedEditOutcome;
  readonly policy: OrderedKindPolicy;
  /**
   * What the publish after it composes: the paragraphs it touches, the whole story (a change
   * that can move a block edge: a table, a row, a note, a story block), or what its own
   * suggestions label, in every story (a change that writes in several: a converted note).
   */
  readonly publish: 'paragraphs' | 'story' | 'suggestions';
  /**
   * Stories besides its own that the change writes, composed whole by the publish after it (a
   * watermark is in every header). Read before and after the write.
   */
  stories?(change: C, doc: OrderedDocument): readonly string[];
  /**
   * The suggestions the change rewrites in place (their attribution): the publish draws the
   * paragraphs they label again, in every story.
   */
  rewrites?(change: C): readonly string[];
  /** Where the change leaves the caret, or null to leave it where it was. */
  caret?(change: C): OrderedChangePlace | null;
  /** The change types content: a batch's caret follows the last content it typed. */
  readonly types?: boolean;
  /**
   * The media the change names, by SHA-256: the room refuses it while its storage does not
   * hold one of them (`missing-blob`, design D14).
   */
  blobs?(change: C): readonly string[];
}

/**
 * A spec of one kind, checked against its change type: the registry holds every group's specs
 * in one table, so each is stored as a spec of any edit and asks for its own kind back.
 */
export function orderedKindSpec<C extends OrderedEditChange>(
  spec: OrderedKindSpec<C>
): OrderedKindSpec {
  const own = (change: OrderedEditChange): change is C => change.kind === spec.kind;
  const mine = (change: OrderedEditChange): C => {
    if (!own(change)) throw new TypeError(`A ${spec.kind} spec got a ${change.kind} change.`);
    return change;
  };
  const { stories, caret, blobs, rewrites } = spec;
  return {
    kind: spec.kind,
    fields: spec.fields,
    read: (fields) => spec.read(fields),
    touches: (change, reader) => spec.touches(mine(change), reader),
    apply: (change, input) => spec.apply(mine(change), input),
    policy: spec.policy,
    publish: spec.publish,
    ...(stories === undefined
      ? {}
      : { stories: (change: OrderedEditChange, doc) => stories(mine(change), doc) }),
    ...(caret === undefined ? {} : { caret: (change: OrderedEditChange) => caret(mine(change)) }),
    ...(spec.types === undefined ? {} : { types: spec.types }),
    ...(blobs === undefined ? {} : { blobs: (change: OrderedEditChange) => blobs(mine(change)) }),
    ...(rewrites === undefined
      ? {}
      : { rewrites: (change: OrderedEditChange) => rewrites(mine(change)) }),
  };
}
