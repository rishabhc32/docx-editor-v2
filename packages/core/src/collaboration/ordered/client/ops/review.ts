// Review ops (task 9.3): accept and reject become `decide` changes, a new comment's markers a
// `comment` change.

import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { holdsMathRunContent } from '../../../../store/store/revision-math-runs.ts';
import { revisionGroupKey } from '../../../../store/store/tree-op-revisions.ts';
import { revisionKey } from '../../../../store/store/tree-op-tracked-adjacency.ts';
import { orderedCardOf } from '../../card-of.ts';
import { orderedCardSeen } from '../../change/apply-decide.ts';
import type { OrderedSeenCard } from '../../change/types.ts';
import { isPending, madeBy, storyOfItem } from '../../model/document.ts';
import { makerOf, type OrderedMaker } from '../../model/person.ts';
import {
  fileRecordIdsOf,
  REVISION_ROLE,
  revisionIdFor,
  revisionIdsOf,
} from '../../model/revision-ids.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { attributeOf, isWml } from '../../ooxml-nodes.ts';
import { placeAt } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import {
  DECIDE_RULE,
  translator,
  type OrderedChangeBody,
  type OrderedOpTable,
  type OrderedTranslateBatch,
} from './types.ts';

/** A revision as the editor names it. */
interface RevisionTriple {
  readonly id: string;
  readonly author: string;
  readonly date?: string;
}

/** A revision triple, with no date when the mark has none. */
function tripleOf(id: string, author: string, date: string | undefined): RevisionTriple {
  return { id, author, ...(date === undefined ? {} : { date }) };
}

/** Every element under a part's root, the root first, in one walk with no recursion. */
function* elementsOf(root: OoxmlNode): Generator<OoxmlElement> {
  const stack: OoxmlNode[] = [root];
  for (let node = stack.pop(); node !== undefined; node = stack.pop()) {
    if (node.kind === 'textValue') {
      continue;
    }

    yield node;
    for (const child of node.children) {
      if (child.kind !== 'textValue') {
        stack.push(child);
      }
    }
  }
}

/**
 * The suggestions of this copy by revision triple, read once for each author and date a
 * decision names: every suggestion of that author and date, whatever its status; callers pass
 * over the decided ones. The editor names a revision by its `w:id`, `w:author` and `w:date`, or
 * by the node ids of its marks in the package the client last published (the editor shows that
 * package, so its node ids are the published ones). Compose writes a suggestion's `w:id` for each
 * role as a function of the suggestion's id (`revisionIdFor`), so the triple names one suggestion.
 */
export class TripleIndex {
  readonly #doc: OrderedDocument;
  readonly #byStamp = new Map<string, Map<string, string[]>>();

  constructor(doc: OrderedDocument) {
    this.#doc = doc;
  }

  /** The suggestions a triple names, pending or decided. */
  suggestions(triple: RevisionTriple): readonly string[] {
    const stamp = `${triple.author}\u0000${triple.date ?? ''}`;
    let ids = this.#byStamp.get(stamp);
    if (!ids) {
      ids = new Map();
      for (const suggestion of this.#doc.suggestions.values()) {
        // A change record outside the stories is named by its marks (`outsideRoom`)
        if (this.#doc.partRevisions.has(suggestion.id)) continue;
        if (suggestion.attribution.author !== triple.author) continue;
        if ((suggestion.attribution.date ?? '') !== (triple.date ?? '')) continue;
        for (const id of revisionIdsOf(this.#doc, suggestion.id)) {
          const list = ids.get(id);
          if (list) list.push(suggestion.id);
          else ids.set(id, [suggestion.id]);
        }
      }

      this.#byStamp.set(stamp, ids);
    }

    return ids.get(triple.id) ?? [];
  }
}

/**
 * A revision mark as a host names it: the part it is in, its `w:id`, `w:author` and `w:date`,
 * and in a part outside the stories (the styles part) its change element's local name.
 *
 * @public
 */
export interface OrderedRevisionName {
  /** The part name, such as `/word/document.xml` or `/word/styles.xml`. */
  readonly part: string;
  /** The mark's `w:id`. */
  readonly id: string;
  /** The mark's `w:author`. */
  readonly author: string;
  /** The mark's `w:date`; absent when the mark has none. */
  readonly date?: string;
  /**
   * The change element's local name, such as `rPrChange`. Read only outside the stories, where
   * one card is the records of one element (`OrderedPartRevision`).
   */
  readonly element?: string;
}

/**
 * The pending suggestions a revision mark names, by creation order: the ones whose composed or
 * file `w:id` is `revision.id` in that part, with that author and date. The index is the one the
 * editor's decisions use, so a host's card (a review item's mark) maps to exactly the
 * suggestions the editor's accept of it decides. In a story part, only suggestions with items
 * there; outside the stories, the change records of `doc.partRevisions` by part, element and the
 * three values, so a style's record is never a body suggestion with the same three values.
 * Empty for a decided suggestion or an unknown mark.
 *
 * @public
 */
export function orderedSuggestionsOfRevision(
  doc: OrderedDocument,
  revision: OrderedRevisionName
): readonly string[] {
  const { part, id, author, date } = revision;
  const triple = tripleOf(id, author, date);
  if (!doc.stories.has(part)) {
    if (revision.element === undefined) return [];
    const found = partRecordIndex(doc).get(partRecordKey(part, revision.element, triple));
    return found !== undefined && isPending(doc, found) ? [found] : [];
  }

  const inPart = (suggestion: string) => {
    for (const item of doc.labelled.get(suggestion) ?? [])
      if (storyOfItem(doc, item)?.name === part) return true;
    return false;
  };
  const seqOf = (suggestion: string) => doc.suggestions.get(suggestion)?.seq ?? 0;
  return new TripleIndex(doc)
    .suggestions(triple)
    .filter((suggestion) => isPending(doc, suggestion) && inPart(suggestion))
    .sort((a, b) => seqOf(a) - seqOf(b));
}

/** The nodes of the shown package with these ids, read in one bounded walk of its parts. */
function nodesOf(pkg: OoxmlPackage, ids: readonly string[]): Map<string, OoxmlNode> {
  const wanted = new Set(ids);
  const found = new Map<string, OoxmlNode>();
  for (const part of pkg.parts.values()) {
    for (const node of elementsOf(part.root)) {
      if (!wanted.has(node.id)) {
        continue;
      }

      found.set(node.id, node);
      if (found.size === wanted.size) {
        return found;
      }
    }
  }

  return found;
}

/**
 * Whether the shown package holds a mark of this triple around a math run's content. The list
 * keeps Word's marks on some runs of an equation as file content (it has no unit inside an
 * equation; a whole equation marked is a suggestion on it, `load/equations.ts`), so such a mark
 * names no suggestion and its decision is refused by name (`acceptRevision.equation-part`).
 */
function equationPartShown(batch: OrderedTranslateBatch, triple: RevisionTriple): boolean {
  const shown = batch.shown();
  if (shown === null) {
    return false;
  }

  for (const part of shown.parts.values()) {
    for (const node of elementsOf(part.root)) {
      if (
        holdsMathRunContent(node) &&
        attributeOf(node, 'id') === triple.id &&
        attributeOf(node, 'author') === triple.author &&
        attributeOf(node, 'date') === triple.date
      ) {
        return true;
      }
    }
  }

  return false;
}

/** The triple of the pending suggestion whose table record a `w:tblGridChange` writes. */
function gridRecordTriple(doc: OrderedDocument, node: OoxmlElement): RevisionTriple | null {
  const id = attributeOf(node, 'id');
  if (id === undefined) return null;
  for (const suggestion of doc.suggestions.values())
    if (
      suggestion.status === 'pending' &&
      (revisionIdFor(doc, suggestion.id, REVISION_ROLE.tableProperties) === id ||
        fileRecordIdsOf(doc, suggestion.id).includes(id))
    ) {
      const { author, date } = suggestion.attribution;
      return tripleOf(id, author, date);
    }

  return null;
}

/** The revision triples of marks named by node id in the shown package. */
export function triplesOfSites(batch: OrderedTranslateBatch, siteNodeIds: readonly string[]) {
  const shown = batch.shown();
  if (shown === null)
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'nothing published yet');
  const nodes = nodesOf(shown, siteNodeIds);
  const triples = new Map<string, RevisionTriple>();
  for (const id of siteNodeIds) {
    const node = nodes.get(id);
    if (!node || node.kind === 'textValue')
      throw new OrderedClientRefusal(
        'ordered-refused',
        'out-of-step',
        `no mark ${id.slice(0, 64)}`
      );

    // A grid record names no author (`CT_Markup`): its suggestion is found by its `w:id`
    if (isWml(node, 'tblGridChange')) {
      const grid = gridRecordTriple(batch.doc, node);
      if (!grid)
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a grid record decided');
      triples.set(revisionKey(grid.id, grid.author, grid.date), grid);
      continue;
    }

    const revision = attributeOf(node, 'id');
    const author = attributeOf(node, 'author');
    if (revision === undefined || author === undefined)
      throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a mark with no id');

    // Every mark of one suggestion names its triple: each is read once
    const triple = tripleOf(revision, author, attributeOf(node, 'date'));
    triples.set(revisionKey(triple.id, triple.author, triple.date), triple);
  }

  return [...triples.values()];
}

type DecisionKind =
  | 'acceptRevision'
  | 'rejectRevision'
  | 'acceptAllRevisions'
  | 'rejectAllRevisions';
type DecisionOp = Extract<TreeDocOp, { readonly op: DecisionKind }>;

const isDecisionOp = (op: TreeDocOp): op is DecisionOp =>
  op.op === 'acceptRevision' ||
  op.op === 'rejectRevision' ||
  op.op === 'acceptAllRevisions' ||
  op.op === 'rejectAllRevisions';

/** A mark a decision op names outside the stories, and the suggestion that holds it, if any. */
interface OutsideMark {
  readonly part: string;
  /** The change record's suggestion (a style's); null for one the list lacks. */
  readonly suggestion: string | null;
}

/**
 * The marks among `ids` in parts the room does not hold as stories, by node id. A style's change
 * record is a suggestion of its own, found by its part, element, `w:id`, author and date: the
 * three values alone can name a body suggestion too (s2 numbers both from 0).
 */
function outsideRoom(
  batch: OrderedTranslateBatch,
  ids: readonly string[]
): Map<string, OutsideMark> {
  const out = new Map<string, OutsideMark>();
  const shown = batch.shown();
  if (shown === null || ids.length === 0) return out;
  const wanted = new Set(ids);
  let records: Map<string, string> | null = null;
  for (const [name, part] of shown.parts) {
    if (batch.doc.stories.has(name)) {
      continue;
    }

    for (const node of elementsOf(part.root)) {
      if (!wanted.has(node.id)) {
        continue;
      }

      records ??= partRecordIndex(batch.doc);
      out.set(node.id, {
        part: name,
        suggestion: records.get(partRecordKeyOf(name, node)) ?? null,
      });
    }
  }

  return out;
}

/** A change record's key outside the stories: its part, element, `w:id`, author and date. */
function partRecordKey(part: string, element: string, revision: RevisionTriple): string {
  return `${part}\u0000${revisionGroupKey(revision, element)}`;
}

/** The key of a change record the shown package holds outside the stories. */
function partRecordKeyOf(part: string, node: OoxmlElement): string {
  return partRecordKey(
    part,
    node.localName,
    tripleOf(
      attributeOf(node, 'id') ?? '',
      attributeOf(node, 'author') ?? '',
      attributeOf(node, 'date')
    )
  );
}

/**
 * The suggestions of the change records outside the stories, by `partRecordKey`: a pending
 * record shows with its suggestion's author and date, also after a change gave it another
 * (`withDecidedPartRevisions`).
 */
function partRecordIndex(doc: OrderedDocument): Map<string, string> {
  const index = new Map<string, string>();
  for (const [id, { part, element, revision }] of doc.partRevisions) {
    const attribution = doc.suggestions.get(id)?.attribution;
    if (attribution !== undefined) {
      index.set(
        partRecordKey(part, element, tripleOf(revision.id, attribution.author, attribution.date)),
        id
      );
    }
  }

  return index;
}

/** The revisions one decision op names: its triple, or the triples of the marks it names. */
function triplesOf(
  op: DecisionOp,
  batch: OrderedTranslateBatch,
  outside: ReadonlyMap<string, OutsideMark>
): readonly RevisionTriple[] {
  if (op.op === 'acceptRevision' || op.op === 'rejectRevision') {
    // A card outside the stories is named by its marks, never by its three values
    if (op.siteNodeIds?.some((id) => outside.has(id))) {
      return [];
    }

    return [tripleOf(op.revision.id, op.revision.author, op.revision.date)];
  }

  // The engine's planner names every mark it resolves; a part-wide op without marks names
  // a story the gate cannot see
  if (op.siteNodeIds === undefined)
    throw new OrderedClientRefusal('ordered-unsupported-op', `${op.op}.without-sites`);
  return triplesOfSites(
    batch,
    op.siteNodeIds.filter((id) => !outside.has(id))
  );
}

/**
 * The cards of `suggestions`, each once, as this copy shows them now. A decided suggestion, and
 * a card named before (by an earlier mark, or for the other status), are passed over.
 */
function seenCards(
  doc: OrderedDocument,
  suggestions: Iterable<string>,
  named: Set<string>
): OrderedSeenCard[] {
  const cards: OrderedSeenCard[] = [];
  for (const suggestion of suggestions) {
    if (named.has(suggestion)) continue;
    const seen = orderedCardSeen(doc, suggestion);
    if (seen === null) continue;
    for (const entry of seen) named.add(entry.suggestion);
    cards.push({ suggestions: seen });
  }

  return cards;
}

/**
 * The cards some triples reach, each once, as this copy shows them now, and how many triples
 * name marks on some runs of an equation only: the list keeps those as file content, so they
 * name no suggestion.
 */
function cardsOf(
  batch: OrderedTranslateBatch,
  index: TripleIndex,
  triples: readonly RevisionTriple[],
  named: Set<string>
): { readonly cards: OrderedSeenCard[]; readonly equationParts: number } {
  const cards: OrderedSeenCard[] = [];
  let equationParts = 0;
  for (const triple of triples) {
    const found = index.suggestions(triple);
    if (found.length === 0) {
      if (!equationPartShown(batch, triple)) {
        throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no such suggestion');
      }

      equationParts += 1;
    } else {
      for (const card of seenCards(batch.doc, found, named)) {
        cards.push(card);
      }
    }
  }

  return { cards, equationParts };
}

/**
 * The sender's own suggestions among `suggestions` that are replies on another author's card: a
 * change the sender made inside that author's suggestion (the organiser's words in a party's
 * insertion). The editor and Word reject such a change alone, and deciding its card would reject
 * the other author's words too, so a reject of it takes it back (`retract`), which never touches
 * another person's text. A card the sender leads is decided whole, as any card.
 */
function ownReplies(
  doc: OrderedDocument,
  maker: OrderedMaker,
  suggestions: Iterable<string>
): string[] {
  const replies: string[] = [];
  for (const suggestion of suggestions) {
    if (!madeBy(doc, suggestion, maker)) {
      continue;
    }

    const card = orderedCardOf(doc, suggestion);
    if (card !== null && !madeBy(doc, card.id, maker)) {
      replies.push(suggestion);
    }
  }

  return replies;
}

/**
 * Say by name what a decision leaves as it is, without refusing it; refuse the write by that
 * name when it decides nothing else.
 */
function leaveOrRefuse(
  batch: OrderedTranslateBatch,
  left: OrderedClientRefusal,
  decides: boolean
): void {
  if (!decides) throw left;
  batch.leave(left.report());
}

/**
 * What one status of a decision names: revisions by their three values, and change records
 * outside the stories by their suggestions.
 */
interface Wanted {
  readonly triples: RevisionTriple[];
  readonly records: Set<string>;
}

interface WantedByStatus {
  readonly accepted: Wanted;
  readonly rejected: Wanted;
}

/**
 * What each status of a batch's decision ops names, and the revisions rejected one at a time: a
 * reply of the sender's among them is taken back.
 */
function wantedOf(
  ops: readonly DecisionOp[],
  batch: OrderedTranslateBatch,
  outside: ReadonlyMap<string, OutsideMark>
): { readonly wanted: WantedByStatus; readonly rejectedAlone: RevisionTriple[] } {
  const wanted: WantedByStatus = {
    accepted: { triples: [], records: new Set<string>() },
    rejected: { triples: [], records: new Set<string>() },
  };
  const rejectedAlone: RevisionTriple[] = [];
  for (const each of ops) {
    const into =
      each.op === 'acceptRevision' || each.op === 'acceptAllRevisions'
        ? wanted.accepted
        : wanted.rejected;
    const triples = triplesOf(each, batch, outside);
    into.triples.push(...triples);
    if (each.op === 'rejectRevision') {
      rejectedAlone.push(...triples);
    }

    for (const id of each.siteNodeIds ?? []) {
      const suggestion = outside.get(id)?.suggestion;
      if (suggestion) {
        into.records.add(suggestion);
      }
    }
  }

  return { wanted, rejectedAlone };
}

/**
 * A change record the list lacks stays as it is, said by name; the cards the room holds are
 * decided. The write is refused when it decides nothing else.
 */
function leaveLacking(
  batch: OrderedTranslateBatch,
  outside: ReadonlyMap<string, OutsideMark>,
  wanted: WantedByStatus
): void {
  const lacking = [...outside.values()].filter((mark) => mark.suggestion === null);
  if (lacking.length === 0) {
    return;
  }

  const parts = [...new Set(lacking.map((mark) => mark.part))].join(', ');
  const decides = Object.values(wanted).some(
    ({ triples, records }) => triples.length > 0 || records.size > 0
  );
  leaveOrRefuse(
    batch,
    new OrderedClientRefusal(
      'ordered-refused',
      'not-in-room',
      `${lacking.length} marks in ${parts}`.slice(0, 200)
    ),
    decides
  );
}

/**
 * The batch's decision, read from the list when its step runs: a `retract` of the sender's own
 * replies rejected alone, then one `decide` for each status that reaches a card.
 */
function decisionChanges(
  batch: OrderedTranslateBatch,
  wanted: WantedByStatus,
  rejectedAlone: readonly RevisionTriple[]
): OrderedChangeBody[] {
  const index = new TripleIndex(batch.doc);
  const named = new Set<string>();
  const out: OrderedChangeBody[] = [];
  const replies = ownReplies(
    batch.doc,
    makerOf(batch.person),
    rejectedAlone.flatMap((triple) => index.suggestions(triple))
  );
  if (replies.length > 0) {
    for (const reply of replies) {
      named.add(reply);
    }

    out.push({ kind: 'retract', suggestions: replies });
  }

  let equationParts = 0;
  for (const status of ['accepted', 'rejected'] as const) {
    const reached = cardsOf(batch, index, wanted[status].triples, named);
    equationParts += reached.equationParts;
    const cards = [...reached.cards, ...seenCards(batch.doc, wanted[status].records, named)];
    if (cards.length > 0) {
      out.push({ kind: 'decide', cards, status });
    }
  }

  // Marks on some runs of an equation only stay file content: the list has no unit there
  if (equationParts > 0) {
    leaveOrRefuse(
      batch,
      new OrderedClientRefusal('ordered-unsupported-op', 'acceptRevision.equation-part'),
      out.length > 0
    );
  }

  return out;
}

/**
 * Every decision op of a batch is one decision: one `decide` change that names every card the
 * ops' revisions reach, as this copy shows it now (the seen check, design D6), so the room leaves
 * a card that changed before it arrives. An Accept All, a Reject All or a choice of cards is one
 * change, so the room orders one change and every copy lays its pages out once. The first
 * decision op makes it, from the revisions every decision op names, read before it applies; the
 * others make nothing. A batch that accepts some revisions and rejects others makes one for
 * each. A reject of one revision that is the sender's own reply on another author's card is one
 * `retract` of it before them (`ownReplies`).
 */
const decision = (kind: DecisionKind) =>
  translator(kind, (op, batch) => {
    const ops = batch.ops.filter(isDecisionOp);
    if (ops[0] !== op) {
      return () => [];
    }

    const outside = outsideRoom(
      batch,
      ops.flatMap((each) => each.siteNodeIds ?? [])
    );
    const { wanted, rejectedAlone } = wantedOf(ops, batch, outside);
    leaveLacking(batch, outside, wanted);
    return () => decisionChanges(batch, wanted, rejectedAlone);
  });

/**
 * A new comment's markers, as one `comment` change with the comment's text: the start op makes
 * it, from its place to the end marker's; the end and the reference are its other half, and
 * make nothing of their own.
 */
const comment = translator('insertCommentMarker', (op, batch) => {
  const text = batch.review?.commentText;
  const half = (side: 'start' | 'end') =>
    batch.ops.find(
      (other) =>
        other.op === 'insertCommentMarker' &&
        other.commentId === op.commentId &&
        other.marker === side
    );
  const start = half('start');
  const end = half('end');
  if (
    text === undefined ||
    start?.op !== 'insertCommentMarker' ||
    end?.op !== 'insertCommentMarker'
  ) {
    throw new OrderedClientRefusal(
      'ordered-unsupported-op',
      'insertCommentMarker',
      'a comment marker without its text or its other half'
    );
  }

  if (op.marker !== 'start') {
    return () => [];
  }

  return () => {
    const from = batch.paragraph(start.paragraphId);
    const to = batch.paragraph(end.paragraphId);
    if (from.story !== to.story) {
      throw new OrderedClientRefusal(
        'ordered-refused',
        'out-of-step',
        'a comment over two stories'
      );
    }

    return [
      {
        kind: 'comment',
        part: from.story.name,
        from: placeAt(from, start.offset),
        to: placeAt(to, end.offset),
        text,
      },
    ];
  };
});

const DECIDE_NOTE =
  "With the batch's other decision ops, one `decide` that names each card their revisions " +
  'reach, with the versions this copy shows.';

export const REVIEW_OPS: OrderedOpTable<
  | 'acceptRevision'
  | 'rejectRevision'
  | 'acceptAllRevisions'
  | 'rejectAllRevisions'
  | 'insertCommentMarker'
> = {
  acceptRevision: {
    status: 'translated',
    group: '9.3',
    changes: ['decide'],
    party: DECIDE_RULE,
    tracking: 'review',
    note: `Accept one revision. ${DECIDE_NOTE}`,
    translate: decision('acceptRevision'),
  },
  rejectRevision: {
    status: 'translated',
    group: '9.3',
    changes: ['decide', 'retract'],
    party: DECIDE_RULE,
    tracking: 'review',
    note: `Reject one revision. ${DECIDE_NOTE} The sender's own reply on another author's card: one \`retract\` of it.`,
    translate: decision('rejectRevision'),
  },
  acceptAllRevisions: {
    status: 'translated',
    group: '9.3',
    changes: ['decide'],
    party: DECIDE_RULE,
    tracking: 'review',
    note: `Accept all, or a card the planner resolves by its marks. ${DECIDE_NOTE}`,
    translate: decision('acceptAllRevisions'),
  },
  rejectAllRevisions: {
    status: 'translated',
    group: '9.3',
    changes: ['decide'],
    party: DECIDE_RULE,
    tracking: 'review',
    note: `Reject all, or a card the planner resolves by its marks. ${DECIDE_NOTE}`,
    translate: decision('rejectAllRevisions'),
  },
  insertCommentMarker: {
    status: 'translated',
    group: '9.3',
    changes: ['comment'],
    party: null,
    tracking: 'review',
    note: "A new comment's markers: one `comment` change over its range, with its text.",
    translate: comment,
  },
};
