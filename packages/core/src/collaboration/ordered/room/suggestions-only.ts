// The suggestions-only policy (design D9): the host's rule for a negotiation room, as checks on
// one change, the items it touches, and what the change did once applied. The host's own check
// allows a party's change when taking back that party's own suggestions gives the same document
// before and after it. This policy gives the same verdict per change, except where that check
// misreads the Word form (design D9 lists them).

import { promptControlOf, promptShows } from '../compose/paragraph.ts';
import {
  isHidden,
  isPending,
  itemsAddedBy,
  madeBy,
  pendingBy,
  pendingByOther,
  pendingDeleters,
} from '../model/document.ts';
import { threadRootOf } from '../model/comments.ts';
import { makerOf, sameMaker, type OrderedMaker } from '../model/person.ts';
import { textBoxHostOf } from '../model/text-boxes.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import { orderedKindSpecOf } from '../change/kinds/registry.ts';
import type { OrderedKindSpec } from '../change/kinds/spec.ts';
import type { OrderedChangeOutcome } from '../change/apply.ts';
import type { OrderedCaret, OrderedTouched } from '../change/targets.ts';
import { fillsOthersPrompt } from './over-prompt.ts';
import {
  isOrderedCommentChange,
  isOrderedEditChange,
  type OrderedChange,
  type OrderedCommentThreadChange,
  type OrderedEditChange,
  type OrderedRetractChange,
} from '../change/types.ts';
import type {
  OrderedActor,
  OrderedPolicyCode,
  OrderedPolicyInput,
  OrderedPolicyReason,
  OrderedPolicyVerdict,
  OrderedRoomPolicy,
} from './policy.ts';

/** The most comments a document holds, the Word file's counted. */
const MAX_COMMENTS = 5000;

const refuse = (
  code: OrderedPolicyCode,
  reason?: OrderedPolicyReason,
  detail?: string
): OrderedPolicyVerdict => ({
  ok: false,
  code,
  ...(reason === undefined ? {} : { reason }),
  ...(detail === undefined ? {} : { detail }),
});

const ALLOW: OrderedPolicyVerdict = Object.freeze({ ok: true });

/** Whether an actor is an organiser: it decides, and does not only suggest. */
const isOrganiser = (actor: OrderedActor): boolean => actor.decides && !actor.suggests;

/** Edits that change the text a reader sees; the others change only properties. */
const TEXT_KINDS: ReadonlySet<OrderedChange['kind']> = new Set([
  'insert',
  'insert-tab',
  'insert-break',
  'delete',
  'split',
  'join',
  'move',
  'insert-note',
  'insert-table',
  'insert-row',
  'delete-row',
  'insert-column',
  'delete-column',
  'control-value',
]);

/** Edits that change only properties: a format or a property change. */
const PROPERTY_KINDS: ReadonlySet<OrderedChange['kind']> = new Set([
  'format',
  'mark-format',
  'paragraph-property',
  'table-property',
]);

/** Edits that add a container beside their target. */
const BESIDE_KINDS: ReadonlySet<OrderedChange['kind']> = new Set(['insert-row', 'insert-column']);

/** Changes that strike what they touch, and leave units struck already as they are (policy 1). */
const STRIKES: ReadonlySet<OrderedChange['kind']> = new Set([
  'delete',
  'join',
  'delete-row',
  'delete-column',
  'control-value',
  'link-target',
]);

/** What a kind of edit does, as the party rules read it. */
interface PolicyFlags {
  /** The spec of a kind a group added, which says what the kind does. */
  readonly spec: OrderedKindSpec | undefined;
  /** It changes the text a reader sees; otherwise only properties or the package. */
  readonly text: boolean;
  /** It sets properties: a format or a property change. */
  readonly properties: boolean;
  /** It adds a row or a column beside its target. */
  readonly beside: boolean;
  /** It strikes what it touches. */
  readonly strikes: boolean;
}

/** What an edit does: a group's kind says so in its spec (`change/kinds/spec.ts`). */
function policyFlags(change: OrderedEditChange): PolicyFlags {
  const spec = orderedKindSpecOf(change.kind);
  return {
    spec,
    text: spec?.policy.text ?? TEXT_KINDS.has(change.kind),
    properties: spec?.policy.properties ?? PROPERTY_KINDS.has(change.kind),
    beside: spec?.policy.beside ?? BESIDE_KINDS.has(change.kind),
    strikes: spec?.policy.strikes ?? STRIKES.has(change.kind),
  };
}

/** The host's reason for an edit that is not a suggestion: the text changed, or the package. */
const directReason = (flags: PolicyFlags): OrderedPolicyReason => (flags.text ? 'text' : 'package');

/** A header's or a footer's story, or a text box's in one. */
function isFurniture(doc: OrderedDocument, story: OrderedStory | null): boolean {
  if (story === null) return false;
  const host = textBoxHostOf(doc, story.name);
  const part = (host === null ? story : doc.stories.get(host.host))?.part;
  return part?.root.localName === 'hdr' || part?.root.localName === 'ftr';
}

/**
 * Other people's pending suggestions that inserted an item or a container around it. With no
 * maker, everyone's: a direct edit is no one's suggestion, so every pending one is another's.
 */
function insertedByOthers(
  doc: OrderedDocument,
  maker: OrderedMaker | null,
  item: OrderedItem
): string[] {
  const other = (id: string) =>
    maker === null ? isPending(doc, id) : pendingByOther(doc, id, maker);
  const out: string[] = [];
  if (item.by !== null && other(item.by)) {
    out.push(item.by);
  }

  for (const id of item.within.by) {
    if (other(id)) {
      out.push(id);
    }
  }

  return out;
}

/**
 * Whether another person's pending deletion holds the unit in Word: their `w:del` around its run.
 * A row or a cell they struck marks only the row or the cell (`within.dels`), so a format on its
 * text sits beside their deletion.
 */
function struckByOthers(doc: OrderedDocument, maker: OrderedMaker, item: OrderedItem): boolean {
  for (const id of item.dels) {
    if (pendingByOther(doc, id, maker)) {
      return true;
    }
  }

  return false;
}

/** The shown cells of the row a cell is in, not those of a table nested in them. */
function rowCellsOf(doc: OrderedDocument, story: OrderedStory, cell: OrderedItem): OrderedItem[] {
  let row: OrderedItem | null = null;
  let depth = 0;
  for (let at = story.tree.prev(cell); at !== null && row === null; at = story.tree.prev(at)) {
    if (at.token === 'tbl') depth += at.kind === 'close' ? 1 : -1;
    else if (depth === 0 && at.kind === 'open' && at.token === 'tr') row = at;
  }

  return row === null ? [cell] : cellsOfRow(doc, story, row);
}

/** The shown cells of a row, not those of a table nested in them. */
function cellsOfRow(doc: OrderedDocument, story: OrderedStory, row: OrderedItem): OrderedItem[] {
  const cells: OrderedItem[] = [];
  let tables = 0;
  for (let at = story.tree.next(row); at !== null; at = story.tree.next(at)) {
    if (at.kind === 'close' && at.pair === row.pair) break;
    if (at.token === 'tbl') tables += at.kind === 'open' ? 1 : -1;
    else if (tables === 0 && at.kind === 'open' && at.token === 'tc' && !isHidden(doc, at))
      cells.push(at);
  }

  return cells;
}

/**
 * The cell a new column takes its width from, and its index in its row. A tracked column takes it
 * from the cell before it (the cell after it at the row's start), the cell a decision gives it
 * back to. A named cell that is not among its row's shown cells has index -1 and lends itself.
 */
function lenderOf(
  doc: OrderedDocument,
  story: OrderedStory,
  cell: OrderedItem,
  where: 'left' | 'right'
): { readonly index: number; readonly cell: OrderedItem | undefined } {
  const cells = rowCellsOf(doc, story, cell);
  const at = cells.indexOf(cell);
  const index = where === 'right' || at <= 0 ? at : at - 1;
  return { index, cell: index < 0 ? cell : cells[index] };
}

/**
 * Whether a new column changes another person's pending row or cell record in `row`: the cell it
 * takes its width from lends from that record too, and a row's formatting record goes on the
 * new cell. The host's check refuses it: rejecting the party's column alone no longer gives
 * the table back as it was (the record keeps the lent width, as Word's reject keeps it), and
 * the record on the new cell ties the other person's change to the column.
 */
function columnChangesOthersRecords(
  doc: OrderedDocument,
  maker: OrderedMaker,
  story: OrderedStory,
  row: OrderedItem,
  lender: number
): boolean {
  const records = (item: OrderedItem | undefined) =>
    item?.props.some(
      (r) =>
        (r.kind !== 'element' || r.untracked !== true) && pendingByOther(doc, r.suggestion, maker)
    ) === true;
  return records(row) || records(cellsOfRow(doc, story, row)[lender]);
}

/** The shown mark that ends the paragraph an item is in: the one Enter there copies. */
function paragraphMarkOf(
  doc: OrderedDocument,
  story: OrderedStory,
  from: OrderedItem
): OrderedItem | null {
  for (let at: OrderedItem | null = from; at !== null; at = story.tree.next(at))
    if (at.kind === 'mark' && !isHidden(doc, at)) return at;
  return null;
}

/**
 * Whether Enter at a caret gives its new paragraph another person's pending format or property
 * record: the new mark copies the mark it splits with its pending records (as Word copies a
 * pending paragraph change), so it would write under their name and join their card.
 */
function copiesOthersRecords(
  doc: OrderedDocument,
  maker: OrderedMaker,
  story: OrderedStory,
  caret: OrderedCaret
): boolean {
  const mark = paragraphMarkOf(doc, story, caret.at);
  if (mark === null) return false;
  return [...mark.fmts, ...mark.props].some((r) => pendingByOther(doc, r.suggestion, maker));
}

/**
 * Another person's pending numbering record (a file's own `w:numPr/w:ins` or `w:numberingChange`,
 * `load/properties.ts`) on a paragraph that a list change reaches, or null. The change would drop
 * the record or copy it into the party's own record, so the other person's suggestion would go
 * without a decision: refused as `others_suggestions`, as a page setup over another person's
 * pending section record is.
 */
function othersNumbering(
  doc: OrderedDocument,
  maker: OrderedMaker,
  change: OrderedEditChange,
  touched: OrderedTouched
): string | null {
  const list =
    change.kind === 'new-list' ||
    change.kind === 'list-level' ||
    (change.kind === 'paragraph-property' &&
      change.properties.some((p) => p.localName === 'numPr'));
  if (!list) {
    return null;
  }

  for (const item of touched.items) {
    for (const record of item.kind === 'mark' ? item.props : []) {
      if (record.target === 'numPr' && pendingByOther(doc, record.suggestion, maker)) {
        return record.suggestion;
      }
    }
  }

  return null;
}

/** Whether the author's own pending suggestion inserted a shown item. */
function insertedByOwn(doc: OrderedDocument, maker: OrderedMaker, item: OrderedItem): boolean {
  if (item.by === null || isHidden(doc, item)) return false;
  return pendingBy(doc, item.by, maker) !== null;
}

/**
 * Another person's pending suggestion on a shown item that the author's own pending suggestion
 * inserted, or null: their deletion over it, their format on it, or their property change on it.
 * Taking the item back takes their record with it, so their suggestion would shrink or go
 * without a decision. The host's check keeps such a record as it stood when it rejects the
 * author's own suggestions, so it finds the record changed and refuses (`others_suggestions`).
 */
function othersOnOwn(doc: OrderedDocument, maker: OrderedMaker, item: OrderedItem): string | null {
  if (!insertedByOwn(doc, maker, item)) {
    return null;
  }

  for (const id of item.dels) {
    if (pendingByOther(doc, id, maker)) {
      return id;
    }
  }

  for (const record of [...item.fmts, ...item.props]) {
    if (pendingByOther(doc, record.suggestion, maker)) {
      return record.suggestion;
    }
  }

  return null;
}

/** Another person's pending format on a unit, or null: Word nests it in the unit's run. */
function othersFormatOn(
  doc: OrderedDocument,
  maker: OrderedMaker,
  item: OrderedItem
): string | null {
  if (!isUnit(item)) {
    return null;
  }

  for (const record of item.fmts) {
    if (pendingByOther(doc, record.suggestion, maker)) {
      return record.suggestion;
    }
  }

  return null;
}

/**
 * Another person's pending suggestion that a strike of a shown item would change, or null: on
 * the author's own items, as `othersOnOwn` reads them; on other units not struck already (policy
 * 1), their format. Word nests a run's format record in the run's deletion, so taking the
 * deletion back takes the format back too, and today's check finds it changed (`package` or
 * `others_suggestions`). A mark's, a row's or a cell's record sits beside the deletion.
 */
function othersUnderStrike(
  doc: OrderedDocument,
  maker: OrderedMaker,
  item: OrderedItem
): string | null {
  if (insertedByOwn(doc, maker, item)) {
    return othersOnOwn(doc, maker, item);
  }

  if (isHidden(doc, item) || item.dels.some((id) => isPending(doc, id))) {
    return null;
  }

  return othersFormatOn(doc, maker, item);
}

/**
 * Whether Enter at a caret removes a block control's shown prompt for good (policy 4): Word's
 * Return replaces it with no deletion mark, so taking the break back leaves the control empty
 * and today's check finds the prompt's words gone (`text`). A paragraph mark before the caret
 * ends the paragraph before, so only the caret's own paragraph counts.
 */
function breaksBlockPrompt(
  doc: OrderedDocument,
  story: OrderedStory,
  caret: OrderedCaret
): boolean {
  const sides =
    caret.before === null || caret.before.kind === 'mark' ? [caret.at] : [caret.before, caret.at];
  for (const side of sides) {
    const control = promptControlOf(story, side);
    if (control?.token === 'sdtBlock' && promptShows(doc, story, control)) {
      return true;
    }
  }

  return false;
}

/** Whether a row, a cell, a table or a note that `id` inserted holds an item. */
const heldByNew = (item: OrderedItem, id: string): boolean =>
  item.within.by.includes(id) || (item.quiet && item.by === id);

/**
 * Whether a caret's new items go inside another person's pending insertion, where Word nests
 * them in it: between two of its units (one run of theirs), or in a row, a cell, a table or a
 * note it made. A caret is where the change's new items go: past struck text, as the engine
 * types there (`OrderedCaret`). Beside their paragraph mark, a link's or a control's edge, or
 * between their deletion and their insertion, the new run sits beside their suggestions, as the
 * host's check reads the Word file. With no maker, inside anyone's (`insertedByOthers`).
 */
function caretInside(
  doc: OrderedDocument,
  maker: OrderedMaker | null,
  caret: OrderedCaret
): boolean {
  const { before, at } = caret;
  if (before === null) return false;
  const around = new Set(insertedByOthers(doc, maker, at));
  return insertedByOthers(doc, maker, before).some(
    (id) =>
      around.has(id) &&
      ((isUnit(before) && isUnit(at)) || heldByNew(before, id) || heldByNew(at, id))
  );
}

/**
 * Whether a direct edit at a caret types into a pending insertion, the author's own too: inside
 * it, or anywhere in a paragraph of a row, a note or a text box it made. The engine types it
 * there as that insertion's author (`directInsert`), as Word extends a tracked insertion with
 * untracked typing, so the edit makes no suggestion of its own.
 */
function typesIntoInsertion(
  doc: OrderedDocument,
  story: OrderedStory | null,
  caret: OrderedCaret
): boolean {
  if (caretInside(doc, null, caret)) {
    return true;
  }

  const mark = story === null ? null : paragraphMarkOf(doc, story, caret.at);
  return mark !== null && insertedByOthers(doc, null, mark).some((id) => heldByNew(mark, id));
}

/**
 * Whether a caret of a party's edit puts its new items inside another person's suggestion: in
 * their insertion, or, for Enter, in a paragraph whose mark has their pending record.
 */
function caretNested(
  doc: OrderedDocument,
  maker: OrderedMaker,
  change: OrderedEditChange,
  story: OrderedStory | null,
  caret: OrderedCaret
): boolean {
  if (caretInside(doc, maker, caret)) {
    return true;
  }

  return change.kind === 'split' && story !== null && copiesOthersRecords(doc, maker, story, caret);
}

/** A new column's story, and the row index of the cell it takes its width from (-1: none). */
interface NewColumn {
  readonly story: OrderedStory | null;
  readonly lender: number;
}

/**
 * Whether a new row or column at one shown item reaches inside another person's suggestion. It
 * is beside their row or column, not inside, unless the table itself is their pending insertion;
 * a new row copies the shown properties of the row beside it and no change record (the engine's
 * `insertTableRow`). But a column cuts through every row, so a new column (`column`) is inside
 * another person's new row, and inside a row they struck.
 */
function besideNested(
  doc: OrderedDocument,
  maker: OrderedMaker,
  item: OrderedItem,
  column: NewColumn | null
): boolean {
  if (item.within.by.some((id) => pendingByOther(doc, id, maker))) {
    return true;
  }

  if (column === null || item.kind !== 'open' || item.token !== 'tr') {
    return false;
  }

  // A new column puts a cell in every row: in another person's new row it is inside their row,
  // and in a row they struck it is inside their deletion, in every place (the host's check passes
  // only the left edge, as that cell comes before the range)
  if (
    (item.by !== null && pendingByOther(doc, item.by, maker)) ||
    struckByOthers(doc, maker, item)
  ) {
    return true;
  }

  const { story, lender } = column;
  return (
    lender >= 0 && story !== null && columnChangesOthersRecords(doc, maker, story, item, lender)
  );
}

/**
 * Whether a range or a target reaches inside another person's pending suggestion at one shown
 * item, as Word nests the change in it. It is inside when it reaches a unit they inserted; a
 * property change on a mark or a container they inserted sits beside their insertion (Word
 * writes its change record beside the `w:ins`). A format on text or an object they struck is
 * inside their deletion too (Word nests the format in it); a property change on a paragraph
 * mark, a row or a table they struck is not (Word writes its change record beside the deletion
 * mark). A control value or a new link target reaches the units it strikes.
 */
function itemNested(
  doc: OrderedDocument,
  maker: OrderedMaker,
  item: OrderedItem,
  properties: boolean
): boolean {
  if (insertedByOthers(doc, maker, item).length > 0 && (!properties || isUnit(item))) {
    return true;
  }

  return properties && isUnit(item) && struckByOthers(doc, maker, item);
}

/**
 * Whether a party's edit reaches inside another person's pending suggestion. "Inside" is per
 * kind of change, as Word nests the change in the other person's suggestion.
 */
function nested(
  doc: OrderedDocument,
  maker: OrderedMaker,
  change: OrderedEditChange,
  touched: OrderedTouched,
  flags: PolicyFlags
): boolean {
  const story = touched.story;

  // Content in place of a control's prompt drops the prompt's runs, with their suggestions
  if (fillsOthersPrompt(doc, maker, change, touched)) {
    return true;
  }

  if (touched.carets.some((caret) => caretNested(doc, maker, change, story, caret))) {
    return true;
  }

  let column: NewColumn | null = null;
  if (change.kind === 'insert-column') {
    const named = touched.items.find((item) => item.kind === 'open' && item.token === 'tc');
    const lender = story && named ? lenderOf(doc, story, named, change.where) : null;

    // A new column that takes its width from another person's new cell changes that cell, under
    // their card: read in the named cell's row, as their new column has a cell in every row
    const by = lender?.cell?.by ?? null;
    if (by !== null && pendingByOther(doc, by, maker)) {
      return true;
    }

    column = { story, lender: lender?.index ?? -1 };
  }

  for (const item of touched.items) {
    // A deletion leaves what is struck already as it is (policy 1)
    if (isHidden(doc, item) || (flags.strikes && pendingDeleters(doc, item).length > 0)) {
      continue;
    }

    const inside = flags.beside
      ? besideNested(doc, maker, item, column)
      : itemNested(doc, maker, item, flags.properties);
    if (inside) {
      return true;
    }
  }

  return false;
}

/** Whether a suggestion is the actor's own and pending, another person's, or decided. */
function ownershipOf(
  doc: OrderedDocument,
  maker: OrderedMaker,
  id: string
): 'own' | 'other' | 'decided' {
  const suggestion = doc.suggestions.get(id);
  if (suggestion === undefined || !sameMaker(suggestion.attribution, maker)) {
    return 'other';
  }

  return suggestion.status === 'pending' ? 'own' : 'decided';
}

/** Whether any item a suggestion labels is in a header or a footer. */
function labelsFurniture(doc: OrderedDocument, id: string): boolean {
  const trees = new Set<object>();
  for (const story of doc.stories.values()) if (isFurniture(doc, story)) trees.add(story.tree);
  if (trees.size === 0) return false;
  for (const item of doc.labelled.get(id) ?? []) {
    const tree = item.leaf?.tree;
    if (tree !== undefined && trees.has(tree)) return true;
  }

  return false;
}

/**
 * Another person's pending suggestion that taking back `id` on one of its items would take back
 * too, or null: their deletion, format or property change on the author's own new item, or
 * their format Word nests in the author's own struck run.
 */
function othersTakenWith(
  doc: OrderedDocument,
  maker: OrderedMaker,
  item: OrderedItem,
  id: string
): string | null {
  if (item.by === id) {
    return othersOnOwn(doc, maker, item);
  }

  if (item.dels.includes(id)) {
    return othersFormatOn(doc, maker, item);
  }

  return null;
}

/**
 * A take-back of the author's own pending suggestions. Taking back their own pending items takes
 * back another person's deletion, format or property change on those items too, and taking back
 * a deletion takes back the formats Word nests in the deleted runs, so both are refused
 * (`others_suggestions`): the host's check keeps such a record as it stood when it rejects the
 * party's own suggestions.
 */
function checkRetract(
  input: OrderedPolicyInput,
  { suggestions, undo = [] }: OrderedRetractChange
): OrderedPolicyVerdict {
  const { actor, document: doc } = input;
  const maker = makerOf(actor.person);
  for (const id of suggestions) {
    const ownership = ownershipOf(doc, maker, id);
    if (ownership === 'other') return refuse('policy-not-own', 'others_suggestions', id);
    if (ownership === 'decided') return refuse('policy-direct', 'text', `${id} is decided`);
    if (actor.suggests && labelsFurniture(doc, id))
      return refuse('policy-furniture', 'furniture', id);
  }

  if (!actor.suggests) return ALLOW;
  for (const id of suggestions) {
    for (const item of doc.labelled.get(id) ?? []) {
      const other = othersTakenWith(doc, maker, item, id);
      if (other !== null) {
        return refuse('policy-not-own', 'others_suggestions', `${other} is on ${id}`);
      }
    }
  }

  // So does the undo of one's own edit, on the words it takes back
  for (const change of undo)
    for (const item of itemsAddedBy(doc, change)) {
      const other = othersOnOwn(doc, maker, item);
      if (other !== null)
        return refuse('policy-not-own', 'others_suggestions', `${other} is on ${item.by}`);
    }

  return ALLOW;
}

/**
 * A comment: only from a person comments are open to, never in a header or a footer, and at most
 * the document's limit. Only an organiser resolves, and a resolved thread takes no reply. A
 * comment may cover anything, another person's pending suggestion too: it is no edit.
 */
function checkComment(
  input: OrderedPolicyInput,
  change: OrderedCommentThreadChange
): OrderedPolicyVerdict {
  const { actor, document: doc } = input;
  if (!actor.comments) {
    return refuse('policy-comment', 'comments');
  }

  if (change.kind === 'comment' && isFurniture(doc, input.touched.story)) {
    return refuse('policy-furniture', 'furniture');
  }

  if (change.kind === 'comment-resolve' && !isOrganiser(actor)) {
    return refuse('policy-comment', 'resolve');
  }

  if (change.kind !== 'comment' && threadRootOf(doc, change.comment).done) {
    return refuse('policy-comment', 'resolved', change.comment.slice(0, 64));
  }

  // A new thread, a reply or a note adds one comment
  const adds = change.kind !== 'comment-resolve' || change.note !== undefined;
  if (adds && doc.comments.size >= MAX_COMMENTS) {
    return refuse('policy-comment', 'limit', `more than ${MAX_COMMENTS}`);
  }

  return ALLOW;
}

/**
 * The refusal of a party's edit that the host's check makes only when the edit changes
 * something: a direct edit, or one in a header or a footer; null for any other edit. That check
 * compares the document before and after, so an edit that changes nothing passes (a deletion of
 * a prompt, which is no content). The room judges it once applied (`checkApplied`): it changed
 * nothing when it made, extended or took back no suggestion.
 */
function appliedOnlyRefusal(
  input: OrderedPolicyInput,
  change: OrderedEditChange,
  flags: PolicyFlags
): OrderedPolicyVerdict | null {
  if (isFurniture(input.document, input.touched.story)) {
    return refuse('policy-furniture', 'furniture');
  }

  if (change.mode !== 'suggest') {
    return refuse('policy-direct', directReason(flags));
  }

  return null;
}

function checkEdit(input: OrderedPolicyInput, change: OrderedEditChange): OrderedPolicyVerdict {
  const { actor, touched, document: doc } = input;
  const maker = makerOf(actor.person);
  const flags = policyFlags(change);
  if (change.into !== undefined) {
    const ownership = ownershipOf(doc, maker, change.into);
    if (ownership === 'other') {
      return refuse('policy-not-own', 'others_suggestions', change.into);
    }

    // A decided suggestion is the document's now: adding to it would not be a suggestion
    if (ownership === 'decided') {
      return refuse('policy-direct', directReason(flags), `${change.into} is decided`);
    }
  }

  if (!actor.suggests) return ALLOW;

  // A kind that names suggestions, not places, has its own party rule
  const party = flags.spec?.policy.party?.(change, doc, maker);
  if (party) {
    return party;
  }

  const numbering = othersNumbering(doc, maker, change, touched);
  if (numbering !== null)
    return refuse('policy-not-own', 'others_suggestions', `${numbering} holds a numbering record`);

  // A kind that writes a header or a footer always changes one (a watermark)
  if (flags.spec?.policy.furniture?.(change, doc)) {
    return refuse('policy-furniture', 'furniture');
  }

  // Judged once applied (`checkApplied`): allowed when it changed nothing. A direct edit that
  // types into a pending insertion makes no suggestion for that check to see: refused now
  const held = appliedOnlyRefusal(input, change, flags);
  if (held !== null) {
    const into =
      change.mode === 'edit' &&
      touched.carets.some((c) => typesIntoInsertion(doc, touched.story, c));
    return into ? held : ALLOW;
  }

  // A strike takes back what it strikes, with another person's records on it
  if (flags.strikes) {
    for (const item of touched.items) {
      const other = othersUnderStrike(doc, maker, item);
      if (other !== null) {
        return refuse('policy-not-own', 'others_suggestions', `${other} is on the struck items`);
      }
    }
  }

  if (nested(doc, maker, change, touched, flags)) {
    return refuse('policy-nested', 'nested');
  }

  const story = touched.story;
  if (
    change.kind === 'split' &&
    story &&
    touched.carets.some((c) => breaksBlockPrompt(doc, story, c))
  )
    return refuse('policy-direct', 'text', "Enter removes a control's prompt for good");
  return ALLOW;
}

/**
 * A party's edit judged by what it did: one that `appliedOnlyRefusal` names is allowed only when
 * it changed nothing, and a format or a property change may not take over another person's
 * pending record (policy 2 is the organiser's alone): the host's check finds their suggestion
 * gone or changed (`others_suggestions` or `package`).
 */
function checkApplied(
  input: OrderedPolicyInput,
  change: OrderedEditChange,
  outcome: OrderedChangeOutcome
): OrderedPolicyVerdict {
  const held = appliedOnlyRefusal(input, change, policyFlags(change));
  if (held !== null && outcome.suggestions.length > 0) return held;
  const doc = input.document;
  const maker = makerOf(input.actor.person);
  const taken = outcome.replaced.find((r) => !madeBy(doc, r.suggestion, maker));
  if (taken !== undefined)
    return refuse('policy-not-own', 'others_suggestions', `${taken.suggestion} is taken over`);
  return ALLOW;
}

/**
 * The suggestions-only rule of a negotiation room as a per-change policy: the organiser edits and
 * decides, and a party only suggests.
 *
 * - A read-only person sends no change.
 * - A party adds suggestions, and adds to or takes back their own pending ones. Nothing in a
 *   header or a footer (`furniture`), nothing inside another person's pending suggestion
 *   (`nested`), no direct edit (`text` or `package`), no decision, and nothing that adds to or
 *   takes back another person's suggestion (`others_suggestions`).
 *   A direct edit, or an edit in a header or a footer, that makes no suggestion changes nothing,
 *   so it passes, judged once applied. A direct edit that types inside a pending insertion, the
 *   party's own too, joins it with no suggestion of its own: refused before it applies. A format
 *   or property change may not take over another person's pending record (policy 2 is the
 *   organiser's).
 * - An organiser edits directly or suggests anywhere, and decides when the login says so. No one
 *   adds to or takes back another person's suggestion: that would write under their name.
 * - Everyone the login gives comments adds comment threads and replies, on any words (another
 *   person's pending suggestion too), but not in a header or a footer (`furniture`); only an
 *   organiser resolves a thread (`resolve`); a resolved thread takes no reply and no resolve
 *   (`resolved`); a document holds at most 5,000 comments, the Word file's counted (`limit`).
 *   Without comments, every comment is refused (`comments`).
 *
 * @public
 */
export function suggestionsOnlyPolicy(): OrderedRoomPolicy {
  return {
    check(input: OrderedPolicyInput): OrderedPolicyVerdict {
      const { actor, change } = input;
      if (actor.readOnly) {
        return refuse('policy-read-only');
      }

      if (change.kind === 'decide') {
        return isOrganiser(actor) ? ALLOW : refuse('policy-decide');
      }

      if (change.kind === 'retract') {
        return checkRetract(input, change);
      }

      // Undo or redo of a direct edit is a direct edit: a party never makes one
      if (change.kind === 'revert') {
        return actor.suggests ? refuse('policy-direct', 'text', 'revert') : ALLOW;
      }

      if (isOrderedCommentChange(change)) {
        return checkComment(input, change);
      }

      if (isOrderedEditChange(change)) {
        return checkEdit(input, change);
      }

      return refuse('policy-failed', undefined, 'unknown kind');
    },
    checkApplied(input: OrderedPolicyInput, outcome: OrderedChangeOutcome): OrderedPolicyVerdict {
      const { actor, change } = input;
      if (!actor.suggests || !isOrderedEditChange(change)) return ALLOW;
      return checkApplied(input, change, outcome);
    },
  };
}
