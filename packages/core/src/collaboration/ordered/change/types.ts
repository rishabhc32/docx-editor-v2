// The change format, version 1 (design D6). A change is JSON: one edit, decision, take-back or
// comment, named by item addresses that never change during a room session. A client
// sends the client fields; the room adds `n`, `by` and `at` and never accepts them from a client.

import type { OrderedTableOp } from '../edits/tables.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedGroupEditChange } from './kinds/registry.ts';

/** The format version a change carries in `v`. @public */
export const ORDERED_CHANGE_VERSION = 1;

/**
 * The format's caps. Client and server share them: a change one accepts, the other accepts.
 *
 * @public
 */
export const ORDERED_CHANGE_LIMITS = Object.freeze({
  /** Changes in one batch. */
  batch: 512,
  /** UTF-16 units of a space name. */
  space: 1024,
  /** UTF-16 units of text in one change (typed text, a note's text, a control's text). */
  text: 65_536,
  /** Paragraphs, rows, cells or suggestions one change names. */
  targets: 1000,
  /**
   * Suggestions one decision names: every card of an Accept All or a Reject All, in one change.
   * A decision of this many suggestions stays under the client frame
   * (`ORDERED_PROTOCOL_LIMITS.clientFrame`) with minted suggestion ids (at most 80 units).
   */
  decisions: 7500,
  /** Properties in one change. */
  properties: 32,
  /** Attributes of one property or one child of a property. */
  attributes: 16,
  /** UTF-16 units of one attribute value. */
  attributeValue: 256,
  /** Children of one property (tab stops, numbering). */
  children: 64,
  /** UTF-16 units of a link target. */
  href: 2048,
  /** UTF-16 units of a link's bookmark name. */
  anchor: 128,
  /** UTF-16 units of a link's tooltip (the engine's own cap on a link attribute). */
  tooltip: 512,
  /** UTF-16 units of a suggestion id. */
  suggestion: 128,
  /** UTF-16 units of a list item's value. */
  listItem: 1024,
  /** UTF-16 units of the Word name of the `by` person the room stamps. */
  author: 256,
  /** UTF-16 units of a comment's text, or of a resolve's note. */
  comment: 10_000,
  /** UTF-16 units of a comment's `w:id`. */
  commentId: 64,
  /** Rows of a new table. */
  tableRows: 1000,
  /** Columns of a new table (Word's own limit). */
  tableColumns: 63,
  /** Width of one column of a new table, in twips (22 inches). */
  columnWidth: 31_680,
});

/**
 * A unit's address: the space it was inserted in and its offset there. A place `{ space,
 * offset }` is the caret right before that unit.
 *
 * @public
 */
export interface OrderedChangeAddress {
  readonly space: string;
  readonly offset: number;
}

/**
 * The caret at the end of a paragraph, named by the address of the paragraph's mark. It is the
 * place right before the mark; when the mark is hidden (the paragraph was joined), it is where
 * the paragraph's own text ends inside the paragraph that took it.
 *
 * @public
 */
export interface OrderedChangeMarkPlace {
  readonly mark: OrderedChangeAddress;
}

/** A caret: before one unit, or at the end of a paragraph. @public */
export type OrderedChangePlace = OrderedChangeAddress | OrderedChangeMarkPlace;

/**
 * How an edit is made: as a suggestion (`suggest`, tracked) or directly (`edit`, untracked).
 *
 * @public
 */
export type OrderedChangeMode = 'suggest' | 'edit';

/**
 * One property an edit writes: a `w:` element name on the target's allowlist, its `w:`
 * attributes, and for `w:tabs` and `w:numPr` one level of child elements. The format holds no
 * XML: `validate.ts` checks every name and attribute. A run, mark or paragraph format may
 * instead take a property off (`remove`, with no attributes or children), as the engine's
 * property ops do with a property they leave out (Clear All Formatting, a toggle off).
 *
 * @public
 */
export interface OrderedChangeProperty {
  readonly localName: string;
  readonly attributes?: Readonly<Record<string, string>>;
  readonly children?: readonly OrderedChangePropertyChild[];
  readonly remove?: true;
}

/** A child element of a property: a name on the parent's allowlist and its attributes. @public */
export interface OrderedChangePropertyChild {
  readonly localName: string;
  readonly attributes?: Readonly<Record<string, string>>;
}

/** The value a content control takes. The same shapes as the engine's control value. @public */
export type OrderedControlValue =
  | { readonly kind: 'text'; readonly text: string }
  | { readonly kind: 'listItem'; readonly value: string }
  | { readonly kind: 'checkbox'; readonly checked: boolean }
  /** `YYYY-MM-DD`, or a full instant `YYYY-MM-DDTHH:MM:SS(.fff)Z`. */
  | { readonly kind: 'date'; readonly iso: string };

/** Fields every change has. @public */
export interface OrderedChangeBase {
  /** The format version: {@link ORDERED_CHANGE_VERSION}. */
  readonly v: 1;
  /** `<namespace>:<counter>`: the sender's namespace and a counter it never reuses. */
  readonly id: string;
}

/**
 * Fields every edit has. `part` names the story part (`/word/document.xml`). `into` names the
 * sender's own pending suggestion that the edit adds to, so one card holds several edits.
 *
 * @public
 */
export interface OrderedEditBase extends OrderedChangeBase {
  readonly part: string;
  readonly mode: OrderedChangeMode;
  readonly into?: string;
}

/** Type text before a place. @public */
export interface OrderedInsertChange extends OrderedEditBase {
  readonly kind: 'insert';
  readonly place: OrderedChangePlace;
  readonly text: string;
  /**
   * The open item of the inline control the caret is in: at its edge the text goes in it, as
   * the keyboard types on in a control (the engine's `inside`); other text goes beside it.
   */
  readonly inside?: OrderedChangeAddress;
}

/** Delete the units from `from` up to, not including, `to`. @public */
export interface OrderedDeleteChange extends OrderedEditBase {
  readonly kind: 'delete';
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
}

/** Set run properties on the units from `from` up to `to`. @public */
export interface OrderedFormatChange extends OrderedEditBase {
  readonly kind: 'format';
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
  readonly properties: readonly OrderedChangeProperty[];
}

/**
 * Set run properties on paragraph marks, each named by its address: the face of the paragraph's
 * own mark, which its list number takes. The engine's `setParagraphMarkProperties`, which the
 * editor sends when a format covers a paragraph's end (a whole paragraph, the last paragraph of
 * a selection). A `format` over several paragraphs already reaches the marks between.
 *
 * @public
 */
export interface OrderedMarkFormatChange extends OrderedEditBase {
  readonly kind: 'mark-format';
  readonly marks: readonly OrderedChangeAddress[];
  readonly properties: readonly OrderedChangeProperty[];
}

/** Break the paragraph at a place (Enter). @public */
export interface OrderedSplitChange extends OrderedEditBase {
  readonly kind: 'split';
  readonly place: OrderedChangePlace;
}

/** Join a paragraph with the next one: remove the paragraph's mark. @public */
export interface OrderedJoinChange extends OrderedEditBase {
  readonly kind: 'join';
  readonly mark: OrderedChangeAddress;
}

/** Set paragraph properties on paragraphs, each named by its mark. @public */
export interface OrderedParagraphPropertyChange extends OrderedEditBase {
  readonly kind: 'paragraph-property';
  readonly marks: readonly OrderedChangeAddress[];
  readonly properties: readonly OrderedChangeProperty[];
}

/**
 * Set table, row or cell properties with the engine's own table op, as the editor sends it. The
 * containers are named by their open items: tables for `setTableProperties`, rows for
 * `setTableRowHeight`, and cells of one table for the cell ops (a border op reads the cells as
 * one selection, so its outside and inside edges are the selection's).
 *
 * @public
 */
export interface OrderedTablePropertyChange extends OrderedEditBase {
  readonly kind: 'table-property';
  readonly on: readonly OrderedChangeAddress[];
  readonly op: OrderedTableOp;
}

/** Insert a row above or below a row, named by its open item. @public */
export interface OrderedInsertRowChange extends OrderedEditBase {
  readonly kind: 'insert-row';
  readonly row: OrderedChangeAddress;
  readonly where: 'above' | 'below';
}

/** Delete a row, named by its open item. @public */
export interface OrderedDeleteRowChange extends OrderedEditBase {
  readonly kind: 'delete-row';
  readonly row: OrderedChangeAddress;
}

/** Insert a column left or right of a cell's column; the cell is named by its open item. @public */
export interface OrderedInsertColumnChange extends OrderedEditBase {
  readonly kind: 'insert-column';
  readonly cell: OrderedChangeAddress;
  readonly where: 'left' | 'right';
}

/** Delete the column of a cell, named by its open item. @public */
export interface OrderedDeleteColumnChange extends OrderedEditBase {
  readonly kind: 'delete-column';
  readonly cell: OrderedChangeAddress;
}

/** Insert a table at a place. Every column takes `columnWidth` twips. @public */
export interface OrderedInsertTableChange extends OrderedEditBase {
  readonly kind: 'insert-table';
  readonly place: OrderedChangePlace;
  readonly rows: number;
  readonly columns: number;
  readonly columnWidth: number;
}

/** Set a content control's value; the control is named by its open item. @public */
export interface OrderedControlValueChange extends OrderedEditBase {
  readonly kind: 'control-value';
  readonly control: OrderedChangeAddress;
  readonly value: OrderedControlValue;
}

/**
 * Point a link at a new target: an address (`href`) or a bookmark in the document (`anchor`),
 * exactly one of them. The link is named by its open item. `tooltip` replaces its tooltip.
 * `from` and `to` name a part of its words, which becomes a link of its own (a direct edit).
 *
 * @public
 */
export interface OrderedLinkTargetChange extends OrderedEditBase {
  readonly kind: 'link-target';
  readonly link: OrderedChangeAddress;
  readonly href?: string;
  readonly anchor?: string;
  readonly tooltip?: string;
  readonly from?: OrderedChangePlace;
  readonly to?: OrderedChangePlace;
}

/** Move the units from `from` up to `to` to `place`. @public */
export interface OrderedMoveChange extends OrderedEditBase {
  readonly kind: 'move';
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
  readonly place: OrderedChangePlace;
}

/** Type a tab (`w:tab`) before a place, as the engine's Tab key writes it. @public */
export interface OrderedInsertTabChange extends OrderedEditBase {
  readonly kind: 'insert-tab';
  readonly place: OrderedChangePlace;
}

/** What a break breaks: the line, the page, or the column. @public */
export type OrderedBreakKind = 'line' | 'page' | 'column';

/**
 * Type a break (`w:br`) before a place: a line break (Shift+Enter), a page break (Ctrl+Enter) or
 * a column break. A line break may restart below the floating objects on one side or both
 * (`w:clear`).
 *
 * @public
 */
export interface OrderedInsertBreakChange extends OrderedEditBase {
  readonly kind: 'insert-break';
  readonly place: OrderedChangePlace;
  readonly break: OrderedBreakKind;
  readonly clear?: 'left' | 'right' | 'all';
}

/** Insert a footnote or an endnote reference at a place, with the note's text. @public */
export interface OrderedInsertNoteChange extends OrderedEditBase {
  readonly kind: 'insert-note';
  readonly place: OrderedChangePlace;
  readonly note: 'footnote' | 'endnote';
  readonly text: string;
}

/** A suggestion of the card an organiser saw, and its version then. @public */
export interface OrderedSeenSuggestion {
  readonly suggestion: string;
  /** The suggestion's `version` when the organiser saw the card. */
  readonly version: number;
}

/** A review card as the organiser saw it: every suggestion of the card. @public */
export interface OrderedSeenCard {
  readonly suggestions: readonly OrderedSeenSuggestion[];
}

/**
 * Accept or reject review cards, as the organiser saw them: every suggestion of each card with
 * the version the organiser saw. One card, or every card of an Accept All or a Reject All, is
 * one change. Each card is decided on its own: a card that changed since (a version, or the set
 * of its suggestions) stays as it is, and the others are decided. When no card named is as the
 * organiser saw it, the decision is refused with `stale-decision` (`not-pending` when the first
 * card is decided already).
 *
 * @public
 */
export interface OrderedDecideChange extends OrderedChangeBase {
  readonly kind: 'decide';
  readonly cards: readonly OrderedSeenCard[];
  readonly status: 'accepted' | 'rejected';
}

/**
 * Take back the sender's own pending suggestions: they never existed for anyone else. Undo
 * sends it, so a client starts a new suggestion at each undo step. At least one of
 * `suggestions` and `undo` names something.
 *
 * @public
 */
export interface OrderedRetractChange extends OrderedChangeBase {
  readonly kind: 'retract';
  readonly suggestions: readonly string[];
  /**
   * Earlier changes of the sender (its own namespace) whose edits of the sender's own pending
   * words are taken back exactly, as the editor's undo of them: the words they took back show
   * again, the words they typed into those suggestions go, and their formats come off.
   */
  readonly undo?: readonly string[];
}

/**
 * Take back the sender's own direct edits, or make them again (the organiser's undo and redo):
 * the suggestions a direct edit made and accepted at once turn `rejected` (taken back) or
 * `accepted` (made again). Addresses never change, so each unit, mark and record shows exactly
 * as it did: typed text goes, deleted text and another person's struck suggestion show again,
 * Enter joins back, a join splits again, a format goes off. Only the sender's own direct
 * suggestions; never a decided one. A party never sends it.
 *
 * @public
 */
export interface OrderedRevertChange extends OrderedChangeBase {
  readonly kind: 'revert';
  readonly suggestions: readonly string[];
  readonly status: 'accepted' | 'rejected';
}

/**
 * A new comment thread on the units from `from` up to `to` (a caret is allowed), with `text`:
 * plain text, one paragraph for each line (`\n` ends a line).
 *
 * @public
 */
export interface OrderedCommentChange extends OrderedChangeBase {
  readonly kind: 'comment';
  readonly part: string;
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
  readonly text: string;
}

/**
 * A reply with `text` to the thread of comment `comment` (its `w:id`). A reply answers the
 * thread's first comment, as Word's threads are flat, and its range sits on that comment's,
 * after the thread's earlier replies.
 *
 * @public
 */
export interface OrderedCommentReplyChange extends OrderedChangeBase {
  readonly kind: 'comment-reply';
  readonly comment: string;
  readonly text: string;
}

/**
 * Resolve the thread of comment `comment` (its `w:id`): `note`, when given, is added as the
 * sender's last reply, then every comment of the thread is resolved.
 *
 * @public
 */
export interface OrderedCommentResolveChange extends OrderedChangeBase {
  readonly kind: 'comment-resolve';
  readonly comment: string;
  readonly note?: string;
}

/** A change of a comment thread: a new thread, a reply, or a resolve. @public */
export type OrderedCommentThreadChange =
  | OrderedCommentChange
  | OrderedCommentReplyChange
  | OrderedCommentResolveChange;

/**
 * An edit of one of the format's first kinds, which the shared files handle by name.
 *
 * @public
 */
export type OrderedCoreEditChange =
  | OrderedInsertChange
  | OrderedInsertTabChange
  | OrderedInsertBreakChange
  | OrderedDeleteChange
  | OrderedFormatChange
  | OrderedMarkFormatChange
  | OrderedSplitChange
  | OrderedJoinChange
  | OrderedParagraphPropertyChange
  | OrderedTablePropertyChange
  | OrderedInsertRowChange
  | OrderedDeleteRowChange
  | OrderedInsertColumnChange
  | OrderedDeleteColumnChange
  | OrderedInsertTableChange
  | OrderedControlValueChange
  | OrderedLinkTargetChange
  | OrderedMoveChange
  | OrderedInsertNoteChange;

/**
 * An edit of the document's content: one of the first kinds, or a kind a group adds
 * (`change/kinds/`).
 *
 * @public
 */
export type OrderedEditChange = OrderedCoreEditChange | OrderedGroupEditChange;

/** One change as a client sends it. @public */
export type OrderedChange =
  | OrderedEditChange
  | OrderedDecideChange
  | OrderedRetractChange
  | OrderedRevertChange
  | OrderedCommentThreadChange;

/** The kinds of change. @public */
export type OrderedChangeKind = OrderedChange['kind'];

/** The kinds the shared files handle by name: the first edits, decisions and comments. */
export type OrderedCoreChangeKind = Exclude<OrderedChangeKind, OrderedGroupEditChange['kind']>;

/**
 * The format's first kinds, in the order the format lists them. Every kind, the groups' too,
 * is `ORDERED_CHANGE_KINDS` (`kinds/registry.ts`).
 */
export const ORDERED_CORE_CHANGE_KINDS: readonly OrderedCoreChangeKind[] = Object.freeze([
  'insert',
  'insert-tab',
  'insert-break',
  'delete',
  'format',
  'mark-format',
  'split',
  'join',
  'paragraph-property',
  'table-property',
  'insert-row',
  'delete-row',
  'insert-column',
  'delete-column',
  'insert-table',
  'control-value',
  'link-target',
  'move',
  'insert-note',
  'decide',
  'retract',
  'revert',
  'comment',
  'comment-reply',
  'comment-resolve',
]);

/**
 * The fields the room sets on an accepted change. A client never sends them.
 *
 * @public
 */
export interface OrderedServerFields {
  /** The order number: 1 for the first change after the Word file, then one more each. */
  readonly n: number;
  /** The person, from the login: their Word name is written as `w:author`. */
  readonly by: OrderedPerson;
  /** The room's clock, `YYYY-MM-DDTHH:MM:SSZ`, written as `w:date`. */
  readonly at: string;
}

/** A change as the room ordered it. @public */
export type OrderedStampedChange = OrderedChange & OrderedServerFields;

/** Whether a change is a new comment thread, a reply, or a resolve. @public */
export function isOrderedCommentChange(
  change: OrderedChange
): change is OrderedCommentThreadChange {
  return (
    change.kind === 'comment' ||
    change.kind === 'comment-reply' ||
    change.kind === 'comment-resolve'
  );
}

/** Whether a change is an edit of content (it has a part, a mode and maybe `into`). @public */
export function isOrderedEditChange(change: OrderedChange): change is OrderedEditChange {
  return (
    change.kind !== 'decide' &&
    change.kind !== 'retract' &&
    change.kind !== 'revert' &&
    !isOrderedCommentChange(change)
  );
}

const CORE_KINDS: ReadonlySet<string> = new Set(ORDERED_CORE_CHANGE_KINDS);

/** Whether a kind is one of the format's first kinds, handled by name in the shared files. */
export function isOrderedCoreKind(kind: string): kind is OrderedCoreChangeKind {
  return CORE_KINDS.has(kind);
}

/** Whether an edit is of one of the format's first kinds, which the shared files handle by name. */
export function isOrderedCoreEditChange(change: OrderedChange): change is OrderedCoreEditChange {
  return isOrderedEditChange(change) && isOrderedCoreKind(change.kind);
}
