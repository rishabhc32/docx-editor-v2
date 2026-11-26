import type { OrderedItemAddress } from './types.ts';

/**
 * Why the ordered model refused an input or an edit. Stable: a caller branches on it.
 *
 * Names and loads:
 *
 * - `unknown-story`, `unknown-address`, `unknown-suggestion`, `unknown-comment`: a name that is
 *   not in the document.
 * - `other-file`: an address a load of another Word file of the room made (design D8): the
 *   change was made against an older file and never applies to this one.
 * - `not-text`: a split asked of an item that is not text.
 * - `splits-surrogate-pair`: a place between the two UTF-16 units of one character.
 * - `item-in-tree`: an item that is already in a story.
 * - `too-deep`, `too-many-items`: a part past the load limits.
 * - `malformed-story`: a part or story the model cannot read: a document part with no body, a
 *   part that is not a story, a container without its other edge, inline content after the last
 *   paragraph mark, a block without its shell, or a paragraph that went during the edit.
 *
 * Edits and decisions:
 *
 * - `invalid-edit`: an edit whose own fields are out of range (empty text, no rows).
 * - `empty-range`: a range that holds nothing, or ends before it starts.
 * - `not-pending`: a decision or a retraction of a suggestion that is already decided.
 * - `stale-decision`: a decision on a card that changed since the organiser saw it (a
 *   suggestion's version, or the card's set of suggestions).
 * - `not-own-suggestion`: a retraction or a revert of another author's suggestion.
 * - `fixed-author`: an `attribution` change that names another author than a suggestion's own. A
 *   room item's author comes only from its person, and a file item's stays the name Word saved.
 * - `not-direct`: a revert of a suggestion no direct edit made (a pending one, or one a
 *   decision accepted or rejected).
 * - `stale-revert`: a revert of a direct edit that is already taken back (or already made
 *   again).
 * - `no-inverse`: a revert that has no exact inverse: it would leave words before a block with no
 *   paragraph (`others-words`), which another person put in a paragraph the direct edit made.
 * - `break-in-container`: Enter in an inline control's prompt, or where the characters on its two
 *   sides share no container (Enter inside a link or a control splits it).
 * - `single-line-control`: Enter in a control that holds one line (a plain text control without
 *   `w:multiLine`, a list, a date, a check box, a picture), or inside an inline rich text
 *   control, which Word leaves whole.
 * - `paragraphs-in-link`: paragraphs pasted inside a link's text, which Word turns into a
 *   `HYPERLINK` field across them (not written yet).
 * - `no-paragraph-before`: Backspace at the start of a story.
 * - `no-paragraph-after`: a join of the last paragraph of a story with the next one.
 * - `hidden-target`: a change names a paragraph mark, a table, a row, a cell, a control or a
 *   link that no longer shows (it was joined, rejected or taken back), or a place in a text box
 *   or a note that no longer shows, or in one of Word's own notes (a separator or a continuation
 *   notice), which the editor does not show to edit; or a decision names a suggestion only in
 *   such a note.
 * - `not-paragraph-start`: a join at a place that is not the start of a paragraph.
 * - `not-in-table`, `not-in-link`, `unknown-control`: a place outside what the edit changes.
 * - `link-not-plain`: a link target change on a link that holds more than text of one format.
 * - `control-without-properties`: a control value on a control with no `w:sdtPr`.
 * - `move-of-suggestion`, `move-of-container`, `move-of-prompt`, `move-into-itself`: a move of
 *   text another suggestion holds, of a container's edge, of a prompt, or into itself.
 * - `others-in-row`: taking back a row or a note that holds another author's change.
 * - `comment-without-range`: a reply to a thread whose first comment has no whole range in a
 *   story: the reply's range has nowhere to sit.
 * - `note-outside-body`: a note reference outside the main document's body.
 * - `engine-refused`: the engine's own builder refuses the edit on a scratch part (the detail
 *   gives its reason).
 * - `name-taken`: a resource's key or Word name (a part, a relationship id, a list number) is
 *   one the file or another resource already holds (design D14).
 * - `foreign-resource`: an edit uses a resource another person's pending suggestion added.
 * - `missing-blob`: media the room's storage does not hold, or bytes a save does not have.
 * - `not-a-section`: a section change names an item that ends no section now.
 * - `not-one-field`: a table of contents' result names a separator and an end that are not one
 *   TOC field's own pair (an end of a field inside the result, or of a later field), or a
 *   separator inside a link or another inline container, or a result that crosses a
 *   container's edge. No copy could compose the result.
 * - `untracked-only`: a suggestion of a change Word cannot track (a picture's size, crop,
 *   position, wrap, rotation, alt text, locks, link or image), or one that inserts or strikes a
 *   simple field with no result, whose mark the Word file has no place for: only a direct edit
 *   makes it.
 * - `locked`: a content control's lock (`w:lock`, ECMA-376 Part 1 §17.5.2.23) forbids the
 *   change, as the engine's `locked` does: an edit of the content of a `contentLocked` or
 *   `sdtContentLocked` control (a control inside one too), or the removal of an `sdtLocked` or
 *   `sdtContentLocked` control. It binds the organiser and a party alike.
 * - `not-covered`: content or an edit the first release does not cover: a place inside content
 *   kept whole (a text box, a shape), an edit kind the model does not know (a paste with
 *   formatting), a list number the document does not define (a new list), an edit of a table
 *   of contents' generated result, or of a control bound to custom XML.
 *
 * @public
 */
export type OrderedRefusalReason =
  | 'unknown-story'
  | 'unknown-address'
  | 'unknown-suggestion'
  | 'unknown-comment'
  | 'other-file'
  | 'not-text'
  | 'splits-surrogate-pair'
  | 'item-in-tree'
  | 'too-deep'
  | 'too-many-items'
  | 'malformed-story'
  | 'invalid-edit'
  | 'empty-range'
  | 'not-pending'
  | 'stale-decision'
  | 'not-own-suggestion'
  | 'fixed-author'
  | 'not-direct'
  | 'stale-revert'
  | 'no-inverse'
  | 'break-in-container'
  | 'single-line-control'
  | 'paragraphs-in-link'
  | 'no-paragraph-before'
  | 'no-paragraph-after'
  | 'hidden-target'
  | 'not-paragraph-start'
  | 'not-in-table'
  | 'not-in-link'
  | 'unknown-control'
  | 'link-not-plain'
  | 'control-without-properties'
  | 'move-of-suggestion'
  | 'move-of-container'
  | 'move-of-prompt'
  | 'move-into-itself'
  | 'others-in-row'
  | 'comment-without-range'
  | 'note-outside-body'
  | 'engine-refused'
  | 'name-taken'
  | 'foreign-resource'
  | 'missing-blob'
  | 'not-a-section'
  | 'not-one-field'
  | 'untracked-only'
  | 'locked'
  | 'not-covered';

/**
 * An input or an edit the ordered model refuses by name. Nothing changed when it is thrown
 * inside an undo record: the record takes back every write first.
 *
 * @public
 */
export class OrderedRefusal extends Error {
  readonly reason: OrderedRefusalReason;
  readonly detail: string | undefined;

  constructor(reason: OrderedRefusalReason, detail?: string) {
    super(detail === undefined ? reason : `${reason}: ${detail}`);
    this.name = 'OrderedRefusal';
    this.reason = reason;
    this.detail = detail;
  }
}

/** An address as a refusal names it: its space (cut short) and offset. */
export const addressLabel = (address: OrderedItemAddress): string =>
  `${address.space.slice(0, 160)}|${address.offset}`;

/** An error's name for a log line or a report; never its message, which can hold document text. */
export const errorName = (error: unknown): string =>
  error instanceof Error ? error.name : typeof error;

/**
 * The engine's refusal of its own op on the view, by the room's name: a control's lock as
 * `locked`, any other reason as `engine-refused` with the reason in its detail.
 */
export function engineRefusal(op: string, reason: string): OrderedRefusal {
  return reason === 'locked'
    ? new OrderedRefusal('locked', `${op}: locked`)
    : new OrderedRefusal('engine-refused', `${op}: ${reason}`);
}
