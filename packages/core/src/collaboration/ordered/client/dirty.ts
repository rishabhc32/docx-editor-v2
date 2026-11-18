import { orderedKindSpecOf } from '../change/kinds/registry.ts';
import { orderedChangeTargets } from '../change/targets.ts';
import { notesStoryOfReference } from '../compose/note-references.ts';
import { notesStoryOf } from '../edits/notes.ts';
import { replacedPromptPlace } from '../edits/places.ts';
import {
  isOrderedCommentChange,
  isOrderedEditChange,
  type OrderedChange,
  type OrderedCommentChange,
  type OrderedDecideChange,
  type OrderedEditChange,
  type OrderedRetractChange,
  type OrderedRevertChange,
} from '../change/types.ts';
import { commentRangeItems, threadRootOf } from '../model/comments.ts';
import { itemAt, itemsAddedBy, storyOfItem } from '../model/document.ts';
import { takePromptMarkChanges } from '../model/prompt-marks.ts';
import { itemsInside } from '../placement.ts';
import {
  BLOCK_TOKENS,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import { nextShownMark } from './paragraphs.ts';
import type { OrderedDirty } from './publish.ts';

/**
 * Changes that can move block edges: their story is composed whole, with the notes parts of the
 * note references they reach. Other changes name the paragraphs they touched by their marks.
 */
const STRUCTURAL = new Set<OrderedChange['kind']>([
  'insert-table',
  'insert-row',
  'delete-row',
  'insert-column',
  'delete-column',
  'table-property',
  'insert-note',
  'move',
  'control-value',
]);

/**
 * Most items `noteItems` reads one by one before it names whole stories. A take-back reads one
 * budget for its suggestions and one for each own edit it undoes.
 */
const ITEM_BUDGET = 4096;

/** The shown mark that ends the paragraph an item is in, or null at the story's end. */
function markAfter(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem
): OrderedItem | null {
  // In a prompt's paragraph that content replaced: the control's paragraph a place there goes to
  const replaced = replacedPromptPlace(doc, story, item);
  if (replaced) return replaced.mark;
  return nextShownMark(doc, story, item);
}

/** An item whose change can move a block edge. */
const edgeItem = (item: OrderedItem): boolean =>
  ((item.kind === 'open' || item.kind === 'close') &&
    item.token !== undefined &&
    BLOCK_TOKENS.has(item.token)) ||
  (item.kind === 'zero' && item.block === true);

/**
 * Name the paragraph an item is in; a paragraph mark names itself (shown or not, the publisher
 * finds the paragraphs a split or a join made of it); a block edge names its story.
 */
function noteItem(
  doc: OrderedDocument,
  story: OrderedStory,
  item: OrderedItem,
  dirty: OrderedDirty
): void {
  if (item.kind === 'mark') {
    dirty.mark(story.name, { space: item.space, offset: item.offset });
    return;
  }

  // A note shows only while its reference does: its notes part is drawn again with it
  const notes = notesStoryOfReference(doc, item);
  if (notes) dirty.story(notes.name);
  if (edgeItem(item)) {
    dirty.story(story.name);
    return;
  }

  const mark = markAfter(doc, story, item);
  if (mark) dirty.mark(story.name, { space: mark.space, offset: mark.offset });
  else dirty.story(story.name);
}

/**
 * The notes parts of the references a change that names its whole story reaches: in a row or a
 * column it deletes, in a control whose words its value replaces, in words it moves. A note
 * shows only while its reference does, so its notes part is drawn again with the story.
 */
function noteReferencesReached(
  doc: OrderedDocument,
  change: OrderedEditChange,
  dirty: OrderedDirty
): void {
  const targets = orderedChangeTargets(doc, change);
  const story = doc.stories.get(change.part);
  if (!targets.ok || !story) return;
  const reach = (item: OrderedItem): void => {
    const notes = notesStoryOfReference(doc, item);
    if (notes) dirty.story(notes.name);
  };
  for (const item of targets.touched.items) {
    reach(item);

    // The story is composed whole, so reading a container's items costs no more than that
    if (item.kind === 'open') for (const inside of itemsInside(story, item)) reach(inside);
  }
}

/** Whether a suggestion added something outside the story text (design D14). */
function holdsResource(doc: OrderedDocument, id: string): boolean {
  for (const resource of doc.resources.added.values()) if (resource.suggestion === id) return true;
  return false;
}

/** The paragraphs of these items; past the budget, their whole stories. */
function noteItems(
  doc: OrderedDocument,
  groups: readonly Iterable<OrderedItem>[],
  dirty: OrderedDirty
): void {
  let budget = ITEM_BUDGET;
  for (const items of groups)
    for (const item of items) {
      const story = storyOfItem(doc, item);
      if (!story) continue;
      budget -= 1;
      if (budget < 0) {
        dirty.story(story.name);
        continue;
      }

      noteItem(doc, story, item, dirty);
    }
}

/** The paragraphs of the items a decision or a take-back labels. */
function noteSuggestions(doc: OrderedDocument, ids: readonly string[], dirty: OrderedDirty): void {
  // A resource (a setting, a list) or a change record outside the stories (a style's) can show
  // or go with no item: the package goes out
  if (ids.some((id) => holdsResource(doc, id) || doc.partRevisions.has(id))) dirty.stale = true;
  noteItems(
    doc,
    ids.map((id) => doc.labelled.get(id) ?? []),
    dirty
  );
}

/** The paragraphs of a block prompt that a change showed or hid (`change/settle-prompts.ts`). */
function notePromptMarks(doc: OrderedDocument, dirty: OrderedDirty): void {
  for (const mark of takePromptMarkChanges(doc)) {
    const story = storyOfItem(doc, mark);
    if (story) {
      dirty.mark(story.name, { space: mark.space, offset: mark.offset });
    }
  }
}

/** The paragraphs of the suggestions a decision, a take-back or a revert names. */
function noteDecision(
  doc: OrderedDocument,
  change: OrderedDecideChange | OrderedRetractChange | OrderedRevertChange,
  dirty: OrderedDirty
): void {
  if (change.kind === 'decide') {
    noteSuggestions(
      doc,
      change.cards.flatMap((card) => card.suggestions.map((seen) => seen.suggestion)),
      dirty
    );
    return;
  }

  noteSuggestions(doc, change.suggestions, dirty);
  if (change.kind === 'retract') {
    // The words an undo of one's own edits changes: read before the write, which forgets them
    for (const id of change.undo ?? []) {
      noteItems(doc, [doc.ownEdited.get(id) ?? [], itemsAddedBy(doc, id)], dirty);
    }
  }
}

/**
 * A change that added resources outside the story text publishes the package with them, and
 * the stories it added as they are now: applied again after another person's change, a copy
 * holds what it copied then. Its keys come from the change id (`p:<change>.<k>`, in order).
 */
function noteAddedResources(
  doc: OrderedDocument,
  change: OrderedChange,
  dirty: OrderedDirty
): void {
  for (let k = 1; ; k += 1) {
    const resource = doc.resources.added.get(`p:${change.id}.${k}`);
    if (resource === undefined) {
      break;
    }

    dirty.stale = true;
    if (resource.kind === 'story' && doc.stories.has(resource.part)) {
      dirty.story(resource.part);
    }
  }
}

/** The ids of a change's own suggestions, read after the write: they come from the change id. */
function ownSuggestions(doc: OrderedDocument, change: OrderedChange): string[] {
  const own: string[] = [];
  for (let k = 1; doc.suggestions.has(`S:${change.id}.${k}`); k += 1) {
    own.push(`S:${change.id}.${k}`);
  }

  return own;
}

/** A change that can move a block edge: its whole story, and the notes parts it reaches. */
function noteStructural(
  doc: OrderedDocument,
  change: OrderedEditChange,
  dirty: OrderedDirty
): void {
  dirty.story(change.part);

  // A new note is drawn in its notes part too, which the change makes when the file has none
  const notes = change.kind === 'insert-note' ? notesStoryOf(doc, change.note) : null;
  if (notes) {
    dirty.story(notes.name);
  }

  noteReferencesReached(doc, change, dirty);
}

/** The paragraphs of the items and carets a change touches; past the budget, its story. */
function noteTouched(
  doc: OrderedDocument,
  change: OrderedEditChange | OrderedCommentChange,
  dirty: OrderedDirty
): void {
  const targets = orderedChangeTargets(doc, change);
  const story = doc.stories.get(change.part);
  if (!targets.ok || !story) {
    // A change whose targets do not resolve writes nothing; its story is drawn again to be sure
    dirty.story(change.part);
    return;
  }

  const { items, carets } = targets.touched;
  if (items.length > ITEM_BUDGET) {
    dirty.story(change.part);
    return;
  }

  for (const item of items) {
    noteItem(doc, story, item, dirty);
  }

  // Typed units go before the caret's unit, or before the mark at a paragraph's end
  for (const { at } of carets) {
    noteItem(doc, story, at, dirty);
  }

  // Enter's new mark ends the first half. Its space comes from the change id (`k:<change>.1`,
  // `orderedIdMinter`): read after the write, it names the first half too
  if (change.kind === 'split') {
    const mark = itemAt(story, { space: `k:${change.id}.1`, offset: 0 });
    if (mark && mark.kind === 'mark') {
      noteItem(doc, story, mark, dirty);
    }
  }
}

/** The paragraphs of a thread's range, where a reply's markers go: its start, end and reference. */
function noteThread(doc: OrderedDocument, comment: string, dirty: OrderedDirty): void {
  if (!doc.comments.has(comment)) {
    return;
  }

  const range = commentRangeItems(doc, threadRootOf(doc, comment).id);
  if (range === null) {
    return;
  }

  for (const item of [range.start, range.end, range.reference]) {
    if (item !== null) {
      noteItem(doc, range.story, item, dirty);
    }
  }
}

/**
 * Add to `dirty` what a change writes, read from the document as it is now. Call it before the
 * write (what it will change) and after it (what it changed), so a take-back, which removes
 * labels, still names the paragraphs it leaves.
 */
export function noteDirty(doc: OrderedDocument, change: OrderedChange, dirty: OrderedDirty): void {
  notePromptMarks(doc, dirty);
  if (change.kind === 'decide' || change.kind === 'retract' || change.kind === 'revert') {
    noteDecision(doc, change, dirty);
    return;
  }

  if (isOrderedCommentChange(change)) {
    dirty.comments = true;
    if (change.kind === 'comment') {
      noteTouched(doc, change, dirty);
    } else {
      noteThread(doc, change.comment, dirty);
    }

    return;
  }

  const spec = orderedKindSpecOf(change.kind);

  // A change that rewrites suggestions in place (their attribution): their paragraphs
  const rewritten = isOrderedEditChange(change) ? spec?.rewrites?.(change) : undefined;
  if (rewritten !== undefined) {
    noteSuggestions(doc, rewritten, dirty);
    return;
  }

  if (!doc.stories.has(change.part)) {
    return;
  }

  noteAddedResources(doc, change, dirty);
  if (spec?.publish === 'suggestions') {
    noteSuggestions(doc, ownSuggestions(doc, change), dirty);
    return;
  }

  if (isOrderedEditChange(change)) {
    for (const name of spec?.stories?.(change, doc) ?? []) {
      dirty.story(name);
    }

    if (spec?.publish === 'story' || STRUCTURAL.has(change.kind)) {
      noteStructural(doc, change, dirty);
      return;
    }
  }

  noteTouched(doc, change, dirty);
}
