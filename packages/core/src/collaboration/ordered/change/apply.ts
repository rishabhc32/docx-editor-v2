// Apply one change to a document: the room and the client use this one function, so every copy
// makes the same document with the same ids.

import { orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedEditContext, OrderedReplacedFormat } from '../edits/context.ts';
import { runOrderedEdit, type OrderedEdit, type OrderedEditOutcome } from '../edits/apply.ts';
import { replyToComment, resolveComment } from '../edits/comments.ts';
import { refuseMarkedEmptyFields } from '../edits/fields.ts';
import { orderedPlaceAt, shownItemAt, type OrderedPlace } from '../edits/places.ts';
import { retractSuggestion, undoOwnEdits } from '../edits/retract.ts';
import { revertDirect } from '../edits/revert.ts';
import { isHidden, itemAt, madeBy, storyNamed } from '../model/document.ts';
import { makerOf, personAttribution, type OrderedMaker } from '../model/person.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { recordUndo } from '../model/undo-record.ts';
import {
  isUnit,
  type OrderedDocument,
  type OrderedItem,
  type OrderedStory,
} from '../model/types.ts';
import type { OrderedApplyContext } from '../room/room.ts';
import { applyDecision } from './apply-decide.ts';
import { parseOrderedChangeId } from './ids.ts';
import { placeOfChange } from './change-place.ts';
import { settleRemovedColumns } from './settle-columns.ts';
import { settleHiddenNotes } from './settle-notes.ts';
import { settleLooseWords } from './settle-paragraphs.ts';
import { settleEmptiedControls } from './settle-emptied.ts';
import { settleNamedPrompts } from './settle-named-prompts.ts';
import { settlePromptMarks } from './settle-prompts.ts';
import { formatWrite, paragraphOps } from './apply-properties.ts';
import { orderedKindSpecOf } from './kinds/registry.ts';
import { tableOpToken } from './table-op.ts';
import {
  isOrderedCoreEditChange,
  type OrderedChange,
  type OrderedCoreEditChange,
  type OrderedEditChange,
} from './types.ts';

/** What a change did. @public */
export interface OrderedChangeOutcome {
  /** The suggestions the change made, extended, decided or took back. */
  readonly suggestions: readonly string[];
  /** Policy 2: the format and property records the change took over from other suggestions. */
  readonly replaced: readonly OrderedReplacedFormat[];
}

/** The shown paragraph a shown mark ends, by index. */
function viewOfMark(doc: OrderedDocument, story: OrderedStory, mark: OrderedItem) {
  const view = orderedViewParagraph(doc, story.name, story.tree.before(mark).marks);
  if (!view || view.mark !== mark) throw new OrderedRefusal('hidden-target', mark.space);
  return view;
}

/**
 * The start of the first shown paragraph right inside a container: in its first cell, not in a
 * table nested there. A table edit at that place targets the container itself.
 */
function insidePlace(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): OrderedPlace {
  const close = open.pair === undefined ? null : itemAt(story, { space: open.pair, offset: 1 });
  let nested = 0;
  for (let item = story.tree.next(open); item && item !== close; item = story.tree.next(item)) {
    if (item.token === 'tbl' && item.kind === 'open') {
      nested += 1;
    }

    if (item.token === 'tbl' && item.kind === 'close') {
      nested -= 1;
    }

    if (nested > 0 || item.kind !== 'mark' || isHidden(doc, item)) {
      continue;
    }

    const view = viewOfMark(doc, story, item);
    if (!view.blocks.includes(open)) {
      break;
    }

    return orderedPlaceAt(view, 0);
  }

  throw new OrderedRefusal('hidden-target', `${open.space} shows no paragraph`);
}

/** The first shown unit inside a link: the place its new target is read at. */
function linkPlace(doc: OrderedDocument, story: OrderedStory, open: OrderedItem): OrderedPlace {
  const close = open.pair === undefined ? null : itemAt(story, { space: open.pair, offset: 1 });
  for (let item = story.tree.next(open); item && item !== close; item = story.tree.next(item))
    if (isUnit(item) && !isHidden(doc, item)) return { space: item.space, offset: item.offset };
  throw new OrderedRefusal('hidden-target', `${open.space} shows no text`);
}

/** The place at the start of the paragraph after the one a mark ends: a join's Backspace. */
function joinPlace(doc: OrderedDocument, story: OrderedStory, mark: OrderedItem): OrderedPlace {
  const next = orderedViewParagraph(doc, story.name, story.tree.before(mark).marks + 1);
  if (!next) throw new OrderedRefusal('no-paragraph-after', mark.space);
  return orderedPlaceAt(next, 0);
}

/** The edits a change of one of the first kinds is made of, in order. */
function editsOf(
  doc: OrderedDocument,
  story: OrderedStory,
  change: OrderedCoreEditChange
): OrderedEdit[] {
  switch (change.kind) {
    case 'insert': {
      const inside = change.inside === undefined ? {} : { inside: change.inside };
      return [{ kind: 'insert', at: placeOfChange(change.place), text: change.text, ...inside }];
    }
    case 'insert-tab':
      return [{ kind: 'insertBreak', at: placeOfChange(change.place), break: 'tab' }];
    case 'insert-break':
      return [
        {
          kind: 'insertBreak',
          at: placeOfChange(change.place),
          break: change.break,
          ...(change.clear === undefined ? {} : { clear: change.clear }),
        },
      ];
    case 'delete':
      return [{ kind: 'delete', from: placeOfChange(change.from), to: placeOfChange(change.to) }];
    case 'format':
      return [
        {
          kind: 'format',
          from: placeOfChange(change.from),
          to: placeOfChange(change.to),
          ...formatWrite(change.properties),
        },
      ];
    case 'mark-format': {
      const write = formatWrite(change.properties);

      // Each mark is its own press, all in the change's one suggestion
      return change.marks.map((address) => {
        const mark = shownItemAt(doc, story, address, 'mark');
        return { kind: 'formatMark', at: { space: mark.space, offset: mark.offset }, ...write };
      });
    }
    case 'split':
      return [{ kind: 'split', at: placeOfChange(change.place) }];
    case 'join': {
      const mark = shownItemAt(doc, story, change.mark, 'mark');
      return [{ kind: 'join', at: joinPlace(doc, story, mark) }];
    }
    case 'paragraph-property': {
      const ops = paragraphOps(change.properties);

      // Each paragraph is its own press, all in the change's one suggestion
      return change.marks.map((address) => {
        const mark = shownItemAt(doc, story, address, 'mark');
        return { kind: 'paragraph', at: { space: mark.space, offset: mark.offset }, ops };
      });
    }
    case 'table-property': {
      const op = change.op;
      const token = tableOpToken(op);
      const opens = change.on.map((address) => shownItemAt(doc, story, address, 'open', [token]));
      const first = opens[0];
      if (!first) throw new OrderedRefusal('invalid-edit', 'a table press needs a target');

      // A cell op is one selection of cells; a table or a row op is one press for each
      if (token === 'tc')
        return [
          { kind: 'tableProperty', at: insidePlace(doc, story, first), op, cells: change.on },
        ];
      return opens.map((open) => ({
        kind: 'tableProperty',
        at: insidePlace(doc, story, open),
        op,
      }));
    }
    case 'insert-row':
    case 'delete-row': {
      const row = shownItemAt(doc, story, change.row, 'open', ['tr']);
      const at = insidePlace(doc, story, row);
      return [
        change.kind === 'insert-row'
          ? { kind: 'insertRow', at, where: change.where }
          : { kind: 'deleteRow', at },
      ];
    }
    case 'insert-column':
    case 'delete-column': {
      const cell = shownItemAt(doc, story, change.cell, 'open', ['tc']);
      const at = insidePlace(doc, story, cell);
      return [
        change.kind === 'insert-column'
          ? { kind: 'insertColumn', at, where: change.where }
          : { kind: 'deleteColumn', at },
      ];
    }
    case 'insert-table':
      return [
        {
          kind: 'insertTable',
          at: placeOfChange(change.place),
          rows: change.rows,
          cols: change.columns,
          columnWidth: change.columnWidth,
        },
      ];
    case 'control-value':
      shownItemAt(doc, story, change.control, 'open', ['sdt', 'sdtBlock']);
      return [{ kind: 'controlValue', control: change.control, value: change.value }];
    case 'link-target': {
      const link = shownItemAt(doc, story, change.link, 'open', ['hyperlink']);
      const at = linkPlace(doc, story, link);
      const shared = {
        link: change.link,
        ...(change.tooltip === undefined ? {} : { tooltip: change.tooltip }),
        ...(change.from === undefined || change.to === undefined
          ? {}
          : { from: placeOfChange(change.from), to: placeOfChange(change.to) }),
      };
      return [
        change.href !== undefined
          ? { kind: 'linkTarget', at, url: change.href, ...shared }
          : {
              kind: 'linkTarget',
              at,
              ...(change.anchor === undefined ? {} : { anchor: change.anchor }),
              ...shared,
            },
      ];
    }
    case 'move':
      return [
        {
          kind: 'move',
          from: placeOfChange(change.from),
          to: placeOfChange(change.to),
          at: placeOfChange(change.place),
        },
      ];
    case 'insert-note':
      return [
        {
          kind: 'insertNote',
          at: placeOfChange(change.place),
          note: change.note,
          text: change.text,
        },
      ];
  }
}

/** `into` names the sender's own pending suggestion; a direct edit adds to none. */
function checkInto(doc: OrderedDocument, change: OrderedEditChange, maker: OrderedMaker): void {
  if (change.into === undefined) {
    return;
  }

  if (change.mode === 'edit') {
    throw new OrderedRefusal('invalid-edit', 'a direct edit adds to no suggestion');
  }

  const into = doc.suggestions.get(change.into);
  if (!into) {
    throw new OrderedRefusal('unknown-suggestion', change.into);
  }

  if (into.status !== 'pending') {
    throw new OrderedRefusal('not-pending', change.into);
  }

  if (!madeBy(doc, change.into, maker)) {
    throw new OrderedRefusal('not-own-suggestion', change.into);
  }
}

/** The edit context of a change: the room's stamp and clock; `mode: 'edit'` makes it direct. */
function editContext(
  doc: OrderedDocument,
  change: OrderedEditChange,
  ctx: OrderedApplyContext
): OrderedEditContext {
  return {
    doc,
    story: change.part,
    attribution: personAttribution(ctx.person, ctx.date),
    ids: ctx.ids,
    ...(change.mode === 'edit' ? { direct: true } : {}),
    ...(ctx.contentLockPolicy === undefined ? {} : { contentLockPolicy: ctx.contentLockPolicy }),
  };
}

function outcome(outcomes: readonly OrderedEditOutcome[]): OrderedChangeOutcome {
  return {
    suggestions: [...new Set(outcomes.flatMap((o) => o.suggestions))],
    replaced: outcomes.flatMap((o) => o.replaced),
  };
}

/**
 * Apply one change to a document, inside the undo record the caller opened (the room's, or the
 * client's for a pending change). It holds whole, or it throws `OrderedRefusal` and the
 * caller's record takes back every write. It has the room's `applyChange` signature.
 *
 * @public
 */
export function applyOrderedChange(
  doc: OrderedDocument,
  change: OrderedChange,
  ctx: OrderedApplyContext
): OrderedChangeOutcome {
  if (doc.undo === null) throw new Error('A change applies inside an undo record.');

  // The person of the change's namespace, whom the typer read names
  const namespace = parseOrderedChangeId(change.id)?.namespace;
  if (namespace !== undefined && !doc.namespaces.has(namespace)) {
    doc.namespaces.set(namespace, ctx.person);
    recordUndo(doc, () => doc.namespaces.delete(namespace));
  }

  const done = applyChange(doc, change, ctx);

  // A suggestion's mark on a simple field with no result has no place in the Word file
  refuseMarkedEmptyFields(doc, done.suggestions);

  // A control a decision emptied shows its prompt as units, so typing can replace it
  settleEmptiedControls(doc, ctx);

  // A control that names a glossary prompt shows it once its content went, as the engine does
  settleNamedPrompts(doc, ctx);

  // A block prompt of two or more paragraphs keeps only the paragraphs content took over
  settlePromptMarks(doc);
  ctx.onApplied?.(done);
  return done;
}

/**
 * A change as the model's own operations: an edit as `OrderedEdit`s with one id minter, a
 * take-back as `retractSuggestion`, a decision as status changes, a comment as a comment write.
 */
function applyChange(
  doc: OrderedDocument,
  change: OrderedChange,
  ctx: OrderedApplyContext
): OrderedChangeOutcome {
  switch (change.kind) {
    case 'decide': {
      const decided = applyDecision(doc, change);
      settleRemovedColumns(doc, decided, ctx.ids);
      settleHiddenNotes(doc, decided);
      settleLooseWords(doc, decided);
      return { suggestions: decided, replaced: [] };
    }
    case 'retract': {
      for (const id of change.suggestions) {
        retractSuggestion(doc, id, makerOf(ctx.person));
      }

      const undone = undoOwnEdits(doc, change.undo ?? []);
      const suggestions = [...new Set([...change.suggestions, ...undone])];
      settleLooseWords(doc, suggestions);
      return { suggestions, replaced: [] };
    }
    case 'revert':
      for (const id of change.suggestions) {
        revertDirect(doc, id, makerOf(ctx.person), change.status);
      }

      return { suggestions: change.suggestions, replaced: [] };
    case 'comment': {
      const edit: OrderedEdit = {
        kind: 'comment',
        from: placeOfChange(change.from),
        to: placeOfChange(change.to),
        text: change.text,
      };
      const attribution = personAttribution(ctx.person, ctx.date);
      return outcome([
        runOrderedEdit({ doc, story: change.part, attribution, ids: ctx.ids }, edit),
      ]);
    }
    case 'comment-reply': {
      const writer = { doc, attribution: personAttribution(ctx.person, ctx.date), ids: ctx.ids };
      replyToComment(writer, change.comment, change.text);
      return { suggestions: [], replaced: [] };
    }
    case 'comment-resolve': {
      const writer = { doc, attribution: personAttribution(ctx.person, ctx.date), ids: ctx.ids };
      resolveComment(writer, change.comment, change.note);
      return { suggestions: [], replaced: [] };
    }
    default: {
      const story = storyNamed(doc, change.part);
      checkInto(doc, change, makerOf(ctx.person));
      const context = editContext(doc, change, ctx);

      // A kind a group added applies through its spec, in the same record
      const spec = orderedKindSpecOf(change.kind);
      if (spec !== undefined) {
        const edit = change.into === undefined ? context : { ...context, join: change.into };
        return outcome([spec.apply(change, { doc, story, edit })]);
      }

      if (!isOrderedCoreEditChange(change)) {
        // A kind neither the first kinds nor a group's spec covers: a client of another code
        throw new OrderedRefusal('not-covered', `a change of kind ${change.kind.slice(0, 64)}`);
      }

      return applyCoreEdits(doc, story, change, context);
    }
  }
}

/**
 * Run the edits a change of the first kinds is made of. Places are addresses: each edit
 * resolves them in the view the edits before it left. A suggesting change is one suggestion:
 * its later edits join the first one's.
 */
function applyCoreEdits(
  doc: OrderedDocument,
  story: OrderedStory,
  change: OrderedCoreEditChange,
  context: OrderedEditContext
): OrderedChangeOutcome {
  const own = `S:${change.id}.`;
  let join = change.into;
  const done: OrderedEditOutcome[] = [];
  for (const edit of editsOf(doc, story, change)) {
    const result = runOrderedEdit({ ...context, ...(join === undefined ? {} : { join }) }, edit);
    done.push(result);
    if (join === undefined && change.mode === 'suggest') {
      join = result.suggestions.find((id) => id.startsWith(own));
    }
  }

  return outcome(done);
}
