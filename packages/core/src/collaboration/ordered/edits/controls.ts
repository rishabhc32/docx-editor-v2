// A content control's new value: what the engine proposes, the old words struck and the new
// words inserted. A link's new target is `links.ts`.

import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlElement } from '../../../store/package/ooxml-tree.ts';
import {
  liveRuns,
  proposedProperties,
} from '../../../store/store/tracked-content-control-value.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import type { ContentControlValueInput } from '../../../store/store/tree-op-content-controls.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { insertedPending } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import { isUnit, type OrderedItem } from '../model/types.ts';
import { childNamed, childWhere, isControlContent, isWml } from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import { pendingOwn, strikeItems, suggestionFor, takeBack, type Ctx } from './context.ts';
import { controlPropertiesRecord, sdtPrRecord } from './control-shell.ts';
import { takeBackNotesOf } from './note-body.ts';
import { noteOwnAddition, struckAlready, unitsOf } from './text.ts';
import { shellId, templateItems, viewNode, viewPart } from './templates.ts';

/**
 * A control's new value: the engine sets it on the view, and its words and properties become
 * the suggestion's: the old words struck (your own pending words taken back), the new words
 * inserted at the control's end, and the new `w:sdtPr` shown while pending, with no change
 * record (Word records none). A value the control shows already makes no suggestion (null).
 */
export function setControlValue(
  ctx: Ctx,
  open: OrderedItem,
  value: string | ContentControlValueInput
): string | null {
  if (open.kind !== 'open' || (open.token !== 'sdt' && open.token !== 'sdtBlock'))
    throw new OrderedRefusal('unknown-control', open.space);
  const shell = open.node;
  if (!shell || shell.kind === 'textValue' || !childNamed(shell, 'sdtPr'))
    throw new OrderedRefusal('control-without-properties', open.space);
  const view = viewPart(ctx);
  const id = shellId(open);
  const result = applyTreeOp(
    view,
    { op: 'setContentControlValue', controlId: id, value },
    {
      deferValidation: true,
    }
  );
  if (!result.ok) throw engineRefusal('setContentControlValue', result.reason);
  const control = viewNode(result.part, id);
  const sdtPr = childNamed(control, 'sdtPr');
  if (!sdtPr) throw new OrderedRefusal('engine-refused', 'the value removed the control');

  // A value the control shows already changes nothing: no record
  if (ooxmlTreesEqual(control, viewNode(view, id))) {
    return null;
  }

  // A value in a control the author's own pending suggestion put in is part of that suggestion,
  // as the engine writes it inside the author's own insertion
  const own = pendingOwn(ctx, open.by);
  const suggestion = suggestionFor(ctx, own);
  strikeOldValue(ctx, open, suggestion.id);
  const close = closeOf(ctx.story, open, ctx.doc);

  // A direct edit types the new words into the pending insertion that put the control in, as its
  // author's (`directInsert`), so a reject takes them with the control; else the edit's own
  const holder =
    ctx.direct && open.by !== null && insertedPending(ctx.doc, open) ? open.by : suggestion.id;
  const items = valueItems(ctx, control, open, holder);

  // At the control's end: in a block control, in its paragraph, before the mark
  const last = ctx.story.tree.prev(close);
  const end = open.token === 'sdtBlock' && last?.kind === 'mark' ? ctx.story.tree.prev(last) : last;
  if (items.length > 0) {
    insertItems(ctx.doc, ctx.story, end, items);
  }

  noteOwnAddition(ctx, holder, unitsOf(ctx.story, items));

  // The record sets only the children the value changed (its last value, its date, its
  // check box state), so a change of the control's title or tag stays whatever its order; the
  // prompt flag is compose's to say: it shows the prompt while the control holds nothing else.
  // A suggestion keeps a value record the control stores beside its prompt, as the engine's; a
  // direct value, and one in the author's own pending control, are written as they are
  const shown = viewNode(view, id);
  const before = childNamed(shown, 'sdtPr');
  const after = ctx.direct || own ? sdtPr : proposedProperties(shown, sdtPr);
  const record =
    controlPropertiesRecord(ctx, suggestion.id, before, after, result.part) ??
    sdtPrRecord(ctx, suggestion.id, after, []);
  relabel(ctx.doc, ctx.story, open, { props: [...open.props, record] });
  return suggestion.id;
}

/**
 * Strike a control's old words with the suggestion: the author's own pending words are taken
 * back instead, with a new note whose own reference they held.
 */
function strikeOldValue(ctx: Ctx, open: OrderedItem, suggestion: string): void {
  const taken: OrderedItem[] = [];
  for (const item of itemsInside(ctx.story, open, ctx.doc)) {
    if (!isUnit(item) || item.retracted || item.prompt) {
      continue;
    }

    const own = pendingOwn(ctx, item.by);
    if (own && insertedPending(ctx.doc, item)) {
      takeBack(ctx, ctx.story, item);
      taken.push(item);
    } else if (!struckAlready(ctx, item.dels)) {
      strikeItems(ctx, ctx.story, [item], suggestion);
    }
  }

  takeBackNotesOf(ctx, taken);
}

/**
 * The value's items: the runs the engine wrote in the control (in a block control, in its one
 * paragraph), loaded as a paragraph would be (a check box's symbol is an object), inserted by the
 * suggestion. Only the runs the value shows (`liveRuns`), as the engine proposes a value: the
 * struck words the write kept beside it are the control's items already.
 */
function valueItems(
  ctx: Ctx,
  control: OoxmlElement,
  open: OrderedItem,
  suggestion: string
): OrderedItem[] {
  const content = control.children.find(isControlContent);
  if (!content || content.kind === 'textValue') {
    return [];
  }

  const block = childWhere(content, (c) => c.kind === 'paragraph' || isWml(c, 'p'));
  const paragraph = build(ctx.ids.node(), 'paragraph', 'p', [], liveRuns(block ?? content));
  return templateItems(ctx, paragraph, suggestion, isUnit).map(
    (item): OrderedItem => ({ ...item, by: suggestion, within: open.within })
  );
}
