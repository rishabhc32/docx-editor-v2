// A content control's new value: what the engine proposes, the old words struck and the new
// words inserted. A link's new target is `links.ts`.

import type { OoxmlElement, OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { contentControlTextOf } from '../../../store/package/content-control-nodes.ts';
import { applyTreeOp } from '../../../store/store/tree-op-apply.ts';
import type { ContentControlValueInput } from '../../../store/store/tree-op-content-controls.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import { createOrderedItem, insertedPending } from '../model/document.ts';
import { insertItems, relabel } from '../model/mutations.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import { isUnit, type OrderedItem } from '../model/types.ts';
import {
  childNamed,
  childWhere,
  isControlContent,
  isRunProperties,
  isWml,
} from '../ooxml-nodes.ts';
import { closeOf, itemsInside } from '../placement.ts';
import {
  pendingOwn,
  strikeItems,
  suggestionFor,
  takeBack,
  withNewIds,
  type Ctx,
} from './context.ts';
import { controlPropertiesRecord, sdtPrRecord } from './control-shell.ts';
import { takeBackNotesOf } from './note-body.ts';
import { struckAlready } from './text.ts';
import { shellId, templateItems, viewNode, viewPart } from './templates.ts';

/** The first run's properties in a control's content, as the engine wrote its value. */
function valueFace(control: OoxmlNode, depth = 0): OoxmlElement | null {
  if (control.kind === 'textValue' || depth > 16) return null;
  if (control.kind === 'run') return childWhere(control, isRunProperties) ?? null;
  for (const child of control.children) {
    if (child.kind === 'textValue' || isWml(child, 'sdtPr')) continue;
    const found = valueFace(child, depth + 1);
    if (found) return found;
  }

  return null;
}

/**
 * A control's new value: the engine sets it on the view, and its words and properties become
 * the suggestion's: the old words struck (your own pending words taken back), the new words
 * inserted at the control's end, and the new `w:sdtPr` shown while pending, with no change
 * record (Word records none).
 */
export function setControlValue(
  ctx: Ctx,
  open: OrderedItem,
  value: string | ContentControlValueInput
): string {
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
  const suggestion = suggestionFor(ctx);
  strikeOldValue(ctx, open, suggestion.id);
  const close = closeOf(ctx.story, open, ctx.doc);
  const items = valueItems(ctx, control, open, suggestion.id);
  if (items.length > 0) insertItems(ctx.doc, ctx.story, ctx.story.tree.prev(close), items);

  // The record sets only the children the value changed (its last value, its date, its
  // check box state), so a change of the control's title or tag stays whatever its order; the
  // prompt flag is compose's to say: it shows the prompt while the control holds nothing else
  const before = childNamed(viewNode(view, id), 'sdtPr');
  const record =
    controlPropertiesRecord(ctx, suggestion.id, before, sdtPr, result.part) ??
    sdtPrRecord(ctx, suggestion.id, sdtPr, []);
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
 * The value's items: the runs the engine wrote in the control, loaded as a paragraph would be
 * (a check box's symbol is an object), inserted by the suggestion.
 */
function valueItems(
  ctx: Ctx,
  control: OoxmlElement,
  open: OrderedItem,
  suggestion: string
): OrderedItem[] {
  const content = control.children.find(isControlContent);
  const runs = content && content.kind !== 'textValue' ? content.children : [];
  if (runs.some((c) => c.kind === 'paragraph' || isWml(c, 'p'))) {
    // A block control's value: its text, in the first run's face
    const display = contentControlTextOf(control);
    const face = valueFace(control);
    if (display.length === 0) return [];
    const space = ctx.ids.units();
    return [
      createOrderedItem({
        space,
        kind: 'text',
        length: display.length,
        text: display,
        face: face ? withNewIds(ctx, face) : null,
        run: space,
        by: suggestion,
        within: open.within,
      }),
    ];
  }

  const paragraph = build(ctx.ids.node(), 'paragraph', 'p', [], runs);
  return templateItems(ctx, paragraph, suggestion, isUnit).map(
    (item): OrderedItem => ({ ...item, by: suggestion, within: open.within })
  );
}
