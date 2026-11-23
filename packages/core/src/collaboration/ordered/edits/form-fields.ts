// Legacy form fields (task 9.6): a check box's state, a drop-down's entry, a text field's
// options and default, its result formatted when the user leaves it, and typing in its result.

import { ooxmlTreesEqual } from '../../../store/package/ooxml-serialize.ts';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import { isInlineRunContainer } from '../../../store/package/ooxml-shared.ts';
import { parsedFieldSpansOf } from '../../../store/package/field-nodes.ts';
import type {
  OoxmlNode,
  OoxmlParagraphNode,
  OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { legacyCheckboxFieldsOf } from '../../../store/store/legacy-checkbox-fields.ts';
import { legacyDropdownFieldsOf } from '../../../store/store/legacy-dropdown-fields.ts';
import type { TextFormFieldOptions } from '../../../store/store/text-form-field-options.ts';
import { textFormFieldsOf } from '../../../store/store/text-form-fields.ts';
import { applyProtectedTextFormEdit } from '../../../store/store/tree-op-field-results.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import type { OrderedViewParagraph } from '../compose/context.ts';
import { isHidden } from '../model/document.ts';
import { insertItems } from '../model/mutations.ts';
import { engineRefusal, OrderedRefusal } from '../model/refusal.ts';
import type { OrderedItem } from '../model/types.ts';
import { isControlContent, isWml } from '../ooxml-nodes.ts';
import { refuseLocked } from '../placement.ts';
import { strikeItems, suggestionFor, type Ctx } from './context.ts';
import { engineApply } from './control-shell.ts';
import { templateItems, viewPart } from './templates.ts';

/** What a form field edit does, past the field it names. @public */
export type OrderedFormFieldOp =
  | { readonly op: 'setLegacyCheckbox'; readonly checked: boolean }
  | { readonly op: 'setLegacyDropdown'; readonly selectedIndex: number }
  | {
      readonly op: 'setTextFormFieldDefault';
      readonly text: string;
      readonly options?: TextFormFieldOptions;
      readonly locale?: string;
    }
  | { readonly op: 'commitTextFormField'; readonly locale?: string }
  /** Typing in a text field's result: `offset` counts from the result's start. */
  | { readonly op: 'insertText'; readonly offset: number; readonly text: string }
  /** A deletion in a text field's result, `[start, end)` from the result's start. */
  | { readonly op: 'deleteText'; readonly start: number; readonly end: number };

/** A form field in a paragraph of the engine's offsets: its begin node and its span. */
interface FieldSpan {
  readonly fieldNodeId: string;
  readonly start: number;
  readonly end: number;
}

/** The fields of the kind an op edits, in a paragraph. */
function fieldsFor(op: OrderedFormFieldOp, paragraph: OoxmlParagraphNode): readonly FieldSpan[] {
  switch (op.op) {
    case 'setLegacyCheckbox':
      return legacyCheckboxFieldsOf(paragraph);
    case 'setLegacyDropdown':
      return legacyDropdownFieldsOf(paragraph);
    default:
      return textFormFieldsOf(paragraph);
  }
}

const paragraphOf = (part: OoxmlPart, id: string): OoxmlParagraphNode => {
  const node = findNode(part, id);
  if (!node || node.kind !== 'paragraph')
    throw new OrderedRefusal('malformed-story', `${id} is not a paragraph`);
  return node;
};

/** The engine's op for a form field edit, on a paragraph of the view. */
function engineOp(op: OrderedFormFieldOp, paragraphId: string, field: FieldSpan): TreeDocOp {
  const named = { paragraphId, fieldNodeId: field.fieldNodeId };
  switch (op.op) {
    case 'setLegacyCheckbox':
      return { op: 'setLegacyCheckbox', ...named, checked: op.checked };
    case 'setLegacyDropdown':
      return { op: 'setLegacyDropdown', ...named, selectedIndex: op.selectedIndex };
    case 'setTextFormFieldDefault':
      return {
        op: 'setTextFormFieldDefault',
        ...named,
        text: op.text,
        ...(op.options === undefined ? {} : { options: op.options }),
        ...(op.locale === undefined ? {} : { locale: op.locale }),
      };
    case 'commitTextFormField':
      return {
        op: 'commitTextFormField',
        ...named,
        ...(op.locale === undefined ? {} : { locale: op.locale }),
      };
    case 'insertText':
      return {
        op: 'insertText',
        paragraphId,
        offset: field.start + op.offset,
        text: op.text,
        textFormFieldId: field.fieldNodeId,
      };
    case 'deleteText':
      return {
        op: 'deleteText',
        paragraphId,
        start: field.start + op.start,
        end: field.start + op.end,
        textFormFieldId: field.fieldNodeId,
      };
  }
}

/** The engine's result for the op, refused by name as the engine refuses it. */
function engineResult(
  view: OoxmlPart,
  op: TreeDocOp,
  typing: boolean,
  field: FieldSpan
): OoxmlPart {
  if (typing) {
    // Filling a text field: the store's own protected edit, which keeps the text in the result
    const paragraph = paragraphOf(view, 'paragraphId' in op ? op.paragraphId : '');
    const range = textFormFieldsOf(paragraph).find((f) => f.fieldNodeId === field.fieldNodeId);
    if (!range) throw new OrderedRefusal('engine-refused', `${op.op}: no text field`);
    const applied = applyProtectedTextFormEdit(view, op, range, { deferValidation: true });
    if (!applied.ok) throw engineRefusal(op.op, applied.reason);
    return applied.part;
  }

  return engineApply(view, op);
}

/** Whether a node is or holds a `w:fldChar` of a type. */
function holdsFieldChar(node: OoxmlNode | null, type: string, depth = 0): boolean {
  if (!node || node.kind === 'textValue' || depth > 8) return false;
  if (
    node.localName === 'fldChar' &&
    node.attributes.some((a) => a.localName === 'fldCharType' && a.value === type)
  )
    return true;
  return node.children.some((child) => holdsFieldChar(child, type, depth + 1));
}

/** The list's items of the field the view character at `start` begins, first to last. */
function fieldItems(ctx: Ctx, p: OrderedViewParagraph, field: FieldSpan): OrderedItem[] {
  const first = p.chars[field.start];
  const last = p.chars[field.end - 1];
  if (!first || !last) throw new OrderedRefusal('unknown-address', 'no form field there');

  // A field the list keeps whole: one object
  if (first.item.pieces !== undefined && first.item === last.item) return [first.item];
  const items = p.items;
  const firstIndex = items.indexOf(first.item);
  const lastIndex = items.indexOf(last.item);
  if (firstIndex < 0 || lastIndex < 0)
    throw new OrderedRefusal('malformed-story', 'a field outside its paragraph');

  // Back to the field's begin, past its instruction and separator; on to its end
  const from = fieldEdge(ctx, items, firstIndex, -1);
  const to = fieldEdge(ctx, items, lastIndex, 1);
  return items.slice(from, to + 1).filter((item) => !isHidden(ctx.doc, item));
}

/**
 * The index a walk from `at` stops at: back (`-1`) at the field's begin, on (`1`) at its end,
 * past the fields nested in it; else at the paragraph's edge. Hidden items do not count.
 */
function fieldEdge(ctx: Ctx, items: readonly OrderedItem[], at: number, step: -1 | 1): number {
  const [stop, nest] = step < 0 ? ['begin', 'end'] : ['end', 'begin'];
  let index = at;
  for (let depth = 0; index + step >= 0 && index + step < items.length; ) {
    index += step;
    const item = items[index];
    if (!item || isHidden(ctx.doc, item)) {
      continue;
    }

    if (holdsFieldChar(item.node, nest)) {
      depth += 1;
    } else if (holdsFieldChar(item.node, stop)) {
      if (depth === 0) {
        break;
      }

      depth -= 1;
    }
  }

  return index;
}

const MAX_DEPTH = 32;

/**
 * The children of an inline container that the list holds as open and close items (a control,
 * a link, a smart tag), or null for any other node.
 */
function containerChildren(node: OoxmlNode | undefined): readonly OoxmlNode[] | null {
  if (!node || node.kind === 'textValue') return null;
  if (node.kind === 'contentControl' || isWml(node, 'sdt')) {
    const content = node.children.find(isControlContent);
    return content && content.kind !== 'textValue' ? content.children : null;
  }

  return node.kind === 'hyperlink' || (node.kind === 'generic' && isInlineRunContainer(node))
    ? node.children
    : null;
}

/**
 * The nodes that hold a field, from its begin to its end, in the engine's result: the
 * paragraph's children, or those of the inline container that holds the whole field. The
 * container's other words are no part of the field (seed 1014: a text field in a control).
 */
function fieldNodes(paragraph: OoxmlParagraphNode, fieldNodeId: string): OoxmlNode[] {
  const span = parsedFieldSpansOf(paragraph).find((f) => f.node.id === fieldNodeId);
  const endId = span?.removeNodeIds[span.removeNodeIds.length - 1];
  if (!span || endId === undefined)
    throw new OrderedRefusal('engine-refused', 'the engine wrote no field');
  let children: readonly OoxmlNode[] = paragraph.children;
  for (let depth = 0; ; depth += 1) {
    const from = children.findIndex((child) => containsId(child, fieldNodeId));
    const to = children.findIndex((child) => containsId(child, endId));
    if (from < 0 || to < from) throw new OrderedRefusal('engine-refused', 'the field is not whole');
    const inner = from === to && depth < MAX_DEPTH ? containerChildren(children[from]) : null;
    if (inner === null) return children.slice(from, to + 1);
    children = inner;
  }
}

/** Whether a node is, or holds, the node with this id. */
function containsId(node: OoxmlNode, id: string, depth = 0): boolean {
  if (node.id === id) return true;
  if (node.kind === 'textValue' || depth > MAX_DEPTH) return false;
  return node.children.some((child) => containsId(child, id, depth + 1));
}

/**
 * A form field edit: the engine's op on the view, and the field it wrote in place of the
 * field's items. Direct only. Returns the edit's suggestion, or null when the op changes nothing.
 */
export function editFormField(
  ctx: Ctx,
  p: OrderedViewParagraph,
  start: number,
  op: OrderedFormFieldOp
): string | null {
  // Word fills legacy form fields only in a document protected for forms, which turns tracking
  // off: ECMA-376 `ST_DocProtect` takes one value, and `CT_FFData` has no change record
  if (!ctx.direct)
    throw new OrderedRefusal(
      'invalid-edit',
      'a form field has no tracked form: a direct edit only'
    );
  refuseLocked(ctx, p, start, start + 1);
  const view = viewPart(ctx);
  const field = fieldsFor(op, paragraphOf(view, p.id)).find((f) => f.start === start);
  if (!field) throw new OrderedRefusal('unknown-address', `no form field at ${start}`);
  const typing = op.op === 'insertText' || op.op === 'deleteText';
  const result = engineResult(view, engineOp(op, p.id, field), typing, field);

  // An op that leaves the paragraph as it shows changes nothing: no record
  if (ooxmlTreesEqual(paragraphOf(result, p.id), paragraphOf(view, p.id))) {
    return null;
  }

  const nodes = fieldNodes(paragraphOf(result, p.id), field.fieldNodeId);
  const old = fieldItems(ctx, p, field);
  const suggestion = suggestionFor(ctx);

  // The old items are struck and the new ones go after them, so `revert` gives the field back
  strikeItems(ctx, ctx.story, old, suggestion.id);
  const scratch = build(ctx.ids.node(), 'paragraph', 'p', [], nodes);
  const made = templateItems(ctx, scratch, suggestion.id)
    .filter((item) => item.kind !== 'mark')
    .map((item): OrderedItem => ({ ...item, by: suggestion.id, within: p.mark.within }));
  const last = old[old.length - 1];
  if (!last) throw new OrderedRefusal('malformed-story', 'a field with no items');
  insertItems(ctx.doc, ctx.story, last, made);
  return suggestion.id;
}
