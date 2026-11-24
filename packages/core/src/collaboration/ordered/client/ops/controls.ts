// Control ops: a control's value (task 6.1); new, removed and changed controls, repeating
// section items, building blocks and legacy form fields (group 9.6).

import { contentControlPropertiesOf } from '../../../../store/package/content-control-nodes.ts';
import { customNodeBindingTarget } from '../../../../store/package/custom-node-payloads.ts';
import { readCustomXmlNode } from '../../../../store/package/custom-xml-nodes.ts';
import { customXmlDataParts } from '../../../../store/package/custom-xml-part.ts';
import { glossaryPartOf } from '../../../../store/package/building-blocks.ts';
import { canonicalTreeDifference } from '../../../../store/package/ooxml-difference.ts';
import { findNode } from '../../../../store/package/ooxml-edit.ts';
import type { OoxmlNode, OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { inheritedNamespaceBindings } from '../../../../store/package/wml-namespace.ts';
import { paragraphTextOf } from '../../../../store/store/paragraph-model-text.ts';
import { applyTreeOp } from '../../../../store/store/tree-op-apply.ts';
import { legacyCheckboxFieldsOf } from '../../../../store/store/legacy-checkbox-fields.ts';
import { legacyDropdownFieldsOf } from '../../../../store/store/legacy-dropdown-fields.ts';
import { textFormFieldsOf } from '../../../../store/store/text-form-fields.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import type { OrderedFormFieldOp } from '../../edits/form-fields.ts';
import type { ContentControlValueInput } from '../../../../store/store/tree-op-content-controls.ts';
import { parseCheckboxValue } from '../../../../store/store/tree-op-nodes.ts';
import type { OrderedControlBinding, OrderedCustomPayload } from '../../change/kinds/controls.ts';
import type { OrderedControlValue } from '../../change/types.ts';
import { isHidden } from '../../model/document.ts';
import type { OrderedItem } from '../../model/types.ts';
import { addressOf, placeAt, type ListParagraph } from '../paragraphs.ts';
import { OrderedClientRefusal } from '../refusal.ts';
import { containerOf, type ListContainer } from './containers.ts';
import {
  translator,
  type OrderedChangeBody,
  type OrderedOpOption,
  type OrderedOpTable,
  type OrderedTranslateBatch,
  type OrderedTranslateStep,
} from './types.ts';
import { untrackedRule } from './untracked.ts';

/**
 * The node a custom node's binding names, as the package the ops apply to holds it: the write's
 * own store edit put it there. A binding of another shape names nothing the room carries.
 */
function customPayloadOf(
  batch: OrderedTranslateBatch,
  binding: OrderedControlBinding
): { readonly payload?: OrderedCustomPayload } {
  const target = customNodeBindingTarget(binding);
  if (!target) return {};
  const pkg = batch.gated();
  const storeItemId = binding.storeItemId.toUpperCase();
  const store = pkg
    ? customXmlDataParts(pkg, pkg.mainDocumentPart).find(
        (candidate) => candidate.itemId.toUpperCase() === storeItemId
      )
    : undefined;
  const node = pkg && store ? readCustomXmlNode(pkg, store.partName, target.nodeId) : null;
  if (!node)
    throw new OrderedClientRefusal('ordered-unsupported-op', 'insertInlineContentControl.payload');
  return { payload: { label: node.label, data: node.data } };
}

/** A control's value as the change format says it, read by the control's own type. */
function controlValueOf(
  open: OrderedItem,
  value: string | ContentControlValueInput
): OrderedControlValue {
  if (typeof value !== 'string') return value;
  if (!open.node) throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a control');

  // The editor's string form means what the control's type says, as the engine reads it: a
  // combo box's names an item by its value, or else is the combo box's own text
  const properties = contentControlPropertiesOf(open.node);
  switch (properties.type) {
    case 'dropDownList':
      return { kind: 'listItem', value };
    case 'comboBox':
      return properties.listItems.some((item) => item.value === value)
        ? { kind: 'listItem', value }
        : { kind: 'text', text: value };
    case 'checkbox': {
      const checked = parseCheckboxValue(value);
      if (checked === null)
        throw new OrderedClientRefusal('ordered-unsupported-op', 'setContentControlValue.checkbox');
      return { kind: 'checkbox', checked };
    }
    case 'date':
      return { kind: 'date', iso: value };
    default:
      return { kind: 'text', text: value };
  }
}

/** The engine's reader of each kind of legacy form field in a paragraph. */
const FORM_FIELD_READERS = {
  checkbox: legacyCheckboxFieldsOf,
  dropdown: legacyDropdownFieldsOf,
  text: textFormFieldsOf,
} as const;

/** A legacy form field the editor names: its paragraph, and its span as the editor counts it. */
function formFieldAt(
  batch: OrderedTranslateBatch,
  paragraphId: string,
  fieldNodeId: string,
  kind: 'checkbox' | 'dropdown' | 'text'
): { p: ListParagraph; start: number; end: number } {
  const p = batch.paragraph(paragraphId);
  const shown = batch.shown()?.parts.get(p.story.name);
  const node = shown ? findNode(shown, paragraphId) : null;
  if (!node || node.kind !== 'paragraph')
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'a form field');
  const field = FORM_FIELD_READERS[kind](node).find((f) => f.fieldNodeId === fieldNodeId);
  if (!field)
    throw new OrderedClientRefusal('ordered-refused', 'out-of-step', 'no such form field');
  return { p, start: field.start, end: field.end };
}

/** A form field edit: `form-text` when it changes the text a reader sees, else `form-field`. */
const formFieldChange = (
  p: ListParagraph,
  start: number,
  op: OrderedFormFieldOp,
  text: boolean
): OrderedChangeBody => ({
  kind: text ? 'form-text' : 'form-field',
  part: p.story.name,
  field: placeAt(p, start),
  op,
});

/** Whether the engine's op changes a paragraph's text on the part the editor shows. */
function changesText(
  batch: OrderedTranslateBatch,
  story: string,
  op: TreeDocOp & { readonly paragraphId: string }
): boolean {
  const part = batch.shown()?.parts.get(story);
  if (!part) return true;
  const applied = applyTreeOp(part, op, { deferValidation: true });
  return (
    !applied.ok ||
    paragraphTextOf(applied.part, op.paragraphId) !== paragraphTextOf(part, op.paragraphId)
  );
}

/**
 * Typing or a deletion in a legacy text form field's result (`textFormFieldId`), which the
 * editor sends only under forms protection: the store's protected edit keeps the text in the
 * result, so the list makes the same edit of the field. The engine refuses it as a suggestion.
 */
export function formFieldTyping(
  op: Extract<TreeDocOp, { readonly op: 'insertText' | 'deleteText' }>,
  batch: OrderedTranslateBatch,
  fieldNodeId: string
): OrderedTranslateStep {
  if (op.revision !== undefined)
    throw new OrderedClientRefusal('ordered-unsupported-op', `${op.op}.textFormFieldId`, 'tracked');
  return () => {
    const { p, start } = formFieldAt(batch, op.paragraphId, fieldNodeId, 'text');
    if (op.op === 'insertText')
      return [
        formFieldChange(
          p,
          start,
          { op: 'insertText', offset: op.offset - start, text: op.text },
          true
        ),
      ];
    if (op.end <= op.start) return [];
    return [
      formFieldChange(
        p,
        start,
        { op: 'deleteText', start: op.start - start, end: op.end - start },
        true
      ),
    ];
  };
}

/** A control the editor names, inline or block. */
const controlOf = (doc: Parameters<typeof containerOf>[0], nodeId: string): ListContainer =>
  containerOf(doc, nodeId, ['sdt', 'sdtBlock']);

/**
 * Check the paragraph an inline control is in through the batch: its text is checked against
 * the editor's once, so a later op of the batch (a custom node's update puts a new control where
 * the old one was) reads the paragraph as this change leaves it.
 */
function checkHoldingParagraph(batch: OrderedTranslateBatch, control: ListContainer): void {
  if (control.open.token !== 'sdt') {
    return;
  }

  for (let at = control.story.tree.next(control.open); at; at = control.story.tree.next(at)) {
    if (at.kind === 'mark' && !isHidden(batch.doc, at)) {
      if (at.node) {
        batch.paragraph(at.node.id);
      }

      return;
    }
  }
}

/** A control op's tag, alias and lock, each when the op names it. */
function controlNames<Name, Lock>(op: {
  readonly tag?: Name;
  readonly alias?: Name;
  readonly lock?: Lock;
}) {
  return {
    ...(op.tag === undefined ? {} : { tag: op.tag }),
    ...(op.alias === undefined ? {} : { alias: op.alias }),
    ...(op.lock === undefined ? {} : { lock: op.lock }),
  };
}

/**
 * Whether the blocks the editor picked are the ones a building block's name gives every copy.
 * Both are the glossary's: they compare under its scope, where a shape's choice finds its
 * prefixes.
 */
function sameBlocks(
  named: readonly OoxmlNode[],
  picked: readonly OoxmlNode[],
  glossary: OoxmlPart | null
): boolean {
  if (named.length !== picked.length) {
    return false;
  }

  return named.every((block, k) => {
    const other = picked[k];
    const bindings = glossary ? inheritedNamespaceBindings(glossary, block) : undefined;
    return other !== undefined && canonicalTreeDifference(block, other, { bindings }).equal;
  });
}

export const CONTROL_OPS: OrderedOpTable<
  | 'setContentControlValue'
  | 'insertInlineContentControl'
  | 'insertContentControl'
  | 'removeContentControl'
  | 'setContentControlProperties'
  | 'addRepeatingSectionItem'
  | 'removeRepeatingSectionItem'
  | 'insertBuildingBlock'
  | 'setTextFormFieldDefault'
  | 'commitTextFormField'
  | 'setLegacyCheckbox'
  | 'setLegacyDropdown'
> = {
  setContentControlValue: {
    status: 'translated',
    group: '6.1',
    changes: ['control-value'],
    party: null,
    note: "A control's value; the editor's string form read by the control's type.",
    translate: translator('setContentControlValue', (op, batch) => () => {
      const control = controlOf(batch.doc, op.controlId);
      const value = controlValueOf(control.open, op.value);
      return [
        {
          kind: 'control-value',
          part: control.story.name,
          control: addressOf(control.open),
          value,
        },
      ];
    }),
  },
  insertInlineContentControl: {
    status: 'translated',
    group: '9.6',
    changes: ['insert-control'],
    party: untrackedRule('text', 'a new inline control with its text'),
    note: 'A custom node: a control holding text at a caret, with a tag, an alias, a lock, a data binding and the payload node the binding names.',
    translate: translator('insertInlineContentControl', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      return [
        {
          kind: 'insert-control',
          part: p.story.name,
          place: placeAt(p, op.offset),
          control: {
            type: 'custom',
            tag: op.tag,
            text: op.text,
            ...(op.alias === undefined ? {} : { alias: op.alias }),
            ...(op.lock === undefined ? {} : { lock: op.lock }),
            ...(op.dataBinding === undefined ? {} : { binding: op.dataBinding }),
            ...(op.dataBinding === undefined ? {} : customPayloadOf(batch, op.dataBinding)),
          },
        },
      ];
    }),
  },
  insertContentControl: {
    status: 'translated',
    group: '9.6',
    changes: ['insert-control', 'wrap-control'],
    party: null,
    note: 'A control around a range, of a type (tracked in suggesting mode, from the toolbar and the editing API alike: the words struck and a copy in the new control), or an empty one at a caret showing its prompt. A control at a caret, or around words that hold review markup, has no tracked form and goes untracked: a party is then refused as a direct edit (`text` at a caret, `package` for a range).',
    translate: translator('insertContentControl', (op, batch) => () => {
      const p = batch.paragraph(op.paragraphId);
      const control = { type: op.type, ...controlNames(op) };
      if (op.start === op.end)
        return [
          { kind: 'insert-control', part: p.story.name, place: placeAt(p, op.start), control },
        ];
      return [
        {
          kind: 'wrap-control',
          part: p.story.name,
          from: placeAt(p, op.start),
          to: placeAt(p, op.end),
          control,
        },
      ];
    }),
  },
  removeContentControl: {
    status: 'translated',
    group: '9.6',
    changes: ['remove-control', 'delete-control'],
    party: untrackedRule(
      'package',
      'the control goes (its text too without `keepContent`, `text`)'
    ),
    note: 'Remove a control, keeping its content (`remove-control`) or not (`delete-control`).',
    translate: translator('removeContentControl', (op, batch) => () => {
      const control = controlOf(batch.doc, op.controlId);
      checkHoldingParagraph(batch, control);
      return [
        {
          kind: op.keepContent === false ? 'delete-control' : 'remove-control',
          part: control.story.name,
          control: addressOf(control.open),
        },
      ];
    }),
  },
  setContentControlProperties: {
    status: 'translated',
    group: '9.6',
    changes: ['control-property'],
    party: untrackedRule('package', "a control's tag, alias and lock"),
    note: "A control's tag, alias and lock: Word records no change of them.",
    translate: translator('setContentControlProperties', (op, batch) => () => {
      const control = controlOf(batch.doc, op.controlId);
      return [
        {
          kind: 'control-property',
          part: control.story.name,
          control: addressOf(control.open),
          ...controlNames(op),
        },
      ];
    }),
  },
  addRepeatingSectionItem: {
    status: 'translated',
    group: '9.6',
    changes: ['repeating-item'],
    party: untrackedRule('text', 'a copy of a repeating section item'),
    note: 'A new item of a repeating section: a copy of the item before it, its controls back to their prompts, as Word adds one.',
    translate: translator('addRepeatingSectionItem', (op, batch) => () => {
      const section = controlOf(batch.doc, op.controlId);
      return [
        {
          kind: 'repeating-item',
          part: section.story.name,
          control: addressOf(section.open),
          action: 'add',
          ...(op.index === undefined ? {} : { index: op.index }),
        },
      ];
    }),
  },
  removeRepeatingSectionItem: {
    status: 'translated',
    group: '9.6',
    changes: ['repeating-item'],
    party: untrackedRule('text', 'an item of a repeating section goes'),
    note: 'Remove an item of a repeating section control; a section keeps one.',
    translate: translator('removeRepeatingSectionItem', (op, batch) => () => {
      const section = controlOf(batch.doc, op.controlId);
      return [
        {
          kind: 'repeating-item',
          part: section.story.name,
          control: addressOf(section.open),
          action: 'remove',
          index: op.index,
        },
      ];
    }),
  },
  insertBuildingBlock: {
    status: 'translated',
    group: '9.6',
    changes: ['building-block'],
    party: untrackedRule('text', "a gallery control's content is replaced"),
    note: "A building block gallery control's content from the glossary, named: every copy reads the block from the room's Word file.",
    translate: translator('insertBuildingBlock', (op, batch) => () => {
      const control = controlOf(batch.doc, op.controlId);
      const named = batch.doc.placeholders.get(op.name);
      const source = batch.gated() ?? batch.shown();
      const glossary = source ? glossaryPartOf(source) : null;
      if (!named || !sameBlocks(named, op.blocks, glossary)) {
        throw new OrderedClientRefusal(
          'ordered-refused',
          'engine-refused',
          'insertBuildingBlock: another block under that name'
        );
      }

      return [
        {
          kind: 'building-block',
          part: control.story.name,
          control: addressOf(control.open),
          name: op.name,
        },
      ];
    }),
  },
  setTextFormFieldDefault: {
    status: 'translated',
    group: '9.6',
    changes: ['form-field', 'form-text'],
    party: untrackedRule(
      'package',
      "a legacy text form field's options (`text` when its shown default changes)"
    ),
    note: "A legacy text form field's default and options: forms protection turns tracking off.",
    translate: translator('setTextFormFieldDefault', (op, batch) => () => {
      const { p, start } = formFieldAt(batch, op.paragraphId, op.fieldNodeId, 'text');
      return [
        formFieldChange(
          p,
          start,
          {
            op: 'setTextFormFieldDefault',
            text: op.text,
            ...(op.options === undefined ? {} : { options: op.options }),
            ...(op.locale === undefined ? {} : { locale: op.locale }),
          },
          changesText(batch, p.story.name, op)
        ),
      ];
    }),
  },
  commitTextFormField: {
    status: 'translated',
    group: '9.6',
    changes: ['form-text'],
    party: untrackedRule('text', "the field's result is formatted"),
    note: "A legacy text form field's result, formatted when the user leaves it.",
    translate: translator('commitTextFormField', (op, batch) => () => {
      const { p, start } = formFieldAt(batch, op.paragraphId, op.fieldNodeId, 'text');
      return [
        formFieldChange(
          p,
          start,
          { op: 'commitTextFormField', ...(op.locale === undefined ? {} : { locale: op.locale }) },
          true
        ),
      ];
    }),
  },
  setLegacyCheckbox: {
    status: 'translated',
    group: '9.6',
    changes: ['form-field'],
    party: untrackedRule('package', "a legacy check box's state"),
    note: 'A legacy check box form field on or off.',
    translate: translator('setLegacyCheckbox', (op, batch) => () => {
      const { p, start } = formFieldAt(batch, op.paragraphId, op.fieldNodeId, 'checkbox');
      return [formFieldChange(p, start, { op: 'setLegacyCheckbox', checked: op.checked }, false)];
    }),
  },
  setLegacyDropdown: {
    status: 'translated',
    group: '9.6',
    changes: ['form-field'],
    party: untrackedRule('package', "a legacy drop-down's entry (it shows as one object)"),
    note: 'A legacy drop-down form field: one of its entries.',
    translate: translator('setLegacyDropdown', (op, batch) => () => {
      const { p, start } = formFieldAt(batch, op.paragraphId, op.fieldNodeId, 'dropdown');
      return [
        formFieldChange(
          p,
          start,
          { op: 'setLegacyDropdown', selectedIndex: op.selectedIndex },
          false
        ),
      ];
    }),
  },
};

/** Options of the control ops that change their meaning. */
export const CONTROL_OPTIONS: readonly OrderedOpOption[] = [];
