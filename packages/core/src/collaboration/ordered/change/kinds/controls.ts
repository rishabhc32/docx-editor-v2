// Change kinds of group 9.6 (content controls): new and removed controls, a control's properties.

import type { ContentControlLock } from '../../../../store/package/content-control-nodes.ts';
import { customNodeBindingTarget } from '../../../../store/package/custom-node-payloads.ts';
import {
  MAX_CUSTOM_NODE_LABEL_LENGTH,
  MAX_CUSTOM_NODE_PAYLOAD_LENGTH,
} from '../../../../store/store/custom-node-write-plan.ts';
import { INSERTABLE_CONTENT_CONTROL_TYPES } from '../../../../store/store/tree-op-content-controls.ts';
import type { Ctx } from '../../edits/context.ts';
import {
  deleteControl,
  insertBuildingBlock,
  insertControl,
  removeControl,
  setControlProperties,
  wrapControl,
  type OrderedNewControl,
} from '../../edits/control-structure.ts';
import { addRepeatingItem, removeRepeatingItem } from '../../edits/repeating-items.ts';
import { addCustomNodePayload } from '../../edits/custom-xml.ts';
import { shownItemAt } from '../../edits/places.ts';
import { OrderedRefusal } from '../../model/refusal.ts';
import type { OrderedItem } from '../../model/types.ts';
import {
  at,
  fieldPath,
  type Fields,
  fieldsOf,
  int,
  Invalid,
  oneOf,
  plainText,
  required,
} from '../read.ts';
import {
  ORDERED_CHANGE_LIMITS,
  type OrderedChangeAddress,
  type OrderedChangePlace,
  type OrderedEditBase,
} from '../types.ts';
import { resolveChangePlace } from '../change-place.ts';
import { address, place } from '../values.ts';
import {
  FORM_FIELD_KIND,
  FORM_TEXT_KIND,
  type OrderedFormFieldChange,
  type OrderedFormTextChange,
} from './control-form-fields.ts';
import { runKindEdit } from './run.ts';
import { orderedKindSpec, type OrderedKindApply, type OrderedKindSpec } from './spec.ts';

/** The control types a new control may have, as the engine inserts them. @public */
export type OrderedControlType = (typeof INSERTABLE_CONTENT_CONTROL_TYPES)[number];

/** A control's lock (`w:lock`). @public */
export type OrderedControlLock = ContentControlLock;

/** A new control's own properties. @public */
export interface OrderedControlProperties {
  readonly type: OrderedControlType;
  readonly tag?: string;
  /** The control's title (`w:alias`). */
  readonly alias?: string;
  readonly lock?: OrderedControlLock;
}

/** The custom-data binding of a custom node (`w:dataBinding`). @public */
export interface OrderedControlBinding {
  readonly prefixMappings: string;
  readonly xpath: string;
  /** The custom XML part's `ds:itemID`, a GUID in braces. */
  readonly storeItemId: string;
}

/**
 * A custom node's payload: what the node its binding names holds, in the custom XML store of
 * the binding's namespace. The namespace, the store's root and the node's id are the binding's.
 *
 * @public
 */
export interface OrderedCustomPayload {
  /** The text Word paints the control from. */
  readonly label: string;
  /** The payload, as its host serialized it; never parsed by the room. */
  readonly data: string;
}

/** A custom node: an inline control that holds `text`, named by its tag. @public */
export interface OrderedCustomControl {
  readonly type: 'custom';
  readonly tag: string;
  readonly text: string;
  readonly alias?: string;
  readonly lock?: Exclude<OrderedControlLock, 'unlocked'>;
  readonly binding?: OrderedControlBinding;
  /** The node the binding names, which the room writes into its store (design D14). */
  readonly payload?: OrderedCustomPayload;
}

/**
 * A new inline control at a caret: an empty control of a type, showing its type's prompt, or a
 * custom node holding its text. A direct edit only.
 *
 * @public
 */
export interface OrderedInsertControlChange extends OrderedEditBase {
  readonly kind: 'insert-control';
  readonly place: OrderedChangePlace;
  readonly control: OrderedControlProperties | OrderedCustomControl;
}

/**
 * A new inline control around the units `[from, to)` of one paragraph. Directly, it wraps them;
 * as a suggestion, the words are struck and a copy of them goes into the new control.
 *
 * Word records a control change under tracking only as range markers (MS-OI29500 §17.13.5.7)
 * and has no record for its properties (`CT_SdtPr`), so the editor writes control edits with
 * tracking off. This kind is the one exception, as the engine proposes it; every other kind of
 * the group is a direct edit, which the suggestions-only policy refuses for a party, as the
 * host's check does.
 *
 * @public
 */
export interface OrderedWrapControlChange extends OrderedEditBase {
  readonly kind: 'wrap-control';
  readonly from: OrderedChangePlace;
  readonly to: OrderedChangePlace;
  readonly control: OrderedControlProperties;
}

/** Remove a control and keep its content (Word's "Remove content control"). @public */
export interface OrderedRemoveControlChange extends OrderedEditBase {
  readonly kind: 'remove-control';
  /** The control's open item. */
  readonly control: OrderedChangeAddress;
}

/** Remove a control with its content. @public */
export interface OrderedDeleteControlChange extends OrderedEditBase {
  readonly kind: 'delete-control';
  readonly control: OrderedChangeAddress;
}

/** A control's tag, title (`w:alias`) and lock; null removes the tag or the title. @public */
export interface OrderedControlPropertyChange extends OrderedEditBase {
  readonly kind: 'control-property';
  readonly control: OrderedChangeAddress;
  readonly tag?: string | null;
  readonly alias?: string | null;
  readonly lock?: OrderedControlLock;
}

/**
 * A gallery control's content replaced by the glossary block of this name (Word's building
 * block gallery). A direct edit only: the block is read from the room's Word file.
 *
 * @public
 */
export interface OrderedBuildingBlockChange extends OrderedEditBase {
  readonly kind: 'building-block';
  readonly control: OrderedChangeAddress;
  /** The block's `w:docPartPr/w:name`. */
  readonly name: string;
}

/**
 * Add an item to a repeating section (a copy of the item before `index`, its controls back to
 * their prompts; at the end without `index`), or remove the item at `index`. A direct edit only.
 *
 * @public
 */
export interface OrderedRepeatingItemChange extends OrderedEditBase {
  readonly kind: 'repeating-item';
  /** The repeating section's open item. */
  readonly control: OrderedChangeAddress;
  readonly action: 'add' | 'remove';
  readonly index?: number;
}

/**
 * The change kinds of group 9.6, as one union.
 *
 * @public
 */
export type OrderedControlChange =
  | OrderedInsertControlChange
  | OrderedWrapControlChange
  | OrderedRemoveControlChange
  | OrderedDeleteControlChange
  | OrderedControlPropertyChange
  | OrderedBuildingBlockChange
  | OrderedFormFieldChange
  | OrderedFormTextChange
  | OrderedRepeatingItemChange;

/** The engine's longest tag or title. */
const METADATA = 4096;
const LOCKS = ['unlocked', 'sdtLocked', 'contentLocked', 'sdtContentLocked'] as const;
const NEW_LOCKS = ['sdtLocked', 'contentLocked', 'sdtContentLocked'] as const;
const GUID = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;
const CONTROL_TOKENS = ['sdt', 'sdtBlock'] as const;

const metadata = (value: unknown, path: string): string => plainText(value, path, 0, METADATA);

/** A tag or a title, or null to remove it. */
const metadataOrNull = (value: unknown, path: string): string | null =>
  value === null ? null : metadata(value, path);

const lock = (value: unknown, path: string) => oneOf(value, path, LOCKS);

/** An optional field, read when present. */
function optional<K extends string, T>(
  fields: Fields,
  key: K,
  read: (value: unknown, path: string) => T
): { [P in K]?: T } {
  if (!fields.has(key)) return {};
  const out: { [P in K]?: T } = {};
  out[key] = read(fields.get(key), fieldPath(fields, key));
  return out;
}

function binding(value: unknown, path: string): OrderedControlBinding {
  const fields = fieldsOf(value, path, new Set(['prefixMappings', 'xpath', 'storeItemId']));
  const storeItemId = plainText(required(fields, 'storeItemId'), at(path, 'storeItemId'), 38, 38);
  if (!GUID.test(storeItemId)) throw new Invalid('invalid-value', at(path, 'storeItemId'));
  return {
    prefixMappings: metadata(required(fields, 'prefixMappings'), at(path, 'prefixMappings')),
    xpath: plainText(required(fields, 'xpath'), at(path, 'xpath'), 1, METADATA),
    storeItemId,
  };
}

function controlProperties(value: unknown, path: string): OrderedControlProperties {
  const fields = fieldsOf(value, path, new Set(['type', 'tag', 'alias', 'lock']));
  return {
    type: oneOf(required(fields, 'type'), at(path, 'type'), INSERTABLE_CONTENT_CONTROL_TYPES),
    ...optional(fields, 'tag', metadata),
    ...optional(fields, 'alias', metadata),
    ...optional(fields, 'lock', lock),
  };
}

function newControl(value: unknown, path: string): OrderedControlProperties | OrderedCustomControl {
  const type =
    typeof value === 'object' && value !== null
      ? Object.getOwnPropertyDescriptor(value, 'type')?.value
      : undefined;
  if (type !== 'custom') return controlProperties(value, path);
  const fields = fieldsOf(
    value,
    path,
    new Set(['type', 'tag', 'text', 'alias', 'lock', 'binding', 'payload'])
  );
  const control: OrderedCustomControl = {
    type: 'custom',
    tag: plainText(required(fields, 'tag'), at(path, 'tag'), 1, METADATA),
    text: plainText(required(fields, 'text'), at(path, 'text'), 1, 65_536),
    ...optional(fields, 'alias', metadata),
    ...optional(fields, 'lock', (v, p) => oneOf(v, p, NEW_LOCKS)),
    ...optional(fields, 'binding', binding),
    ...optional(fields, 'payload', customPayload),
  };

  // A payload is the node a binding of the engine's own shape names
  if (
    control.payload !== undefined &&
    (control.binding === undefined || customNodeBindingTarget(control.binding) === null)
  )
    throw new Invalid('invalid-value', at(path, 'binding'), 'the binding of a custom node');
  return control;
}

function customPayload(value: unknown, path: string): OrderedCustomPayload {
  const fields = fieldsOf(value, path, new Set(['label', 'data']));
  return {
    label: plainText(required(fields, 'label'), at(path, 'label'), 0, MAX_CUSTOM_NODE_LABEL_LENGTH),
    data: plainText(required(fields, 'data'), at(path, 'data'), 0, MAX_CUSTOM_NODE_PAYLOAD_LENGTH),
  };
}

/** Run a control edit inside the caller's undo record, settled at once when it is direct. */
const runControlEdit = (input: OrderedKindApply, edit: (ctx: Ctx) => string | null) =>
  runKindEdit(input, edit, { settleDirect: true });

/** The view paragraph and offset of a place. */
const placed = (input: OrderedKindApply, where: OrderedChangePlace) =>
  resolveChangePlace(input.doc, input.story.name, where);

/** The shown control a change names by its open item. */
const controlAt = (input: OrderedKindApply, where: OrderedChangeAddress): OrderedItem =>
  shownItemAt(input.doc, input.story, where, 'open', ['sdt', 'sdtBlock']);

/** The edit of a new control from the format's form. */
function editControl(control: OrderedInsertControlChange['control']): OrderedNewControl {
  if (control.type !== 'custom') return control;
  return {
    type: 'custom',
    tag: control.tag,
    text: control.text,
    ...(control.alias === undefined ? {} : { alias: control.alias }),
    ...(control.lock === undefined ? {} : { lock: control.lock }),
    ...(control.binding === undefined ? {} : { dataBinding: control.binding }),
  };
}

/** The group's specs: one for each kind of {@link OrderedControlChange}. */
export const CONTROL_KINDS: { readonly [K in OrderedControlChange['kind']]: OrderedKindSpec } = {
  'insert-control': orderedKindSpec<OrderedInsertControlChange>({
    kind: 'insert-control',
    fields: ['place', 'control'],
    read: (f) => ({
      place: place(required(f, 'place'), fieldPath(f, 'place')),
      control: newControl(required(f, 'control'), fieldPath(f, 'control')),
    }),
    touches: (change, reader) => reader.caret(change.place),
    apply: (change, input) =>
      runControlEdit(input, (ctx) => {
        const where = placed(input, change.place);
        const { control } = change;
        const made = insertControl(ctx, where.paragraph, where.offset, editControl(control));
        if (control.type === 'custom' && control.binding && control.payload)
          addCustomNodePayload(ctx, made, control.binding, control.payload);
        return made.suggestion;
      }),
    // A prompt or a custom node's text is new text a reader sees
    policy: { text: true, properties: false },
    publish: 'paragraphs',
  }),
  'wrap-control': orderedKindSpec<OrderedWrapControlChange>({
    kind: 'wrap-control',
    fields: ['from', 'to', 'control'],
    read: (f) => ({
      from: place(required(f, 'from'), fieldPath(f, 'from')),
      to: place(required(f, 'to'), fieldPath(f, 'to')),
      control: controlProperties(required(f, 'control'), fieldPath(f, 'control')),
    }),
    touches: (change, reader) => reader.range(change.from, change.to),
    apply: (change, input) =>
      runControlEdit(input, (ctx) => {
        const from = placed(input, change.from);
        const to = placed(input, change.to);
        if (from.index !== to.index)
          throw new OrderedRefusal('invalid-edit', 'a new control around words of one paragraph');
        return wrapControl(ctx, from.paragraph, from.offset, to.offset, change.control);
      }),
    // Directly, the words stay as they were: only the package changes. As a suggestion it
    // strikes the words it wraps
    policy: { text: false, properties: false, strikes: true },
    publish: 'paragraphs',
  }),
  'remove-control': orderedKindSpec<OrderedRemoveControlChange>({
    kind: 'remove-control',
    fields: ['control'],
    read: (f) => ({ control: address(required(f, 'control'), fieldPath(f, 'control')) }),
    touches: (change, reader) => void reader.item(change.control, 'open', CONTROL_TOKENS),
    apply: (change, input) =>
      runControlEdit(input, (ctx) => removeControl(ctx, controlAt(input, change.control))),
    policy: { text: false, properties: false, strikes: true },
    publish: 'paragraphs',
  }),
  'delete-control': orderedKindSpec<OrderedDeleteControlChange>({
    kind: 'delete-control',
    fields: ['control'],
    read: (f) => ({ control: address(required(f, 'control'), fieldPath(f, 'control')) }),
    touches: (change, reader) => void reader.item(change.control, 'open', CONTROL_TOKENS),
    apply: (change, input) =>
      runControlEdit(input, (ctx) => deleteControl(ctx, controlAt(input, change.control))),
    policy: { text: true, properties: false, strikes: true },
    publish: 'paragraphs',
  }),
  'control-property': orderedKindSpec<OrderedControlPropertyChange>({
    kind: 'control-property',
    fields: ['control', 'tag', 'alias', 'lock'],
    read: (f) => {
      if (!f.has('tag') && !f.has('alias') && !f.has('lock'))
        throw new Invalid('missing-field', f.path, 'a tag, a title or a lock');
      return {
        control: address(required(f, 'control'), fieldPath(f, 'control')),
        ...optional(f, 'tag', metadataOrNull),
        ...optional(f, 'alias', metadataOrNull),
        ...optional(f, 'lock', lock),
      };
    },
    touches: (change, reader) => void reader.item(change.control, 'open', CONTROL_TOKENS),
    apply: (change, input) =>
      runControlEdit(input, (ctx) =>
        setControlProperties(ctx, controlAt(input, change.control), {
          ...(change.tag === undefined ? {} : { tag: change.tag }),
          ...(change.alias === undefined ? {} : { alias: change.alias }),
          ...(change.lock === undefined ? {} : { lock: change.lock }),
        })
      ),
    policy: { text: false, properties: true },
    publish: 'paragraphs',
  }),
  'building-block': orderedKindSpec<OrderedBuildingBlockChange>({
    kind: 'building-block',
    fields: ['control', 'name'],
    read: (f) => ({
      control: address(required(f, 'control'), fieldPath(f, 'control')),
      name: plainText(required(f, 'name'), fieldPath(f, 'name'), 1, 512),
    }),
    touches: (change, reader) => void reader.item(change.control, 'open', CONTROL_TOKENS),
    apply: (change, input) =>
      runControlEdit(input, (ctx) =>
        insertBuildingBlock(ctx, controlAt(input, change.control), change.name)
      ),
    policy: { text: true, properties: false, strikes: true },
    publish: 'paragraphs',
  }),
  'form-field': FORM_FIELD_KIND,
  'form-text': FORM_TEXT_KIND,
  'repeating-item': orderedKindSpec<OrderedRepeatingItemChange>({
    kind: 'repeating-item',
    fields: ['control', 'action', 'index'],
    read: (f) => {
      const action = oneOf(required(f, 'action'), fieldPath(f, 'action'), [
        'add',
        'remove',
      ] as const);
      if (action === 'remove' && !f.has('index'))
        throw new Invalid('missing-field', fieldPath(f, 'index'), 'the item to remove');
      return {
        control: address(required(f, 'control'), fieldPath(f, 'control')),
        action,
        ...optional(f, 'index', (v, p) => int(v, p, 0, ORDERED_CHANGE_LIMITS.targets)),
      };
    },
    touches: (change, reader) => void reader.item(change.control, 'open', CONTROL_TOKENS),
    apply: (change, input) =>
      runControlEdit(input, (ctx) => {
        const section = controlAt(input, change.control);
        return change.action === 'add'
          ? addRepeatingItem(ctx, section, change.index)
          : removeRepeatingItem(ctx, section, change.index ?? 0);
      }),
    policy: { text: true, properties: false, strikes: true },
    publish: 'paragraphs',
  }),
};
