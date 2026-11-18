// The legacy form field kinds of group 9.6: direct edits only, as Word fills these fields only
// under forms protection, which turns tracking off (ECMA-376 `ST_DocProtect`). Two kinds, by
// what a reader sees, so the suggestions-only policy gives a party the room's reason.

import type { OrderedEditOutcome } from '../../edits/apply.ts';
import { editFormField, type OrderedFormFieldOp } from '../../edits/form-fields.ts';
import {
  at,
  bool,
  fieldPath,
  type Fields,
  fieldsOf,
  int,
  Invalid,
  oneOf,
  plainText,
  required,
} from '../read.ts';
import { ORDERED_CHANGE_LIMITS, type OrderedChangePlace, type OrderedEditBase } from '../types.ts';
import { resolveChangePlace } from '../change-place.ts';
import { place } from '../values.ts';
import { runKindEdit } from './run.ts';
import { orderedKindSpec, type OrderedKindApply } from './spec.ts';

/**
 * An edit of the legacy form field whose first unit is at `field` (its one object, or the first
 * unit of a text field's result) that changes no text a reader sees: a check box, a drop-down, a
 * text field's options. A direct edit only.
 *
 * @public
 */
export interface OrderedFormFieldChange extends OrderedEditBase {
  readonly kind: 'form-field';
  readonly field: OrderedChangePlace;
  readonly op: OrderedFormFieldOp;
}

/**
 * An edit of a legacy text form field that changes its shown result: typing or a deletion in
 * it, a new default, its format on leaving it. A direct edit only.
 *
 * @public
 */
export interface OrderedFormTextChange extends OrderedEditBase {
  readonly kind: 'form-text';
  readonly field: OrderedChangePlace;
  readonly op: OrderedFormFieldOp;
}

type FormOpName = OrderedFormFieldOp['op'];

/**
 * The ops each kind takes. A new default is on both lists: with options only it changes the
 * package (`form-field`), and with a new shown result it changes the text (`form-text`).
 */
const STATE_OPS: readonly FormOpName[] = [
  'setLegacyCheckbox',
  'setLegacyDropdown',
  'setTextFormFieldDefault',
];
const TEXT_OPS: readonly FormOpName[] = [
  'setTextFormFieldDefault',
  'commitTextFormField',
  'insertText',
  'deleteText',
];
const LOCALE = /^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$/;

/** Word's own caps: a drop-down holds 25 entries; a text field's default is short. */
const DROPDOWN_ENTRIES = 25;
const FIELD_TEXT = 255;
const FIELD_LENGTH = 32_767;

function locale(fields: Fields): { locale?: string } {
  if (!fields.has('locale')) return {};
  const value = plainText(fields.get('locale'), fieldPath(fields, 'locale'), 2, 35);
  if (!LOCALE.test(value)) throw new Invalid('invalid-value', fieldPath(fields, 'locale'));
  return { locale: value };
}

function options(value: unknown, path: string) {
  const fields = fieldsOf(value, path, new Set(['type', 'maxLength', 'format', 'enabled']));
  return {
    type: oneOf(required(fields, 'type'), at(path, 'type'), ['regular', 'number', 'date'] as const),
    maxLength: int(required(fields, 'maxLength'), at(path, 'maxLength'), 0, FIELD_LENGTH),
    format: plainText(required(fields, 'format'), at(path, 'format'), 0, 64),
    enabled: bool(required(fields, 'enabled'), at(path, 'enabled')),
  };
}

function formFieldOp(
  value: unknown,
  path: string,
  allowed: readonly FormOpName[]
): OrderedFormFieldOp {
  const kind =
    typeof value === 'object' && value !== null
      ? Object.getOwnPropertyDescriptor(value, 'op')?.value
      : undefined;
  const name = oneOf(kind, at(path, 'op'), allowed);
  const read = (keys: readonly string[]) => fieldsOf(value, path, new Set(['op', ...keys]));
  const text = ORDERED_CHANGE_LIMITS.text;
  switch (name) {
    case 'setLegacyCheckbox': {
      const f = read(['checked']);
      return { op: name, checked: bool(required(f, 'checked'), at(path, 'checked')) };
    }
    case 'setLegacyDropdown': {
      const f = read(['selectedIndex']);
      const index = required(f, 'selectedIndex');
      return {
        op: name,
        selectedIndex: int(index, at(path, 'selectedIndex'), 0, DROPDOWN_ENTRIES - 1),
      };
    }
    case 'setTextFormFieldDefault': {
      const f = read(['text', 'options', 'locale']);
      return {
        op: name,
        text: plainText(required(f, 'text'), at(path, 'text'), 0, FIELD_TEXT),
        ...(f.has('options') ? { options: options(f.get('options'), at(path, 'options')) } : {}),
        ...locale(f),
      };
    }
    case 'commitTextFormField':
      return { op: name, ...locale(read(['locale'])) };
    case 'insertText': {
      const f = read(['offset', 'text']);
      return {
        op: name,
        offset: int(required(f, 'offset'), at(path, 'offset'), 0, text),
        text: plainText(required(f, 'text'), at(path, 'text'), 1, text),
      };
    }
    case 'deleteText': {
      const f = read(['start', 'end']);
      const start = int(required(f, 'start'), at(path, 'start'), 0, text);
      const end = int(required(f, 'end'), at(path, 'end'), 1, text);
      if (end <= start) throw new Invalid('invalid-value', at(path, 'end'), 'not after the start');
      return { op: name, start, end };
    }
  }
}

/** A form field change's payload, with its op on the kind's list. */
const readForm = (allowed: readonly FormOpName[]) => (f: Fields) => ({
  field: place(required(f, 'field'), fieldPath(f, 'field')),
  op: formFieldOp(required(f, 'op'), fieldPath(f, 'op'), allowed),
});

/** Apply a form field change: the engine's op on the view, the new field for the old one. */
function applyForm(
  change: OrderedFormFieldChange | OrderedFormTextChange,
  input: OrderedKindApply
): OrderedEditOutcome {
  // A form field edit is direct only (`editFormField` refuses another), so it settles at once
  return runKindEdit(
    input,
    (ctx) => {
      const where = resolveChangePlace(input.doc, input.story.name, change.field);
      return editFormField(ctx, where.paragraph, where.offset, change.op);
    },
    { settleDirect: true }
  );
}

export const FORM_FIELD_KIND = orderedKindSpec<OrderedFormFieldChange>({
  kind: 'form-field',
  fields: ['field', 'op'],
  read: readForm(STATE_OPS),
  touches: (change, reader) => reader.caret(change.field),
  apply: applyForm,
  policy: { text: false, properties: false, strikes: true },
  publish: 'paragraphs',
});

export const FORM_TEXT_KIND = orderedKindSpec<OrderedFormTextChange>({
  kind: 'form-text',
  fields: ['field', 'op'],
  read: readForm(TEXT_OPS),
  touches: (change, reader) => reader.caret(change.field),
  apply: applyForm,
  policy: { text: true, properties: false, strikes: true },
  publish: 'paragraphs',
});
