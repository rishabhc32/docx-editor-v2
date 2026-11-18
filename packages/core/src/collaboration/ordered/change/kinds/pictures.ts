// Change kinds of group 9.4 (pictures and shapes): a new picture or text box (`insert-object`),
// a direct change of an object's size, crop, position, wrap, rotation, alt text, locks or link
// (`object-property`), and a new image for a picture (`object-resource`). A deletion of an
// object is the format's `delete` over its unit. The group's own file: its change types and
// their specs, and nothing else edits it (design, "How to add a group").

import {
  changeObject,
  insertObject,
  type OrderedNewObject,
  type OrderedObjectChangeOp,
} from '../../edits/pictures.ts';
import { shownItemAt } from '../../edits/places.ts';
import type { OrderedItem } from '../../model/types.ts';
import { placeOfChange } from '../change-place.ts';
import { fieldPath, required } from '../read.ts';
import { changeMedia, type OrderedChangeMedia } from '../resources.ts';
import type { OrderedChangeAddress, OrderedChangePlace, OrderedEditBase } from '../types.ts';
import { address, place } from '../values.ts';
import { newObject, objectOps } from './pictures-read.ts';
import { runKindEdit } from './run.ts';
import { orderedKindSpec, type OrderedKindApply, type OrderedKindSpec } from './spec.ts';

/**
 * A new picture or text box before a place. A suggestion when its mode is `suggest`.
 *
 * @public
 */
export interface OrderedInsertObjectChange extends OrderedEditBase {
  readonly kind: 'insert-object';
  readonly place: OrderedChangePlace;
  readonly object: OrderedNewObject;
}

/**
 * A direct change of an object's properties: the ops, in order, on the object whose unit is at
 * `object`. Word has no tracked form for it: a change in `suggest` mode is refused
 * (`untracked-only`).
 *
 * @public
 */
export interface OrderedObjectPropertyChange extends OrderedEditBase {
  readonly kind: 'object-property';
  readonly object: OrderedChangeAddress;
  readonly ops: readonly OrderedObjectChangeOp[];
}

/**
 * A picture's new image, named by its SHA-256, in the same frame. A direct edit, as
 * `object-property`.
 *
 * @public
 */
export interface OrderedObjectResourceChange extends OrderedEditBase {
  readonly kind: 'object-resource';
  readonly object: OrderedChangeAddress;
  readonly media: OrderedChangeMedia;
}

/**
 * The change kinds of group 9.4.
 *
 * @public
 */
export type OrderedPictureChange =
  | OrderedInsertObjectChange
  | OrderedObjectPropertyChange
  | OrderedObjectResourceChange;

/** The shown object's unit at an address. */
const objectAt = (input: OrderedKindApply, at: OrderedChangeAddress): OrderedItem =>
  shownItemAt(input.doc, input.story, at, 'atom');

const insertObjectSpec = orderedKindSpec<OrderedInsertObjectChange>({
  kind: 'insert-object',
  fields: ['place', 'object'],
  read: (f) => ({
    place: place(required(f, 'place'), fieldPath(f, 'place')),
    object: newObject(required(f, 'object'), fieldPath(f, 'object')),
  }),
  touches: (change, reader) => reader.caret(change.place),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => insertObject(ctx, placeOfChange(change.place), change.object)),
  // A new object is a tracked insertion of the run that holds it, and a deletion a tracked
  // deletion (Word 2016 files: `w:ins` and `w:del` around `w:r` with `w:drawing`). A party's
  // untracked picture changes the text a reader sees: one more object in it
  policy: { text: true, properties: false, strikes: false },
  publish: 'paragraphs',
  caret: (change) => change.place,
  types: true,
  blobs: (change) => (change.object.type === 'picture' ? [change.object.media.sha256] : []),
});

const objectPropertySpec = orderedKindSpec<OrderedObjectPropertyChange>({
  kind: 'object-property',
  fields: ['object', 'ops'],
  read: (f) => ({
    object: address(required(f, 'object'), fieldPath(f, 'object')),
    ops: objectOps(required(f, 'ops'), fieldPath(f, 'ops')),
  }),
  touches: (change, reader) => void reader.item(change.object, 'atom'),
  apply: (change, input) =>
    runKindEdit(input, (ctx) => changeObject(ctx, objectAt(input, change.object), change.ops)),
  // Word has no tracked form for an object's properties or its image (`model/objects.ts`): a
  // party's is refused as the host's check refuses it (`policy-direct`, reason `package`). A
  // property change strikes nothing: inside another person's struck object it is inside
  policy: { text: false, properties: true, strikes: false },
  publish: 'paragraphs',
});

const objectResourceSpec = orderedKindSpec<OrderedObjectResourceChange>({
  kind: 'object-resource',
  fields: ['object', 'media'],
  read: (f) => ({
    object: address(required(f, 'object'), fieldPath(f, 'object')),
    media: changeMedia(required(f, 'media'), fieldPath(f, 'media')),
  }),
  touches: (change, reader) => void reader.item(change.object, 'atom'),
  apply: (change, input) =>
    runKindEdit(input, (ctx) =>
      changeObject(ctx, objectAt(input, change.object), [{ op: 'image', media: change.media }])
    ),
  // A direct edit only, as `object-property`. A property change strikes nothing: inside another
  // person's struck object it is inside
  policy: { text: false, properties: true, strikes: false },
  publish: 'paragraphs',
  blobs: (change) => [change.media.sha256],
});

/** The group's specs: one for each kind of {@link OrderedPictureChange}. */
export const PICTURE_KINDS: { readonly [K in OrderedPictureChange['kind']]: OrderedKindSpec } = {
  'insert-object': insertObjectSpec,
  'object-property': objectPropertySpec,
  'object-resource': objectResourceSpec,
};
