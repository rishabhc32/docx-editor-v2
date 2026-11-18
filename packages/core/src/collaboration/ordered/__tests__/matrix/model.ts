// The ordered model, driven by the same actions at the same places: the places are read from
// the model's view text, which must equal the engine's before each action.

import {
  orderedPlaceAt,
  orderedViewParagraph,
  type OrderedDocument,
  type OrderedEdit,
  type OrderedItem,
  type OrderedPlace,
  type OrderedViewParagraph,
} from '../../index.ts';
import { notesStoryOf } from '../../edits/notes.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { RED_PNG } from '../word-fixtures.ts';
import type { Action } from './actions.ts';
import type { Selector } from './catalogue.ts';
import { caretOf, rangeOf } from './engine.ts';

export const BODY = '/word/document.xml';

export function viewOf(
  doc: OrderedDocument,
  story: string,
  index: number
): OrderedViewParagraph | null {
  return orderedViewParagraph(doc, story, index);
}

export const textOf = (view: OrderedViewParagraph | null): string =>
  view ? view.chars.map((c) => c.ch).join('') : '<none>';

/** A change of a kind a group added, without `v`, `id` and `mode`: applied as the room does. */
export interface ModelChange {
  readonly change: Record<string, unknown>;
}

/** The red picture of the matrix's package, as a change names it. */
const RED = {
  sha256: orderedBlobHash(RED_PNG),
  contentType: 'image/png',
  size: RED_PNG.byteLength,
} as const;

/** One pixel of GIF: a new image for Change Picture. */
export const MATRIX_GIF = Uint8Array.from(
  atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'),
  (c) => c.charCodeAt(0)
);

/** The extent the matrix's new picture takes, in EMU. */
export const NEW_PICTURE_EMU = 304_800;

/** The first object of a view (a picture, a shape, a text box), and its offset. */
export function firstObject(view: OrderedViewParagraph): { item: OrderedItem; offset: number } {
  const offset = view.chars.findIndex(
    (c) =>
      c.item.kind === 'atom' &&
      c.item.node !== null &&
      c.item.node.kind !== 'textValue' &&
      !c.item.pieces &&
      (c.item.node.kind === 'drawing' || c.item.node.localName === 'AlternateContent')
  );
  const item = view.chars[offset]?.item;
  if (!item) throw new Error('no object in the paragraph');
  return { item, offset };
}

/** The model's edit for an action, its places read from the model's view now. */
export function editFor(
  doc: OrderedDocument,
  story: string,
  action: Action,
  toggledOff = false
): OrderedEdit | ModelChange {
  const view = (index: number) => {
    const found = viewOf(doc, story, index);
    if (!found) throw new Error(`no model paragraph ${index}`);
    return found;
  };
  const placeIn = (index: number, offset: number): OrderedPlace =>
    orderedPlaceAt(view(index), offset);
  const caret = (index: number, selector: Selector): OrderedPlace =>
    placeIn(index, caretOf(textOf(view(index)), selector));
  const start = placeIn(action.p, 0);
  switch (action.kind) {
    case 'insert':
      return { kind: 'insert', at: caret(action.p, action.at), text: action.text };
    case 'insertBreak':
      return { kind: 'insertBreak', at: caret(action.p, action.at), break: action.break };
    case 'deleteRange':
      return {
        kind: 'delete',
        from: caret(action.p, action.from),
        to: caret(action.to.p, action.to),
      };
    case 'replaceRange':
      return {
        kind: 'replace',
        from: caret(action.p, action.from),
        to: caret(action.to.p, action.to),
        text: action.text,
      };
    case 'formatRange':
      return {
        kind: 'format',
        from: caret(action.p, action.from),
        to: caret(action.to.p, action.to),
        properties: [
          toggledOff
            ? { localName: action.prop, attributes: { val: '0' } }
            : {
                localName: action.prop,
                ...(action.attributes ? { attributes: action.attributes } : {}),
              },
        ],
      };
    case 'enter':
      return { kind: 'split', at: caret(action.p, action.at) };
    case 'join':
      return { kind: 'join', at: start };
    case 'move':
    case 'copy': {
      const [from, to] = rangeOf(textOf(view(action.p)), action.range);
      return {
        kind: action.kind,
        from: placeIn(action.p, from),
        to: placeIn(action.p, to),
        at: caret(action.to.p, action.to.at),
      };
    }
    case 'paragraph':
      return { kind: 'paragraph', at: start, ops: action.ops };
    case 'insertRow':
      return { kind: 'insertRow', at: start, where: action.where };
    case 'deleteRow':
      return { kind: 'deleteRow', at: start };
    case 'insertColumn':
      return { kind: 'insertColumn', at: start, where: action.where };
    case 'deleteColumn':
      return { kind: 'deleteColumn', at: start };
    case 'tableProperty':
      return { kind: 'tableProperty', at: start, op: action.op };
    case 'insertTable':
      return { kind: 'insertTable', at: start, rows: action.rows, cols: action.cols };
    case 'link':
      return { kind: 'linkTarget', at: caret(action.p, action.at), url: action.url };
    case 'control': {
      const open = view(action.p).items.find(
        (item) => item.kind === 'open' && item.token === 'sdt'
      );
      if (!open) throw new Error('no control in the paragraph');
      return {
        kind: 'controlValue',
        control: { space: open.space, offset: open.offset },
        value: action.value,
      };
    }
    case 'insertNote':
      return { kind: 'insertNote', at: caret(action.p, action.at), note: action.note };
    case 'insertPicture':
      return {
        change: {
          kind: 'insert-object',
          part: story,
          place: caret(action.p, action.at),
          object: {
            type: 'picture',
            media: RED,
            extent: { cx: NEW_PICTURE_EMU, cy: NEW_PICTURE_EMU },
          },
        },
      };
    case 'pictureProperty': {
      const { item } = firstObject(view(action.p));
      return {
        change: {
          kind: 'object-property',
          part: story,
          object: { space: item.space, offset: item.offset },
          ops: action.ops,
        },
      };
    }
    case 'replacePicture': {
      const { item } = firstObject(view(action.p));
      return {
        change: {
          kind: 'object-resource',
          part: story,
          object: { space: item.space, offset: item.offset },
          media: {
            sha256: orderedBlobHash(MATRIX_GIF),
            contentType: 'image/gif',
            size: MATRIX_GIF.byteLength,
          },
        },
      };
    }
    case 'deletePicture': {
      const { offset } = firstObject(view(action.p));
      return { kind: 'delete', from: placeIn(action.p, offset), to: placeIn(action.p, offset + 1) };
    }
    case 'deleteNote': {
      // The client's `deleteNote`: a deletion of the reference, which takes the note with it
      const { offset } = noteReferenceOf(view(action.p), action.note);
      return { kind: 'delete', from: placeIn(action.p, offset), to: placeIn(action.p, offset + 1) };
    }
    case 'convertNote': {
      const { id } = noteReferenceOf(view(action.p), action.note);
      const notes = notesStoryOf(doc, action.note);
      const open = notes
        ? [...notes.tree.items()].find(
            (item) =>
              item.kind === 'open' &&
              item.token === 'note' &&
              item.node?.kind !== 'textValue' &&
              item.node?.attributes.some((a) => a.localName === 'id' && a.value === id)
          )
        : undefined;
      if (!notes || !open) throw new Error(`no ${action.note} ${id}`);
      return {
        change: {
          kind: 'convert-note',
          part: notes.name,
          notes: [{ space: open.space, offset: open.offset }],
        },
      };
    }
    case 'noteProperties':
      return {
        change: {
          kind: 'note-properties',
          part: story,
          footnote: { numFmt: action.numFmt, numStart: action.numStart },
        },
      };
    case 'equation': {
      const at = view(action.p).chars.findIndex(
        (c) =>
          c.item.kind === 'atom' &&
          c.item.bare === true &&
          c.item.node?.kind !== 'textValue' &&
          /^oMath(Para)?$/.test(c.item.node?.localName ?? '')
      );
      const item = view(action.p).chars[at]?.item;
      if (!item) throw new Error('no equation in the paragraph');
      // A lone equation goes by a deletion of its object, as the client's `removeMathEquation`
      if (action.linear === null)
        return { kind: 'delete', from: placeIn(action.p, at), to: placeIn(action.p, at + 1) };
      return {
        change: {
          kind: 'equation',
          part: story,
          object: { space: item.space, offset: item.offset },
          equation: 0,
          linear: action.linear,
        },
      };
    }
    case 'pageField':
      return {
        change: {
          kind: 'insert-field',
          part: story,
          place: caret(action.p, action.at),
          field: action.field,
        },
      };
    case 'bookmark':
      return {
        change: {
          kind: 'bookmark',
          part: story,
          name: action.name,
          from: caret(action.p, action.at),
          to: caret(action.p, action.to),
          remove: [],
        },
      };
  }
}

/** A view's first note reference of a kind: its offset and its note's `w:id`. */
function noteReferenceOf(
  view: OrderedViewParagraph,
  kind: 'footnote' | 'endnote'
): { offset: number; id: string } {
  const offset = view.chars.findIndex(
    (c) => c.item.node?.kind !== 'textValue' && c.item.node?.localName === `${kind}Reference`
  );
  const node = view.chars[offset]?.item.node;
  const id =
    node && node.kind !== 'textValue'
      ? node.attributes.find((a) => a.localName === 'id')?.value
      : undefined;
  if (id === undefined) throw new Error(`no ${kind} reference in the paragraph`);
  return { offset, id };
}
