// The `watermark` change kind (task 9.8): set or remove the document's watermark, as the
// engine's `setWatermark` writes it in every header a section shows
// (`edits/sections-watermark.ts`). A picture names its bytes by their SHA-256 (design D14): the
// room refuses it while its storage lacks them.

import { setWatermark, watermarkChanges } from '../../../edits/sections-watermark.ts';
import { storyNamed } from '../../../model/document.ts';
import { OrderedRefusal } from '../../../model/refusal.ts';
import type { OrderedDocument } from '../../../model/types.ts';
import {
  at,
  bool,
  fieldPath,
  type Fields,
  fieldsOf,
  Invalid,
  list,
  plainText,
  required,
  str,
  variantFields,
} from '../../read.ts';
import { changeMedia, type OrderedChangeMedia } from '../../resources.ts';
import {
  ORDERED_CHANGE_LIMITS,
  type OrderedEditBase,
  type OrderedEditChange,
} from '../../types.ts';
import { runKindEdit } from '../run.ts';
import { orderedKindSpec } from '../spec.ts';

/** A watermark's shape size on the pages of one section, in points. @public */
export interface OrderedChangeWatermarkSize {
  readonly widthPt: number;
  readonly heightPt: number;
}

/**
 * What every watermark gives: its shape size, and the size for each section that differs.
 *
 * @public
 */
export interface OrderedChangeWatermarkBox {
  readonly widthPt: number;
  readonly heightPt: number;
  /** By section index; a header takes the size of the first section that shows it. */
  readonly sectionSizes?: readonly (OrderedChangeWatermarkSize | null)[];
}

/** Word's text watermark, as the engine's `setWatermark` writes it. @public */
export interface OrderedChangeTextWatermark extends OrderedChangeWatermarkBox {
  readonly kind: 'text';
  readonly text: string;
  readonly fontFamily: string;
  /** Null for Auto, which fits the text to the shape. */
  readonly fontSizePt: number | null;
  readonly bold: boolean;
  readonly italic: boolean;
  /** `RRGGBB`, upper case. */
  readonly colorHex: string;
  /** From 0 to 1. */
  readonly opacity: number;
  readonly diagonal: boolean;
}

/** A picture watermark, named by its bytes' hash (design D14). @public */
export interface OrderedChangePictureWatermark extends OrderedChangeWatermarkBox {
  readonly kind: 'picture';
  readonly washout: boolean;
  readonly media: OrderedChangeMedia;
}

/** A watermark as the engine's `setWatermark` writes it. @public */
export type OrderedChangeWatermark = OrderedChangeTextWatermark | OrderedChangePictureWatermark;

/**
 * Set the document's watermark in every header a section shows, or remove it (`null`).
 *
 * @public
 */
export interface OrderedWatermarkChange extends OrderedEditBase {
  readonly kind: 'watermark';
  readonly watermark: OrderedChangeWatermark | null;
}

/** The editor's rules for a watermark (`docx-editor-watermark.ts`, `watermarkRefusal`). */
const TEXT_LENGTH = 256;

// A copy of `FONT_FAMILY` in `editor/docx-editor-watermark.ts`: change both together
const FONT_FAMILY = /^[^"';{}<>\\\u0000-\u001f]{1,64}$/u;
const LINE_SEPARATORS = /[\u2028\u2029]/u;
const COLOR = /^[0-9A-F]{6}$/;

/** A shape side up to the largest page side the engine writes (44 inches). */
const SIDE = 3168;

/** A finite number from `min` to `max`. */
function num(value: unknown, path: string, min: number, max: number): number {
  if (typeof value !== 'number') throw new Invalid('invalid-type', path, 'not a number');
  if (!Number.isFinite(value) || value < min || value > max)
    throw new Invalid('invalid-number', path, `not a number from ${min} to ${max}`);
  return value + 0;
}

function side(value: unknown, path: string): number {
  const read = num(value, path, 0, SIDE);
  if (read === 0) throw new Invalid('invalid-number', path, 'not above 0');
  return read;
}

function size(value: unknown, path: string): OrderedChangeWatermarkSize | null {
  if (value === null) return null;
  const fields = fieldsOf(value, path, new Set(['widthPt', 'heightPt']));
  return {
    widthPt: side(required(fields, 'widthPt'), at(path, 'widthPt')),
    heightPt: side(required(fields, 'heightPt'), at(path, 'heightPt')),
  };
}

const BOX_FIELDS = ['widthPt', 'heightPt', 'sectionSizes'];

/** Each kind of watermark, and its own fields. */
const WATERMARK_FIELDS = {
  text: [
    'text',
    'fontFamily',
    'fontSizePt',
    'bold',
    'italic',
    'colorHex',
    'opacity',
    'diagonal',
    ...BOX_FIELDS,
  ],
  picture: ['washout', 'media', ...BOX_FIELDS],
};

function readText(fields: Fields, path: string) {
  const text = plainText(required(fields, 'text'), at(path, 'text'), 1, TEXT_LENGTH);
  if (text.trim() === '' || LINE_SEPARATORS.test(text))
    throw new Invalid('invalid-text', at(path, 'text'), 'not one line of text');
  const fontFamily = str(required(fields, 'fontFamily'), at(path, 'fontFamily'), 64);
  if (!FONT_FAMILY.test(fontFamily) || fontFamily.trim() !== fontFamily)
    throw new Invalid('invalid-value', at(path, 'fontFamily'), 'not a font name');
  const fontSize = required(fields, 'fontSizePt');
  const colorHex = str(required(fields, 'colorHex'), at(path, 'colorHex'), 6);
  if (!COLOR.test(colorHex)) throw new Invalid('invalid-value', at(path, 'colorHex'), 'not RRGGBB');
  return {
    kind: 'text' as const,
    text,
    fontFamily,
    fontSizePt: fontSize === null ? null : num(fontSize, at(path, 'fontSizePt'), 1, 1000),
    bold: bool(required(fields, 'bold'), at(path, 'bold')),
    italic: bool(required(fields, 'italic'), at(path, 'italic')),
    colorHex,
    opacity: num(required(fields, 'opacity'), at(path, 'opacity'), 0, 1),
    diagonal: bool(required(fields, 'diagonal'), at(path, 'diagonal')),
  };
}

function readWatermark(value: unknown, path: string): OrderedChangeWatermark | null {
  if (value === null) {
    return null;
  }

  const { name: kind, fields } = variantFields(value, path, 'kind', WATERMARK_FIELDS);
  const box: OrderedChangeWatermarkBox = {
    widthPt: side(required(fields, 'widthPt'), at(path, 'widthPt')),
    heightPt: side(required(fields, 'heightPt'), at(path, 'heightPt')),
    ...(fields.has('sectionSizes')
      ? {
          sectionSizes: list(
            fields.get('sectionSizes'),
            at(path, 'sectionSizes'),
            0,
            ORDERED_CHANGE_LIMITS.targets
          ).map((entry, k) => size(entry, at(at(path, 'sectionSizes'), k))),
        }
      : {}),
  };
  if (kind === 'text') {
    return { ...readText(fields, path), ...box };
  }

  return {
    kind,
    washout: bool(required(fields, 'washout'), at(path, 'washout')),
    media: changeMedia(required(fields, 'media'), at(path, 'media')),
    ...box,
  };
}

/** Every header story the document holds: a watermark change writes into each one. */
function headerStories(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (const story of doc.stories.values())
    if (story.part.root.localName === 'hdr') out.push(story.name);
  return out;
}

/**
 * Whether a watermark change writes a header part: a removal with no watermark writes none, nor
 * the watermark the headers show already.
 */
function changesFurniture(change: OrderedEditChange, doc: OrderedDocument): boolean {
  if (change.kind !== 'watermark') {
    return false;
  }

  return watermarkChanges(doc, storyNamed(doc, change.part), change.watermark);
}

export const WATERMARK_SPEC = orderedKindSpec<OrderedWatermarkChange>({
  kind: 'watermark',
  fields: ['watermark'],
  read: (fields) => ({
    watermark: readWatermark(required(fields, 'watermark'), fieldPath(fields, 'watermark')),
  }),
  touches: () => undefined,
  apply(change, input) {
    // The editor writes it with tracking off, so it is the organiser's direct edit
    if (input.edit.direct !== true) {
      throw new OrderedRefusal(
        'invalid-edit',
        'a watermark is written with tracking off: the organiser sets it directly'
      );
    }

    return runKindEdit(input, (ctx) => setWatermark(ctx, change.watermark), {
      settleDirect: true,
    });
  },
  // A watermark change writes header parts: the host's check refuses a party's
  policy: { text: false, properties: false, furniture: changesFurniture },
  publish: 'story',
  stories: (_change, doc) => headerStories(doc),
  blobs: (change) => (change.watermark?.kind === 'picture' ? [change.watermark.media.sha256] : []),
});
