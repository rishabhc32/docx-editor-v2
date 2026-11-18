// What a change says about a resource it adds outside the story text (design D14), and the
// validator's readers for it. A change never carries a name: the apply mints the key, the part
// name, the relationship id and the list number from the change id.

import { validDocumentPropertyWrites } from '../../../store/package/document-property-writes.ts';
import {
  ORDERED_PICTURE_MEDIA_TYPES,
  type OrderedListLevel,
  type OrderedMediaType,
  type OrderedStyleType,
} from '../model/resources.ts';
import { properties } from './values.ts';
import { at, bool, fieldsOf, int, Invalid, oneOf, plainText, required, str } from './read.ts';
import type { OrderedChangeProperty } from './types.ts';

/**
 * The caps on what a change adds outside the story text. Client and server share them. A style
 * is a typed value with every field bounded, its properties on the run and paragraph allowlists.
 * The list fields bound no reader: a new list comes from a pasted fragment's XML.
 *
 * @public
 */
export const ORDERED_RESOURCE_LIMITS = Object.freeze({
  /** Bytes of one media file (25 MiB). */
  mediaSize: 26_214_400,
  /** Bytes of one SVG file (2 MiB): it is parsed whole, so it stays small. */
  svgSize: 2_097_152,
  /** Levels of one list definition (Word's own limit). */
  listLevels: 9,
  /** UTF-16 units of a level's text (`%1.`, a bullet). */
  levelText: 32,
  /** The highest start number of a level (Word's own limit). */
  listStart: 32_767,
  /** The widest indent of a level, in twips (22 inches). */
  indent: 31_680,
  /** UTF-16 units of a style's name. */
  styleName: 128,
  /** UTF-16 units of a style id, or a font name. */
  name: 64,
});

/**
 * Media bytes a change adds: named by their SHA-256, never sent inside the change. The room
 * server stores the bytes.
 *
 * @public
 */
export interface OrderedChangeMedia {
  /** 64 lower-case hex digits. */
  readonly sha256: string;
  readonly contentType: OrderedMediaType;
  /** The size in bytes. */
  readonly size: number;
}

/** A list definition a change adds: its levels, from level 0 up. @public */
export interface OrderedChangeList {
  readonly levels: readonly OrderedListLevel[];
}

/** A style a change adds. @public */
export interface OrderedChangeStyle {
  readonly styleId: string;
  readonly type: OrderedStyleType;
  readonly name: string;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  /** `w:uiPriority`, 0 to 99. */
  readonly uiPriority?: number;
  /** `w:unhideWhenUsed`. */
  readonly unhideWhenUsed?: boolean;
  readonly paragraph?: readonly OrderedChangeProperty[];
  readonly run?: readonly OrderedChangeProperty[];
}

/**
 * A settings write a change makes: the engine's own op, or the core document properties by
 * name (`docProps/core.xml`, which Word records no change of either). @public
 */
export type OrderedChangeSetting =
  | { readonly op: 'setDocumentProtection'; readonly enforce: boolean }
  | { readonly op: 'setSectionFurnitureOptions'; readonly evenAndOddHeaders: boolean }
  | {
      readonly op: 'setDocumentProperties';
      readonly values: Readonly<Record<string, string>>;
    };

const SHA256 = /^[0-9a-f]{64}$/;
const STYLE_ID = /^[A-Za-z][A-Za-z0-9_-]{0,63}$/;

/** A SHA-256 as 64 lower-case hex digits. */
export function sha256(value: unknown, path: string): string {
  const hash = str(value, path, 64);
  if (!SHA256.test(hash)) throw new Invalid('invalid-value', path, 'not a SHA-256');
  return hash;
}

/** Media a change adds, of one of `types`: a picture's by default. */
export function changeMedia(
  value: unknown,
  path: string,
  types: readonly OrderedMediaType[] = ORDERED_PICTURE_MEDIA_TYPES
): OrderedChangeMedia {
  const fields = fieldsOf(value, path, new Set(['sha256', 'contentType', 'size']));
  const hash = sha256(required(fields, 'sha256'), at(path, 'sha256'));
  const contentType = oneOf(required(fields, 'contentType'), at(path, 'contentType'), types);
  return {
    sha256: hash,
    contentType,
    size: int(required(fields, 'size'), at(path, 'size'), 1, orderedMediaCap(contentType)),
  };
}

/** The most bytes media of `type` may have: an SVG's own cap, or the media cap. */
export function orderedMediaCap(type: OrderedMediaType): number {
  if (type === 'image/svg+xml') {
    return ORDERED_RESOURCE_LIMITS.svgSize;
  }

  return ORDERED_RESOURCE_LIMITS.mediaSize;
}

/** A style id. */
export function styleId(value: unknown, path: string): string {
  const id = str(value, path, ORDERED_RESOURCE_LIMITS.name);
  if (!STYLE_ID.test(id)) throw new Invalid('invalid-value', path, 'not a style id');
  return id;
}

/** A style a change adds. */
export function changeStyle(value: unknown, path: string): OrderedChangeStyle {
  const fields = fieldsOf(
    value,
    path,
    new Set([
      'styleId',
      'type',
      'name',
      'basedOn',
      'next',
      'link',
      'uiPriority',
      'unhideWhenUsed',
      'paragraph',
      'run',
    ])
  );
  const name = plainText(
    required(fields, 'name'),
    at(path, 'name'),
    1,
    ORDERED_RESOURCE_LIMITS.styleName
  );
  const out: {
    styleId: string;
    type: OrderedStyleType;
    name: string;
    basedOn?: string;
    next?: string;
    link?: string;
    uiPriority?: number;
    unhideWhenUsed?: boolean;
    paragraph?: readonly OrderedChangeProperty[];
    run?: readonly OrderedChangeProperty[];
  } = {
    styleId: styleId(required(fields, 'styleId'), at(path, 'styleId')),
    type: oneOf(required(fields, 'type'), at(path, 'type'), [
      'paragraph',
      'character',
      'table',
    ] as const),
    name,
  };
  for (const key of ['basedOn', 'next', 'link'] as const)
    if (fields.has(key)) out[key] = styleId(fields.get(key), at(path, key));
  if (fields.has('uiPriority'))
    out.uiPriority = int(fields.get('uiPriority'), at(path, 'uiPriority'), 0, 99);
  if (fields.has('unhideWhenUsed'))
    out.unhideWhenUsed = bool(fields.get('unhideWhenUsed'), at(path, 'unhideWhenUsed'));
  if (fields.has('paragraph'))
    out.paragraph = properties(fields.get('paragraph'), at(path, 'paragraph'), 'paragraph');
  if (fields.has('run')) out.run = properties(fields.get('run'), at(path, 'run'), 'run');
  return out;
}

/** A settings write a change makes. */
export function changeSetting(value: unknown, path: string): OrderedChangeSetting {
  const fields = fieldsOf(value, path, new Set(['op', 'enforce', 'evenAndOddHeaders', 'values']));
  const op = oneOf(required(fields, 'op'), at(path, 'op'), [
    'setDocumentProtection',
    'setSectionFurnitureOptions',
    'setDocumentProperties',
  ] as const);
  if (op === 'setDocumentProperties') {
    for (const key of ['enforce', 'evenAndOddHeaders'])
      if (fields.has(key)) throw new Invalid('unknown-field', at(path, key));
    const values = required(fields, 'values');

    // The engine's own check: known property names, XML text of at most 4096 units each
    if (!validDocumentPropertyWrites(values))
      throw new Invalid('invalid-value', at(path, 'values'));
    return { op, values: Object.fromEntries(Object.entries(values)) };
  }

  if (fields.has('values')) throw new Invalid('unknown-field', at(path, 'values'));
  if (op === 'setSectionFurnitureOptions') {
    if (fields.has('enforce')) throw new Invalid('unknown-field', at(path, 'enforce'));
    const flag = required(fields, 'evenAndOddHeaders');
    return { op, evenAndOddHeaders: bool(flag, at(path, 'evenAndOddHeaders')) };
  }

  if (fields.has('evenAndOddHeaders'))
    throw new Invalid('unknown-field', at(path, 'evenAndOddHeaders'));
  return { op, enforce: bool(required(fields, 'enforce'), at(path, 'enforce')) };
}
