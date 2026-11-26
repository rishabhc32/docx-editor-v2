// The validator's readers for the picture kinds (group 9.4). Every value a change gives about
// a picture or a shape is typed and bounded here: sizes and offsets in the ranges OOXML allows
// (`ooxml-drawing-rules.ts`), a crop below the whole image, alt text as the engine's own
// metadata rule allows it, a link through `sanitizeHref`, and media by its SHA-256 (design D14).
// No XML and no bytes travel in a change.

import type {
  DrawingHorizontalReferenceFrame,
  DrawingPositionInput,
  DrawingVerticalReferenceFrame,
} from '../../../../store/package/drawing-projection.ts';
import { IMAGE_WRAP_TARGETS } from '../../../../store/package/drawing-projection.ts';
import { validateDrawingPositionInput } from '../../../../store/package/drawing-position-input.ts';
import {
  ST_COORDINATE_MAX,
  ST_COORDINATE_MIN,
  ST_POSITION_OFFSET_MAX,
  ST_POSITION_OFFSET_MIN,
  ST_POSITIVE_COORDINATE_MAX,
} from '../../../../store/package/ooxml-drawing-rules.ts';
import { isValidXmlText } from '../../../../store/package/sinks.ts';
import type { OrderedNewObject, OrderedObjectChangeOp } from '../../edits/pictures.ts';
import { at, bool, fieldsOf, int, Invalid, list, oneOf, required } from '../read.ts';
import { changeMedia } from '../resources.ts';
import { href } from '../values.ts';

/**
 * The caps on what a picture change gives. Client and server share them.
 *
 * @public
 */
export const ORDERED_OBJECT_LIMITS = Object.freeze({
  /** UTF-16 units of a title or a description (the engine's own limit). */
  metadata: 2048,
  /** Ops in one property change. */
  ops: 16,
  /** The whole image, in thousandths of a percent: a crop of two edges stays below it. */
  crop: 100_000,
});

/** The frames a horizontal offset may be relative to (`ST_RelFromH`). */
export const FRAMES_H: readonly DrawingHorizontalReferenceFrame[] = [
  'character',
  'column',
  'insideMargin',
  'leftMargin',
  'margin',
  'outsideMargin',
  'page',
  'rightMargin',
];

/** The frames a vertical offset may be relative to (`ST_RelFromV`). */
export const FRAMES_V: readonly DrawingVerticalReferenceFrame[] = [
  'bottomMargin',
  'insideMargin',
  'line',
  'margin',
  'outsideMargin',
  'page',
  'paragraph',
  'topMargin',
];

/** A size in EMU: a positive coordinate (`ST_PositiveCoordinate`). */
const extent = (value: unknown, path: string): number =>
  int(value, path, 1, ST_POSITIVE_COORDINATE_MAX);

/**
 * A title or a description: at most 2048 units, with the characters XML 1.0 allows and no
 * control characters but tab, line feed and return (the engine's `isValidXmlMetadata`).
 */
function metadataText(value: unknown, path: string): string {
  if (typeof value !== 'string') throw new Invalid('invalid-type', path, 'not a string');
  if (value.length > ORDERED_OBJECT_LIMITS.metadata)
    throw new Invalid('too-long', path, `over ${ORDERED_OBJECT_LIMITS.metadata}`);
  for (let k = 0; k < value.length; k += 1) {
    const code = value.charCodeAt(k);
    if (code < 0x20 && code !== 0x9 && code !== 0xa && code !== 0xd)
      throw new Invalid('invalid-text', path);
  }

  if (!isValidXmlText(value)) throw new Invalid('invalid-text', path);
  return value;
}

/** A new object: a picture with its media and frame, or a text box with its width. */
export function newObject(value: unknown, path: string): OrderedNewObject {
  const head = fieldsOf(
    value,
    path,
    new Set(['type', 'media', 'extent', 'title', 'description', 'link', 'lockAspect'])
  );
  const type = oneOf(required(head, 'type'), at(path, 'type'), ['picture', 'textbox'] as const);
  const sizes = fieldsOf(
    required(head, 'extent'),
    at(path, 'extent'),
    new Set(type === 'picture' ? ['cx', 'cy'] : ['cx'])
  );
  const cx = extent(required(sizes, 'cx'), at(at(path, 'extent'), 'cx'));
  if (type === 'textbox') {
    for (const key of ['media', 'title', 'description', 'link', 'lockAspect'])
      if (head.has(key)) throw new Invalid('unknown-field', at(path, key));
    return { type, extent: { cx } };
  }

  const cy = extent(required(sizes, 'cy'), at(at(path, 'extent'), 'cy'));
  return {
    type,
    media: changeMedia(required(head, 'media'), at(path, 'media')),
    extent: { cx, cy },
    ...(head.has('title') ? { title: metadataText(head.get('title'), at(path, 'title')) } : {}),
    ...(head.has('description')
      ? { description: metadataText(head.get('description'), at(path, 'description')) }
      : {}),
    ...(head.has('link') ? { link: href(head.get('link'), at(path, 'link')) } : {}),
    ...(head.has('lockAspect')
      ? { lockAspect: bool(head.get('lockAspect'), at(path, 'lockAspect')) }
      : {}),
  };
}

function crop(value: unknown, path: string) {
  const f = fieldsOf(value, path, new Set(['left', 'top', 'right', 'bottom']));
  const edge = (key: string) => int(required(f, key), at(path, key), 0, ORDERED_OBJECT_LIMITS.crop);
  const out = {
    left: edge('left'),
    top: edge('top'),
    right: edge('right'),
    bottom: edge('bottom'),
  };
  if (
    out.left + out.right >= ORDERED_OBJECT_LIMITS.crop ||
    out.top + out.bottom >= ORDERED_OBJECT_LIMITS.crop
  )
    throw new Invalid('invalid-value', path, 'a crop leaves no image');
  return out;
}

function position(value: unknown, path: string): DrawingPositionInput {
  const f = fieldsOf(
    value,
    path,
    new Set(['mode', 'horizontalEmu', 'verticalEmu', 'relativeToH', 'relativeToV'])
  );
  const mode = f.has('mode')
    ? oneOf(f.get('mode'), at(path, 'mode'), ['frame', 'simple'] as const)
    : 'frame';
  const [min, max] =
    mode === 'simple'
      ? [ST_COORDINATE_MIN, ST_COORDINATE_MAX]
      : [ST_POSITION_OFFSET_MIN, ST_POSITION_OFFSET_MAX];
  const offset = (key: string) =>
    f.has(key) ? { [key]: int(f.get(key), at(path, key), min, max) } : {};
  const out: DrawingPositionInput = {
    ...(f.has('mode') ? { mode } : {}),
    ...offset('horizontalEmu'),
    ...offset('verticalEmu'),
    ...(f.has('relativeToH')
      ? { relativeToH: oneOf(f.get('relativeToH'), at(path, 'relativeToH'), FRAMES_H) }
      : {}),
    ...(f.has('relativeToV')
      ? { relativeToV: oneOf(f.get('relativeToV'), at(path, 'relativeToV'), FRAMES_V) }
      : {}),
  };
  if (mode === 'simple' && (f.has('relativeToH') || f.has('relativeToV')))
    throw new Invalid('invalid-value', path, 'a simple position has no frames');
  if (Object.keys(out).filter((key) => key !== 'mode').length === 0)
    throw new Invalid('empty', path);
  if (!validateDrawingPositionInput(out)) throw new Invalid('invalid-value', path);
  return out;
}

function locks(value: unknown, path: string) {
  const f = fieldsOf(value, path, new Set(['select', 'move', 'resize', 'changeAspect']));
  const out: { select?: boolean; move?: boolean; resize?: boolean; changeAspect?: boolean } = {};
  for (const key of ['select', 'move', 'resize', 'changeAspect'] as const)
    if (f.has(key)) out[key] = bool(f.get(key), at(path, key));
  if (Object.keys(out).length === 0) throw new Invalid('empty', path);
  return out;
}

const OPS = ['resize', 'crop', 'position', 'wrap', 'metadata', 'locks', 'transform'] as const;

/** One op of a property change. A new image is its own kind (`object-resource`). */
function objectOp(value: unknown, path: string): OrderedObjectChangeOp {
  const kind =
    typeof value === 'object' && value !== null && !Array.isArray(value)
      ? Object.getOwnPropertyDescriptor(value, 'op')?.value
      : undefined;
  const op = oneOf(kind, at(path, 'op'), OPS);
  switch (op) {
    case 'resize': {
      const f = fieldsOf(value, path, new Set(['op', 'cx', 'cy']));
      return {
        op,
        cx: extent(required(f, 'cx'), at(path, 'cx')),
        cy: extent(required(f, 'cy'), at(path, 'cy')),
      };
    }
    case 'crop': {
      const f = fieldsOf(value, path, new Set(['op', 'crop']));
      return { op, crop: crop(required(f, 'crop'), at(path, 'crop')) };
    }
    case 'position': {
      const f = fieldsOf(value, path, new Set(['op', 'position']));
      return { op, position: position(required(f, 'position'), at(path, 'position')) };
    }
    case 'wrap': {
      const f = fieldsOf(value, path, new Set(['op', 'wrap']));
      return { op, wrap: oneOf(required(f, 'wrap'), at(path, 'wrap'), IMAGE_WRAP_TARGETS) };
    }
    case 'metadata': {
      const f = fieldsOf(value, path, new Set(['op', 'title', 'description', 'link']));
      const link = f.get('link');
      return {
        op,
        title: metadataText(required(f, 'title'), at(path, 'title')),
        description: metadataText(required(f, 'description'), at(path, 'description')),
        ...(f.has('link') ? { link: link === null ? null : href(link, at(path, 'link')) } : {}),
      };
    }
    case 'locks': {
      const f = fieldsOf(value, path, new Set(['op', 'locks']));
      return { op, locks: locks(required(f, 'locks'), at(path, 'locks')) };
    }
    case 'transform': {
      const f = fieldsOf(value, path, new Set(['op', 'action']));
      return {
        op,
        action: oneOf(required(f, 'action'), at(path, 'action'), [
          'rotateCW',
          'rotateCCW',
          'flipH',
          'flipV',
        ] as const),
      };
    }
  }
}

/** The ops of a property change: 1 to 16, in order. */
export function objectOps(value: unknown, path: string): OrderedObjectChangeOp[] {
  return list(value, path, 1, ORDERED_OBJECT_LIMITS.ops).map((entry, k) =>
    objectOp(entry, at(path, k))
  );
}
