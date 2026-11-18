// The matrix's picture scenarios (group 9.4): a new picture next to every kind, tracked (Ana)
// and direct (Olive, the organiser); every property edit the editor makes on every picture kind,
// direct; Change Picture; a deletion, tracked and direct; and two people on one picture.

import type { OrderedObjectChangeOp } from '../../index.ts';
import {
  AFTER,
  ANA,
  BEFORE,
  BEN,
  OLIVE,
  present,
  scenario,
  type Scenario,
  type Who,
} from './actions.ts';
import { KINDS, paragraphOf, type MatrixKind } from './catalogue.ts';

/** Kinds that hold a picture (a `pic:pic`), inline or floating. */
const PICTURES = [
  'picture',
  'anchor',
  'wraptb',
  'behind',
  'tight',
  'crop',
  'picfx',
  'pictable',
  'piccontrol',
  'picheader',
];

/** Of those, the floating ones, which have a position. */
const FLOATING = new Set(['anchor', 'wraptb', 'behind', 'tight']);

/** Property edits by name, with the kinds they apply to. */
const PROPERTIES: readonly (readonly [
  string,
  readonly OrderedObjectChangeOp[],
  (key: string) => boolean,
])[] = [
  ['resize it', [{ op: 'resize', cx: 228_600, cy: 152_400 }], () => true],
  [
    'crop it',
    [{ op: 'crop', crop: { left: 10_000, top: 0, right: 5_000, bottom: 20_000 } }],
    () => true,
  ],
  ['wrap it square', [{ op: 'wrap', wrap: 'square' }], (key) => !FLOATING.has(key)],
  ['make it inline', [{ op: 'wrap', wrap: 'inline' }], (key) => FLOATING.has(key)],
  [
    'describe it',
    [{ op: 'metadata', title: 'Red', description: 'A red dot\non two lines' }],
    () => true,
  ],
  ['turn it', [{ op: 'transform', action: 'rotateCW' }], () => true],
  ['flip it', [{ op: 'transform', action: 'flipH' }], () => true],
  ['lock its aspect', [{ op: 'locks', locks: { changeAspect: true } }], () => true],
  [
    'move it',
    [{ op: 'position', position: { horizontalEmu: 12_700, relativeToH: 'page' } }],
    (key) => FLOATING.has(key),
  ],
  [
    'resize and crop it',
    [
      { op: 'resize', cx: 228_600, cy: 228_600 },
      { op: 'crop', crop: { left: 0, top: 10_000, right: 0, bottom: 0 } },
    ],
    () => true,
  ],
];

const insert = (who: Who, p: number, after = false) => ({
  ...who,
  p,
  kind: 'insertPicture' as const,
  at: after ? AFTER : BEFORE,
});
const property = (who: Who, p: number, ops: readonly OrderedObjectChangeOp[]) => ({
  ...who,
  p,
  kind: 'pictureProperty' as const,
  ops,
});
const replace = (who: Who, p: number) => ({ ...who, p, kind: 'replacePicture' as const });
const remove = (who: Who, p: number) => ({ ...who, p, kind: 'deletePicture' as const });
const RESIZE: readonly OrderedObjectChangeOp[] = [{ op: 'resize', cx: 228_600, cy: 228_600 }];

function forKind(kind: MatrixKind): (Scenario | null)[] {
  const p = paragraphOf(kind);
  const out: (Scenario | null)[] = [
    scenario('pictures', kind, 'a picture before', [insert(ANA, p)]),
    scenario('pictures', kind, 'a picture after, directly', [insert(OLIVE, p, true)]),
    scenario('pictures', kind, 'Ana adds a picture, Olive resizes it', [
      insert(ANA, p),
      property(OLIVE, p, RESIZE),
    ]),
    scenario('pictures', kind, 'Ana adds a picture, Ben deletes it', [
      insert(ANA, p),
      remove(BEN, p),
    ]),
  ];
  if (!PICTURES.includes(kind.key)) return out;
  for (const [name, ops, applies] of PROPERTIES)
    if (applies(kind.key)) out.push(scenario('pictures', kind, name, [property(OLIVE, p, ops)]));
  out.push(
    scenario('pictures', kind, 'a new image', [replace(OLIVE, p)]),
    scenario('pictures', kind, 'delete it', [remove(ANA, p)]),
    scenario('pictures', kind, 'delete it directly', [remove(OLIVE, p)]),
    scenario('pictures', kind, 'Ana deletes it, Olive resizes it', [
      remove(ANA, p),
      property(OLIVE, p, RESIZE),
    ]),
    scenario('pictures', kind, 'Olive resizes it, Ana deletes it', [
      property(OLIVE, p, RESIZE),
      remove(ANA, p),
    ]),
    scenario('pictures', kind, 'Olive gives it a new image, Ben deletes it', [
      replace(OLIVE, p),
      remove(BEN, p),
    ])
  );
  return out;
}

/** The text box: Word's shape with a story inside; its frame and alt text, and a deletion. */
function textBox(kind: MatrixKind): (Scenario | null)[] {
  const p = paragraphOf(kind);
  return [
    scenario('pictures', kind, 'resize it', [property(OLIVE, p, RESIZE)]),
    scenario('pictures', kind, 'describe it', [
      property(OLIVE, p, [{ op: 'metadata', title: 'Box', description: 'A note' }]),
    ]),
    scenario('pictures', kind, 'delete it directly', [remove(OLIVE, p)]),
  ];
}

export const PICTURE_SET: readonly Scenario[] = KINDS.flatMap((kind) =>
  [...forKind(kind), ...(kind.key === 'textbox' ? textBox(kind) : [])].filter(present)
);
