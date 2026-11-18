// The matrix scenarios (spike 5's `matrix-scenarios.ts`, then more), generated from the kinds:
//
//   single:     every kind x every first edit, one author (Ana, suggesting).
//   pairs:      every kind x a second edit over the first: Ben suggesting, or Olive directly.
//   mix:        two kinds side by side, with edits at the place where they meet.
//   nested:     a kind inside a container (link, control, bookmark, comment, ...), edits around it.
//   presses:    paragraph presses (alignment, spacing, indents, tabs, keeps, styles, lists).
//   structures: rows, columns, table presses, new tables, links, control values, notes, copies.
//   trio:       mixed content edited by three people, decided in every order.
//
// A scenario's name is "<kind> | <edit>": the report splits it into rows and columns.

import {
  AFTER,
  ANA,
  BEFORE,
  BEN,
  edits,
  kindNamed,
  OLIVE,
  present,
  scenario,
  WIDE_FROM,
  type Scenario,
  type SetName,
  type Who,
} from './actions.ts';
import { KINDS, paragraphOf, t, type MatrixKind, type Selector } from './catalogue.ts';
import { NOTE_SET } from './notes.ts';
import { PICTURE_SET } from './pictures.ts';
import { PRESSES, STRUCTURES, TRIO } from './sets.ts';

export type { Action, Scenario, SetName } from './actions.ts';

const SINGLE: readonly Scenario[] = KINDS.flatMap((kind) =>
  Object.entries(edits(ANA))
    .map(([name, edit]) => scenario('single', kind, name, [edit(paragraphOf(kind), kind)]))
    .filter(present)
);

/** A second author's edit over the first: [name, first edit, second edit, second author]. */
const PAIRS: readonly (readonly [string, string, string, Who])[] = [
  ['Ana types after, Ben deletes across', 'type after', 'delete across', BEN],
  ['Ana deletes it, Ben types inside', 'delete it', 'type inside', BEN],
  ['Ana deletes it, Ben types before', 'delete it', 'type before', BEN],
  ['Ana bolds, Ben italics it', 'bold across', 'italic it', BEN],
  ['Ana types inside, Ben deletes it', 'type inside', 'delete it', BEN],
  ['Ana deletes across, Ben deletes it', 'delete across', 'delete it', BEN],
  ['Ana deletes it, Ben replaces across', 'delete it', 'replace across', BEN],
  ['Ana deletes across, Olive types after', 'delete across', 'type after', OLIVE],
  ['Ana types after, Olive deletes across', 'type after', 'delete across', OLIVE],
  ['Ana bolds, Olive deletes it', 'bold across', 'delete it', OLIVE],
  // A direct run format after a pending one: Word keeps one `w:rPrChange` (task 4.4)
  ['Ana bolds, Olive italics it', 'bold across', 'italic it', OLIVE],
];

const PAIRS_SET: readonly Scenario[] = KINDS.flatMap((kind) =>
  PAIRS.map(([name, first, second, who]) => {
    const p = paragraphOf(kind);
    const one = edits(ANA)[first];
    const two = edits(who)[second];
    if (!one || !two) throw new Error(`no edit ${first} or ${second}`);
    return scenario('pairs', kind, name, [one(p, kind), two(p, kind)]);
  }).filter(present)
);

// ---- two kinds together ----

/** Inline kinds the mix puts side by side: the spike's 22, then the run formats and objects. */
export const MIX_KEYS = [
  'bold',
  'hidden',
  'tab',
  'br',
  'picture',
  'anchor',
  'textbox',
  'fnref',
  'math',
  'ptab',
  'bookmark',
  'emptymark',
  'comment',
  'link',
  'plain',
  'rich',
  'dropdown',
  'checkbox',
  'pagefield',
  'simple',
  'ins',
  'del',
  'udouble',
  'font',
  'sup',
  'complex',
  'crop',
  'shape',
  'datefield',
  'formtext',
  'dropprompt',
  'piccontrol',
];

const MIX = MIX_KEYS.map(kindNamed);

const together = (x: MatrixKind, y: MatrixKind): MatrixKind => ({
  key: `${x.key}+${y.key}`,
  name: `${x.name}, then ${y.name}`,
  group: 'text',
  xml: (x.xml ?? '') + (y.xml ?? ''),
  text: x.text + y.text,
});

const after = (n: number): Selector => ({ find: 'Alpha ', edge: 'end', delta: n });

const MIX_SET: readonly Scenario[] = MIX.flatMap((x) =>
  MIX.filter((y) => y !== x).flatMap((y) => {
    const both = together(x, y);
    const n = x.text.length;
    const p = 1;
    const keys = [x.key, y.key];
    return [
      scenario(
        'mix',
        both,
        'type between',
        [{ ...ANA, p, kind: 'insert', at: after(n), text: 'mid' }],
        keys
      ),
      both.text
        ? scenario(
            'mix',
            both,
            'delete both',
            [{ ...ANA, p, kind: 'deleteRange', from: BEFORE, to: { p, ...AFTER } }],
            keys
          )
        : null,
      scenario('mix', both, 'Enter between', [{ ...ANA, p, kind: 'enter', at: after(n) }], keys),
      n > 0 && y.text.length > 0
        ? scenario(
            'mix',
            both,
            'delete where they meet',
            [{ ...ANA, p, kind: 'deleteRange', from: after(n - 1), to: { p, ...after(n + 1) } }],
            keys
          )
        : null,
    ].filter(present);
  })
);

// ---- a kind inside a container ----

const CONTAINERS = [
  'link',
  'bookmark',
  'comment',
  'smart',
  'xml',
  'plain',
  'rich',
  'group',
  'bound',
  'gallery',
].map(kindNamed);

const NESTED_SET: readonly Scenario[] = CONTAINERS.flatMap((y) =>
  MIX.filter((x) => x !== y && x.group !== 'revision').flatMap((x) => {
    const wrap = y.wrap ?? ((inner: string) => (y.xml ?? '').replace(t(y.text), inner));
    const inner: MatrixKind = {
      key: `${x.key} in ${y.key}`,
      name: `${x.name} inside ${y.name}`,
      group: 'container',
      xml: wrap(t('in ') + (x.xml ?? '') + t(' it')),
      text: `in ${x.text} it`,
    };
    const p = 1;
    const IN: Selector = { find: 'in ', edge: 'end' };
    const IT: Selector = { find: ' it', edge: 'start' };
    const keys = [x.key, y.key];
    return [
      scenario(
        'nested',
        inner,
        'type before it',
        [{ ...ANA, p, kind: 'insert', at: IN, text: 'X' }],
        keys
      ),
      scenario(
        'nested',
        inner,
        'type after it',
        [{ ...ANA, p, kind: 'insert', at: IT, text: 'X' }],
        keys
      ),
      // Enter inside the container, right before the kind: it splits the container (task 3.5)
      scenario('nested', inner, 'Enter before it', [{ ...ANA, p, kind: 'enter', at: IN }], keys),
      x.text
        ? scenario(
            'nested',
            inner,
            'delete it',
            [{ ...ANA, p, kind: 'deleteRange', from: IN, to: { p, ...IT } }],
            keys
          )
        : null,
      scenario(
        'nested',
        inner,
        'delete the container',
        [{ ...ANA, p, kind: 'deleteRange', from: BEFORE, to: { p, ...AFTER } }],
        keys
      ),
      scenario(
        'nested',
        inner,
        'delete across its start',
        [
          {
            ...ANA,
            p,
            kind: 'deleteRange',
            from: WIDE_FROM,
            to: { p, find: 'in ', edge: 'end', delta: x.text.length },
          },
        ],
        keys
      ),
      scenario(
        'nested',
        inner,
        'Ana types in it, Ben deletes the container',
        [
          { ...ANA, p, kind: 'insert', at: IN, text: 'X' },
          { ...BEN, p, kind: 'deleteRange', from: BEFORE, to: { p, ...AFTER } },
        ],
        keys
      ),
    ].filter(present);
  })
);

export const SETS: Readonly<Record<SetName, readonly Scenario[]>> = {
  single: SINGLE,
  pairs: PAIRS_SET,
  mix: MIX_SET,
  nested: NESTED_SET,
  presses: PRESSES,
  structures: STRUCTURES,
  trio: TRIO,
  pictures: PICTURE_SET,
  notes: NOTE_SET,
};

/** Every n-th scenario of a set from the i-th (from 0): one shard of a set. */
export function shardOf(set: SetName, index: number, count: number): readonly Scenario[] {
  return SETS[set].filter((_, k) => k % count === index);
}
