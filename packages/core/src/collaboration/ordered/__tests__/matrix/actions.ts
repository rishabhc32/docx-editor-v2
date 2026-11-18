// The actions a matrix scenario takes, the people who take them, and the first edits every
// kind meets. Shared by the scenario sets.

import type { OrderedObjectChangeOp, OrderedParagraphOp, OrderedTableOp } from '../../index.ts';
import type { PackageExtras } from '../word-fixtures.ts';
import { aroundOf, bodyOf, extrasOf, KINDS, type MatrixKind, type Selector } from './catalogue.ts';
import type { StoryTarget } from './kinds.ts';

/** A caret in another paragraph than the action's `p`. */
export interface PointSelector extends Selector {
  readonly p: number;
}

interface ActionBase {
  /** Actions with one label are one decision of the scenario's choices. */
  readonly label: string;
  readonly author: string;
  /** `edit`: the organiser edits directly, with tracking off. */
  readonly mode?: 'edit';
  readonly p: number;
}

export type Action = ActionBase &
  (
    | { readonly kind: 'insert'; readonly at: Selector; readonly text: string }
    | {
        readonly kind: 'insertBreak';
        readonly at: Selector;
        readonly break: 'tab' | 'line' | 'page';
      }
    | { readonly kind: 'deleteRange'; readonly from: Selector; readonly to: PointSelector }
    | {
        readonly kind: 'replaceRange';
        readonly from: Selector;
        readonly to: PointSelector;
        readonly text: string;
      }
    | {
        readonly kind: 'formatRange';
        readonly from: Selector;
        readonly to: PointSelector;
        readonly prop: string;
        readonly attributes?: Readonly<Record<string, string>>;
      }
    | { readonly kind: 'enter'; readonly at: Selector }
    | { readonly kind: 'join' }
    | {
        readonly kind: 'move' | 'copy';
        readonly range: Selector;
        readonly to: { readonly p: number; readonly at: Selector };
      }
    | { readonly kind: 'paragraph'; readonly ops: readonly OrderedParagraphOp[] }
    | { readonly kind: 'insertRow'; readonly where: 'above' | 'below' }
    | { readonly kind: 'deleteRow' }
    | { readonly kind: 'insertColumn'; readonly where: 'left' | 'right' }
    | { readonly kind: 'deleteColumn' }
    | { readonly kind: 'tableProperty'; readonly op: OrderedTableOp }
    | { readonly kind: 'insertTable'; readonly rows: number; readonly cols: number }
    | { readonly kind: 'link'; readonly at: Selector; readonly url: string }
    | { readonly kind: 'control'; readonly value: string }
    | { readonly kind: 'insertNote'; readonly at: Selector; readonly note: 'footnote' | 'endnote' }
    | { readonly kind: 'pageField'; readonly at: Selector; readonly field: 'PAGE' | 'PAGE_X_OF_Y' }
    | {
        readonly kind: 'bookmark';
        readonly at: Selector;
        readonly to: Selector;
        readonly name: string;
      }
    /** A new picture (Insert Picture) at a caret. */
    | { readonly kind: 'insertPicture'; readonly at: Selector }
    /** Direct property ops on the paragraph's first object. */
    | { readonly kind: 'pictureProperty'; readonly ops: readonly OrderedObjectChangeOp[] }
    /** A new image for the paragraph's first picture (Change Picture). */
    | { readonly kind: 'replacePicture' }
    /** The paragraph's first object deleted, as the editor's Delete on a selected picture. */
    | { readonly kind: 'deletePicture' }
    /** The paragraph's first note of a kind deleted, as Delete Note (suggesting: struck). */
    | { readonly kind: 'deleteNote'; readonly note: 'footnote' | 'endnote' }
    /** The paragraph's first note of a kind converted to the other kind. */
    | { readonly kind: 'convertNote'; readonly note: 'footnote' | 'endnote' }
    /** The document's footnote properties, as the Footnote and Endnote dialog writes them. */
    | { readonly kind: 'noteProperties'; readonly numFmt: string; readonly numStart: number }
    /** The paragraph's first equation: a new linear form, or removed (`linear` null). */
    | { readonly kind: 'equation'; readonly linear: string | null }
  );

export type SetName =
  | 'single'
  | 'pairs'
  | 'mix'
  | 'nested'
  | 'presses'
  | 'structures'
  | 'trio'
  | 'pictures'
  | 'notes';

export interface Scenario {
  readonly set: SetName;
  readonly name: string;
  readonly body: string;
  readonly actions: readonly Action[];
  readonly extras: PackageExtras;
  /** The story the actions edit, when it is not the body. */
  readonly story?: StoryTarget;
  /** The kind the scenario is about (for the coverage report). */
  readonly kinds: readonly string[];
}

export const BEFORE: Selector = { find: 'Alpha ', edge: 'end' };
export const AFTER: Selector = { find: ' Omega', edge: 'start' };
export const WIDE_FROM: Selector = { find: 'Alpha ', edge: 'end', delta: -3 };
export const wideTo = (p: number): PointSelector => ({
  p,
  find: ' Omega',
  edge: 'start',
  delta: 3,
});

export interface Who {
  readonly author: string;
  readonly label: string;
  readonly mode?: 'edit';
}
export const ANA: Who = { author: 'Ana', label: 'A' };
export const BEN: Who = { author: 'Ben', label: 'B' };
export const CY: Who = { author: 'Cy', label: 'C' };
export const OLIVE: Who = { author: 'Olive', label: 'B', mode: 'edit' };

type Edit = (p: number, kind: MatrixKind) => Action | null;

/** Pictures whose copy the engine inserts whole. */
const COPYABLE = new Set([
  'picture',
  'anchor',
  'wraptb',
  'behind',
  'tight',
  'crop',
  'picfx',
  'pictable',
]);

/** First edits, by name. Null: the edit does not apply to the kind. */
export function edits(who: Who): Readonly<Record<string, Edit>> {
  return {
    'type before': (p) => ({ ...who, p, kind: 'insert', at: BEFORE, text: 'new ' }),
    'type after': (p) => ({ ...who, p, kind: 'insert', at: AFTER, text: ' new' }),
    'tab before': (p) => ({ ...who, p, kind: 'insertBreak', at: BEFORE, break: 'tab' }),
    'line break after': (p) => ({ ...who, p, kind: 'insertBreak', at: AFTER, break: 'line' }),
    'page break before': (p) => ({ ...who, p, kind: 'insertBreak', at: BEFORE, break: 'page' }),
    'type inside': (p, kind) =>
      kind.inside ? { ...who, p, kind: 'insert', at: kind.inside, text: 'X' } : null,
    'delete it': (p, kind) =>
      kind.text ? { ...who, p, kind: 'deleteRange', from: BEFORE, to: { p, ...AFTER } } : null,
    'delete across': (p) => ({ ...who, p, kind: 'deleteRange', from: WIDE_FROM, to: wideTo(p) }),
    'replace across': (p) => ({
      ...who,
      p,
      kind: 'replaceRange',
      from: WIDE_FROM,
      to: wideTo(p),
      text: 'swap',
    }),
    'bold across': (p) => ({
      ...who,
      p,
      kind: 'formatRange',
      from: WIDE_FROM,
      to: wideTo(p),
      prop: 'b',
    }),
    'italic it': (p, kind) =>
      kind.text
        ? { ...who, p, kind: 'formatRange', from: BEFORE, to: { p, ...AFTER }, prop: 'i' }
        : null,
    'Enter inside': (p, kind) =>
      kind.inside ? { ...who, p, kind: 'enter', at: kind.inside } : null,
    'Enter before': (p) => ({ ...who, p, kind: 'enter', at: BEFORE }),
    'Enter after': (p) => ({ ...who, p, kind: 'enter', at: AFTER }),
    'Backspace at start': (p) => ({ ...who, p, kind: 'join' }),
    'Delete at end': (p) => ({ ...who, p: p + 1, kind: 'join' }),
    'delete from the paragraph before': (p, kind) => ({
      ...who,
      p: p - 1,
      kind: 'deleteRange',
      from: { find: aroundOf(kind)[0], edge: 'end' },
      to: wideTo(p),
    }),
    'delete into the paragraph after': (p, kind) => ({
      ...who,
      p,
      kind: 'deleteRange',
      from: WIDE_FROM,
      to: { p: p + 1, find: aroundOf(kind)[1], edge: 'end' },
    }),
    'move it': (p, kind) =>
      kind.text && kind.text !== '￼'
        ? {
            ...who,
            p,
            kind: 'move',
            range: { find: kind.text },
            to: { p: p + 1, at: { find: aroundOf(kind)[1], edge: 'start' } },
          }
        : null,
    'copy it': (p, kind) =>
      COPYABLE.has(kind.key)
        ? {
            ...who,
            p,
            kind: 'copy',
            range: { find: '￼' },
            to: { p: p + 1, at: { find: aroundOf(kind)[1], edge: 'start' } },
          }
        : null,
    center: (p) => ({
      ...who,
      p,
      kind: 'paragraph',
      ops: [
        {
          op: 'setParagraphProperties',
          properties: [{ localName: 'jc', attributes: { val: 'center' } }],
        },
      ],
    }),
  };
}

export function scenario(
  set: SetName,
  kind: MatrixKind,
  edit: string,
  actions: readonly (Action | null)[],
  kinds: readonly string[] = [kind.key]
): Scenario | null {
  const kept = actions.filter((action): action is Action => action !== null);
  if (kept.length !== actions.length) return null;
  return {
    set,
    name: `${kind.key} | ${edit}`,
    body: bodyOf(kind),
    actions: kept,
    extras: extrasOf(kind),
    ...(kind.story ? { story: kind.story } : {}),
    kinds,
  };
}

export const present = (s: Scenario | null): s is Scenario => s !== null;

const byKey = new Map(KINDS.map((kind) => [kind.key, kind]));
export function kindNamed(key: string): MatrixKind {
  const found = byKey.get(key);
  if (!found) throw new Error(`no kind ${key}`);
  return found;
}
