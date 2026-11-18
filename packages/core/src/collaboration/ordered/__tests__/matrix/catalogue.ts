// The item kinds of the matrix: every kind Word puts in a story (`word-fixtures.ts`), with what
// the scenarios need to aim at it: its text in the engine's model, a caret inside it, the
// paragraph it sits in, and how it wraps other items when it is a container. Spike 5's 65
// kinds: the fixture's 64, and the emoji again with the caret after it.

import {
  alphaOmega,
  bodyOf as fixtureBodyOf,
  MATRIX_EXTRAS,
  MATRIX_ITEMS,
  p,
  t,
  type MatrixItem,
  type PackageExtras,
} from '../word-fixtures.ts';
import { MORE_KINDS, type StoryTarget } from './kinds.ts';

/** Where an action applies, found in the paragraph's current text. */
export interface Selector {
  readonly find: string;
  readonly nth?: number;
  /** For a caret: the start (default) or the end of the found text, plus `delta`. */
  readonly edge?: 'start' | 'end';
  readonly delta?: number;
}

export type Group =
  | 'text'
  | 'object'
  | 'marker'
  | 'container'
  | 'control'
  | 'field'
  | 'block'
  | 'table'
  | 'revision'
  | 'story';

export interface MatrixKind extends MatrixItem {
  readonly group: Group;
  /** The item's text in the engine's model, between "Alpha " and " Omega". */
  readonly text: string;
  /** A caret inside the item, for typing in it. */
  readonly inside?: Selector;
  /** The paragraph the edits apply to (default 1). */
  readonly p?: number;
  /** The texts of the paragraphs before and after the edited one. */
  readonly around?: readonly [string, string];
  /** Wraps other inline items (for the nested set). */
  readonly wrap?: (inner: string) => string;
  /** The package's other parts, when the kind needs its own (a header, a background). */
  readonly extras?: PackageExtras;
  /** The whole body, when the kind's paragraph is in another story. */
  readonly body?: string;
  /** The story the edits go to, when it is not the body. */
  readonly story?: StoryTarget;
}

const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const at = (find: string, delta: number): Selector => ({ find, edge: 'start', delta });

/** What each fixture kind adds for the scenarios, by key. */
const EXTRAS: Readonly<Record<string, Omit<MatrixKind, keyof MatrixItem>>> = {
  bold: { group: 'text', text: 'Bold', inside: at('Bold', 2) },
  face: { group: 'text', text: 'Red', inside: at('Red', 1) },
  styled: { group: 'text', text: 'Strong', inside: at('Strong', 3) },
  hidden: { group: 'text', text: 'Hidden', inside: at('Hidden', 3) },
  // A character outside the basic plane is two UTF-16 units: a caret between them is refused
  emoji: { group: 'text', text: 'Doc \u{1F4C4} here', inside: at('\u{1F4C4}', 1) },
  glow: { group: 'text', text: 'Glow', inside: at('Glow', 2) },
  rtl: { group: 'text', text: 'שלום' },
  tab: { group: 'object', text: '\t' },
  br: { group: 'object', text: '\n' },
  page: { group: 'object', text: '\f' },
  column: { group: 'object', text: '\n' },
  cr: { group: 'object', text: '\n' },
  picture: { group: 'object', text: '￼' },
  anchor: { group: 'object', text: '￼' },
  textbox: { group: 'object', text: '￼' },
  vml: { group: 'object', text: '￼' },
  sym: { group: 'object', text: '(' },
  soft: { group: 'object', text: 'co\u001Fop' },
  nbh: { group: 'object', text: 'X\u001Eray' },
  ptab: { group: 'object', text: '' },
  fnref: { group: 'object', text: '￼' },
  enref: { group: 'object', text: '￼' },
  math: { group: 'object', text: '￼' },
  ruby: { group: 'object', text: '' },
  bookmark: {
    group: 'marker',
    text: 'marked',
    inside: at('marked', 3),
    wrap: (inner) => `<w:bookmarkStart w:id="0" w:name="mark"/>${inner}<w:bookmarkEnd w:id="0"/>`,
  },
  emptymark: { group: 'marker', text: '' },
  comment: {
    group: 'marker',
    text: 'noted',
    inside: at('noted', 2),
    wrap: (inner) =>
      `<w:commentRangeStart w:id="0"/>${inner}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`,
  },
  perm: { group: 'marker', text: 'open', inside: at('open', 2) },
  proof: { group: 'marker', text: 'teh', inside: at('teh', 1) },
  lrpb: { group: 'marker', text: 'next', inside: at('next', 2) },
  link: {
    group: 'container',
    text: 'Docs',
    inside: at('Docs', 2),
    wrap: (inner) => `<w:hyperlink r:id="rIdLink">${inner}</w:hyperlink>`,
  },
  anchorlink: { group: 'container', text: 'Jump', inside: at('Jump', 2) },
  smart: {
    group: 'container',
    text: 'Paris',
    inside: at('Paris', 2),
    wrap: (inner) =>
      `<w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="place">${inner}</w:smartTag>`,
  },
  xml: {
    group: 'container',
    text: 'Acme',
    inside: at('Acme', 2),
    wrap: (inner) => `<w:customXml w:element="party">${inner}</w:customXml>`,
  },
  dir: { group: 'container', text: 'abc', inside: at('abc', 1) },
  plain: {
    group: 'control',
    text: 'Name',
    inside: at('Name', 2),
    wrap: (inner) => sdt('<w:alias w:val="Name"/><w:tag w:val="name"/><w:text/>', inner),
  },
  rich: {
    group: 'control',
    text: 'Rich',
    inside: at('Rich', 2),
    wrap: (inner) => sdt('<w:alias w:val="Rich"/><w:tag w:val="rich"/>', inner),
  },
  placeholder: { group: 'control', text: 'Click here', inside: at('Click here', 5) },
  dropdown: { group: 'control', text: 'One', inside: at('One', 1) },
  combo: { group: 'control', text: 'Blue', inside: at('Blue', 2) },
  date: { group: 'control', text: '1/2/2026', inside: at('1/2/2026', 2) },
  checkbox: { group: 'control', text: '☐' },
  pagefield: { group: 'field', text: '￼' },
  ref: { group: 'field', text: '￼' },
  simple: { group: 'field', text: '￼' },
  ins: { group: 'revision', text: 'earlier', inside: at('earlier', 3) },
  del: { group: 'revision', text: 'former', inside: at('former', 3) },
  fmt: { group: 'revision', text: 'heavy', inside: at('heavy', 2) },
  list: { group: 'block', text: 'middle' },
  heading: { group: 'block', text: 'middle' },
  formatted: { group: 'block', text: 'middle' },
  markfmt: { group: 'block', text: 'middle' },
  frame: { group: 'block', text: 'middle' },
  section: { group: 'block', text: 'middle' },
  mathpara: { group: 'block', text: '￼' },
  sdtblock: { group: 'block', text: 'middle' },
  blockprompt: { group: 'block', text: 'middle' },
  xmlblock: { group: 'block', text: 'middle' },
  empty: { group: 'block', text: 'middle', p: 2, around: ['', 'After the item.'] },
  table: { group: 'table', text: 'middle', around: ['Before the item.', 'Fee'] },
  span: { group: 'table', text: 'middle', around: ['Before the item.', 'Due'] },
  vmerge: { group: 'table', text: 'middle', around: ['Before the item.', 'Fee'] },
  nested: { group: 'table', text: 'middle', p: 2, around: ['Outer', 'Outer end'] },
  headrow: { group: 'table', text: 'middle', around: ['Before the item.', 'Row two'] },
};

function kindOf(item: MatrixItem): MatrixKind {
  const extras = EXTRAS[item.key];
  if (!extras) throw new Error(`no matrix extras for ${item.key}`);
  return { ...item, ...extras };
}

const emoji = MATRIX_ITEMS.find((item) => item.key === 'emoji');
if (!emoji) throw new Error('no emoji kind');

/** Spike 5's 65 kinds, in the spike's order. */
export const SPIKE_KINDS: readonly MatrixKind[] = MATRIX_ITEMS.flatMap((item) =>
  item.key === 'emoji'
    ? [
        { ...kindOf(item), name: 'emoji, caret inside it' },
        {
          ...kindOf(emoji),
          key: 'emojiedge',
          name: 'emoji, caret after it',
          inside: { find: '\u{1F4C4}', edge: 'end' },
        },
      ]
    : [kindOf(item)]
);

/** Every kind: the spike's 65, then the ones past them. */
export const KINDS: readonly MatrixKind[] = [...SPIKE_KINDS, ...MORE_KINDS];

/** The body of a kind's scenario. */
export function bodyOf(kind: MatrixKind): string {
  return kind.body ?? fixtureBodyOf(kind);
}

/** The package parts of a kind's scenario. */
export const extrasOf = (kind: MatrixKind): PackageExtras => kind.extras ?? MATRIX_EXTRAS;

/** The paragraph the edits apply to. */
export const paragraphOf = (kind: MatrixKind): number => kind.p ?? 1;

/** The texts of the paragraphs before and after it. */
export const aroundOf = (kind: MatrixKind): readonly [string, string] =>
  kind.around ?? ['Before the item.', 'After the item.'];

export { alphaOmega, p, t };
