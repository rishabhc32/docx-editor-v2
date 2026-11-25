// A content control's lock in a room (security). The room refuses every change that a control's
// `w:lock` forbids (ECMA-376 Part 1 §17.5.2.23), by the organiser and by a party alike, whatever
// a client sends: a modified browser can send any change the format allows. The editor alone is
// the reference (the store's `locked`, `contentControlLockRefusal`): `contentLocked` and
// `sdtContentLocked` forbid every edit of the content, `sdtLocked` and `sdtContentLocked` forbid
// the control's removal. The room refuses by the same name (`apply-refused`, reason `locked`).
// This file runs every change kind against each lock; `room-locked-places.test.ts` checks the
// places around a control, the acceptance run's case and the objects inside one.
//
// Found by the acceptance harness (s100-short seed 1007, ana-1:386): a party typed inside a
// content-locked combo box and the room accepted it.

import { describe, expect, test } from 'bun:test';
import { ORDERED_CHANGE_KINDS } from '../change/kinds/registry.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, at, markOf, opens } from './change-apply-fixture.ts';
import { bodyFingerprint } from './room-harness.ts';
import {
  BOTH,
  content,
  control,
  edit,
  fileOf,
  firstOf,
  LOCKS,
  MEDIA,
  one,
  removal,
  roomOn,
  type Layout,
  type Lock,
  type Who,
} from './room-locked-fixture.ts';

/** Where a case reaches into the locked content: its paragraph and two offsets inside it. */
const INSIDE: Readonly<Record<'inline' | 'block', { line: number; from: number; to: number }>> = {
  inline: { line: 0, from: 8, to: 10 },
  block: { line: 1, from: 2, to: 5 },
};

interface Case {
  readonly name: string;
  readonly layouts: readonly Layout[];
  readonly who: readonly Who[];
  readonly make: (doc: OrderedDocument, who: Who, layout: Layout) => Record<string, unknown>;
  /** When the lock refuses it; by default when the lock forbids editing the content. */
  readonly when?: (lock: Lock, layout: Layout, who: Who) => boolean;
  /** The answer when the lock does not refuse it, when it is not `accepted`. */
  readonly free?: string;
  /** The change kind, when not the name's first word. */
  readonly kind: string;
}

const inside = (doc: OrderedDocument, layout: Layout, end = false) => {
  const i = INSIDE[layout === 'block' ? 'block' : 'inline'];
  return at(doc, i.line, end ? i.to : i.from);
};
const ORGANISER: readonly Who[] = ['organiser'];
const TEXT_LAYOUTS: readonly Layout[] = ['inline', 'block'];

const CASES: readonly Case[] = [
  {
    name: 'typing inside',
    kind: 'insert',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert',
      ...edit(who),
      place: inside(doc, layout),
      text: 'Q',
    }),
  },
  {
    name: 'a tab inside',
    kind: 'insert-tab',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({ kind: 'insert-tab', ...edit(who), place: inside(doc, layout) }),
  },
  {
    name: 'a line break inside',
    kind: 'insert-break',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-break',
      ...edit(who),
      place: inside(doc, layout),
      break: 'line',
    }),
  },
  {
    name: 'a deletion inside',
    kind: 'delete',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'delete',
      ...edit(who),
      from: inside(doc, layout),
      to: inside(doc, layout, true),
    }),
  },
  {
    name: 'a deletion from outside into the control',
    kind: 'delete',
    layouts: ['inline'],
    who: BOTH,
    make: (doc, who) => ({ kind: 'delete', ...edit(who), from: at(doc, 0, 4), to: at(doc, 0, 9) }),
  },
  {
    name: 'a deletion of the whole control and words around it',
    kind: 'delete',
    layouts: ['inline'],
    who: BOTH,
    // The organiser's takes a content-locked control whole, as the editor alone does
    when: (lock, _layout, who) => content(lock) && (who === 'party' || removal(lock)),
    make: (doc, who) => ({ kind: 'delete', ...edit(who), from: at(doc, 0, 6), to: at(doc, 0, 11) }),
  },
  {
    name: 'a deletion from the paragraph before into the control',
    kind: 'delete',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({ kind: 'delete', ...edit(who), from: at(doc, 0, 3), to: at(doc, 1, 3) }),
  },
  {
    name: 'a format inside',
    kind: 'format',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'format',
      ...edit(who),
      from: inside(doc, layout),
      to: inside(doc, layout, true),
      properties: [{ localName: 'b' }],
    }),
  },
  {
    name: 'Enter inside',
    kind: 'split',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({ kind: 'split', ...edit(who), place: inside(doc, layout) }),
    // An inline rich text control refuses Enter inside it for its own reason (D5 policy 5)
    free: 'apply-refused/single-line-control',
  },
  {
    name: 'Backspace joining two paragraphs of the control',
    kind: 'join',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({ kind: 'join', ...edit(who), mark: markOf(doc, 1) }),
  },
  {
    name: "a paragraph mark's format",
    kind: 'mark-format',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'mark-format',
      ...edit(who),
      marks: [markOf(doc, 1)],
      properties: [{ localName: 'b' }],
    }),
  },
  {
    name: "a paragraph's properties",
    kind: 'paragraph-property',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'paragraph-property',
      ...edit(who),
      marks: [markOf(doc, 1)],
      properties: [{ localName: 'jc', attributes: { val: 'center' } }],
    }),
  },
  {
    name: 'a section break',
    kind: 'section-break',
    layouts: ['block'],
    who: ORGANISER,
    make: (doc, who) => ({
      kind: 'section-break',
      ...edit(who),
      mark: markOf(doc, 1),
      type: 'continuous',
    }),
  },
  {
    name: 'a list level',
    kind: 'list-level',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'list-level',
      ...edit(who),
      numId: '1',
      level: 1,
      format: 'ordered',
      marks: [markOf(doc, 1)],
    }),
    // The fixture's paragraphs are in no list, so only the lock can answer before the engine
    // (`editor-list-level.test.ts` checks a free press)
    free: 'apply-refused/engine-refused',
  },
  {
    name: 'a new list',
    kind: 'new-list',
    layouts: ['block'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'new-list',
      ...edit(who),
      list: {
        abstract: {
          e: 'w:abstractNum',
          c: [
            {
              e: 'w:lvl',
              a: { 'w:ilvl': '0' },
              c: [
                { e: 'w:numFmt', a: { 'w:val': 'decimal' } },
                { e: 'w:lvlText', a: { 'w:val': '%1.' } },
              ],
            },
          ],
        },
      },
      paragraphs: [{ mark: markOf(doc, 1), level: 0 }],
    }),
  },
  {
    name: 'a new table at a place inside',
    kind: 'insert-table',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-table',
      ...edit(who),
      place: inside(doc, layout),
      rows: 1,
      columns: 1,
      columnWidth: 2400,
    }),
    // The table goes before the paragraph, which an inline control is in, not around
    when: (lock, layout) => content(lock) && layout === 'block',
  },
  {
    name: 'a move out of the control',
    kind: 'move',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'move',
      ...edit(who),
      from: inside(doc, layout),
      to: inside(doc, layout, true),
      place: at(doc, layout === 'block' ? 3 : 1, 0),
    }),
  },
  {
    name: 'a move into the control',
    kind: 'move',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'move',
      ...edit(who),
      from: at(doc, layout === 'block' ? 3 : 1, 0),
      to: at(doc, layout === 'block' ? 3 : 1, 5),
      place: inside(doc, layout),
    }),
  },
  {
    name: 'a note inside',
    kind: 'insert-note',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-note',
      ...edit(who),
      place: inside(doc, layout),
      note: 'footnote',
      text: ' N',
    }),
  },
  {
    name: 'a page field inside',
    kind: 'insert-field',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-field',
      ...edit(who),
      place: inside(doc, layout),
      field: 'PAGE',
    }),
  },
  {
    name: 'a bookmark inside',
    kind: 'bookmark',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who, layout) => ({
      kind: 'bookmark',
      ...edit(who),
      name: 'Inside_1',
      from: inside(doc, layout),
      to: inside(doc, layout, true),
      remove: [],
    }),
  },
  {
    name: 'a paste inside',
    kind: 'insert-fragment',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-fragment',
      ...edit(who),
      place: inside(doc, layout),
      blocks: [one('P')],
    }),
  },
  {
    name: 'a link over words inside',
    kind: 'insert-link',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-link',
      ...edit(who),
      from: inside(doc, layout),
      to: inside(doc, layout, true),
      href: 'https://example.com/b',
    }),
    // A party's link wraps only its own new words (Word has no tracked link around words)
    free: 'apply-refused/invalid-edit',
  },
  {
    name: 'a picture inside',
    kind: 'insert-object',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'insert-object',
      ...edit(who),
      place: inside(doc, layout),
      object: { type: 'picture', media: MEDIA, extent: { cx: 914_400, cy: 457_200 } },
    }),
  },
  {
    name: "the control's value",
    kind: 'control-value',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who) => ({
      kind: 'control-value',
      ...edit(who),
      control: addressOf(control(doc)),
      value: { kind: 'text', text: 'EUR' },
    }),
  },
  {
    name: 'a new control at a place inside',
    kind: 'insert-control',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who, layout) => ({
      kind: 'insert-control',
      ...edit(who),
      place: inside(doc, layout),
      control: { type: 'plainText' },
    }),
    // The engine puts no new control inside an inline one
    free: 'apply-refused/engine-refused',
  },
  {
    name: 'a new control around words inside',
    kind: 'wrap-control',
    layouts: TEXT_LAYOUTS,
    who: BOTH,
    make: (doc, who, layout) => ({
      kind: 'wrap-control',
      ...edit(who),
      from: inside(doc, layout),
      to: inside(doc, layout, true),
      control: { type: 'richText' },
    }),
    // The engine puts no new control inside an inline one
    free: 'apply-refused/engine-refused',
  },
  {
    name: 'the control deleted with its content',
    kind: 'delete-control',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who) => ({
      kind: 'delete-control',
      ...edit(who),
      control: addressOf(control(doc)),
    }),
    // A content lock guards the words while the control exists; Word deletes the control whole
    when: removal,
  },
  {
    name: 'the control removed, its content kept',
    kind: 'remove-control',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who) => ({
      kind: 'remove-control',
      ...edit(who),
      control: addressOf(control(doc)),
    }),
    when: removal,
  },
  {
    name: "the control's tag",
    kind: 'control-property',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who) => ({
      kind: 'control-property',
      ...edit(who),
      control: addressOf(control(doc)),
      tag: 'renamed',
    }),
    when: (lock) => lock !== 'unlocked',
  },
  {
    name: "the control's lock alone",
    kind: 'control-property',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who) => ({
      kind: 'control-property',
      ...edit(who),
      control: addressOf(control(doc)),
      lock: 'unlocked',
    }),
    // A control's own lock does not guard the lock itself: Word's Properties dialog changes it
    when: () => false,
  },
  {
    name: 'the paragraph that holds the control deleted',
    kind: 'delete-paragraph',
    layouts: ['inline'],
    who: ORGANISER,
    make: (doc, who) => ({ kind: 'delete-paragraph', ...edit(who), mark: markOf(doc, 0) }),
    when: (lock) => lock !== 'unlocked',
  },
  {
    name: 'a paragraph of the control deleted',
    kind: 'delete-paragraph',
    layouts: ['block'],
    who: ORGANISER,
    make: (doc, who) => ({ kind: 'delete-paragraph', ...edit(who), mark: markOf(doc, 1) }),
  },
  {
    name: 'the story replaced',
    kind: 'replace-story',
    layouts: TEXT_LAYOUTS,
    who: ORGANISER,
    make: (doc, who) => ({ kind: 'replace-story', ...edit(who), paragraphs: ['New.'] }),
    when: (lock) => lock !== 'unlocked',
  },
  {
    name: 'a direct join of the paragraph that holds the control',
    kind: 'join',
    layouts: ['inline'],
    who: BOTH,
    make: (doc, who) => ({ kind: 'join', ...edit(who), mark: markOf(doc, 0) }),
    // The editor alone's join rearranges both paragraphs (`joinParagraphs` is structural); a
    // tracked join only marks the paragraph mark (`setParagraphMarkRevision`)
    when: (lock, _, who) => content(lock) && who === 'organiser',
  },
  // ---- tables ----
  {
    name: 'a new row',
    kind: 'insert-row',
    layouts: ['table', 'tableInBlock'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'insert-row',
      ...edit(who),
      row: addressOf(firstOf(doc, 'tr')),
      where: 'below',
    }),
  },
  {
    name: 'the row that holds the control deleted',
    kind: 'delete-row',
    layouts: ['table'],
    who: BOTH,
    make: (doc, who) => ({ kind: 'delete-row', ...edit(who), row: addressOf(firstOf(doc, 'tr')) }),
    when: (lock) => lock !== 'unlocked',
  },
  {
    name: 'another row of the table deleted',
    kind: 'delete-row',
    layouts: ['table', 'tableInBlock'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'delete-row',
      ...edit(who),
      row: addressOf(opens(doc, 'tr')[1] ?? firstOf(doc, 'tr')),
    }),
  },
  {
    name: 'a new column',
    kind: 'insert-column',
    layouts: ['table', 'tableInBlock'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'insert-column',
      ...edit(who),
      cell: addressOf(opens(doc, 'tc')[1] ?? firstOf(doc, 'tc')),
      where: 'right',
    }),
  },
  {
    name: 'a column without the control deleted',
    kind: 'delete-column',
    layouts: ['table', 'tableInBlock'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'delete-column',
      ...edit(who),
      cell: addressOf(opens(doc, 'tc')[1] ?? firstOf(doc, 'tc')),
    }),
    // The editor alone's column deletion reaches every control of the table, and removes
    when: (lock, layout) => (layout === 'table' ? lock !== 'unlocked' : content(lock)),
  },
  {
    name: "a cell's fill",
    kind: 'table-property',
    layouts: ['table', 'tableInBlock'],
    who: BOTH,
    make: (doc, who) => ({
      kind: 'table-property',
      ...edit(who),
      on: [addressOf(opens(doc, 'tc')[3] ?? firstOf(doc, 'tc'))],
      op: { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFFF00' } },
    }),
    // A cell's properties change no control's content: only a control around the table refuses
    when: (lock, layout) => content(lock) && layout === 'tableInBlock',
  },
];

/** The verdict of each case, organiser and party, each lock, each layout; mismatches only. */
function mismatches(cases: readonly Case[]): string[] {
  const out: string[] = [];
  for (const c of cases)
    for (const layout of c.layouts)
      for (const lock of LOCKS)
        for (const who of c.who) {
          const { doc, send } = roomOn(fileOf(layout, lock));
          const before = bodyFingerprint(doc);
          const verdict = send(who, c.make(doc, who, layout));
          const refuses = (c.when ?? ((l: Lock) => content(l)))(lock, layout, who);
          const want = refuses ? 'apply-refused/locked' : (c.free ?? 'accepted');
          const row = `${c.name} | ${layout} ${lock} ${who}: ${verdict}`;
          if (verdict !== want && !(want === c.free && verdict === 'accepted'))
            out.push(`${row} (want ${want})`);
          else if (refuses && bodyFingerprint(doc) !== before) out.push(`${row} (changed)`);
        }
  return out;
}

describe('a room refuses every change a lock forbids, as the editor alone does', () => {
  test('text: typing, tabs, breaks, deletions, formats, Enter, moves, notes, fields, pastes, links, pictures', () => {
    expect(
      mismatches(CASES.filter((c) => ['inline', 'block'].includes(c.layouts[0] ?? '')))
    ).toEqual([]);
  });

  test('tables: rows, columns and cell properties', () => {
    expect(mismatches(CASES.filter((c) => c.layouts[0] === 'table'))).toEqual([]);
  });
});

/**
 * Every change kind is placed: either a case above reaches a locked control with it, or it
 * cannot reach a control's content, for the reason given. A new kind that nobody places fails
 * here, as the engine's `TREE_OP_REACH` fails a new op.
 */
const NO_REACH: Readonly<Record<string, string>> = {
  'building-block': "the engine's own op on the view refuses it (a value write)",
  'repeating-item': "the engine's own op on the view refuses it",
  'form-field': "the engine's own op on the view refuses it",
  'field-result': 'refused by `refuseLocked` on the field, as typing is',
  'toc-pages': 'a contents result is refused whole (`not-covered`)',
  'toc-result': 'a contents result is refused whole (`not-covered`)',
  'insert-toc': 'a new contents block goes before a paragraph, as a new table does',
  'note-properties': "a document setting: no control's content",
  'section-property': "a document setting: no control's content",
  'header-footer': "a part's wiring: no control's content",
  watermark: "a document setting: no control's content",
  setting: "a document setting: no control's content",
  attribution: "a suggestion's author: no words change",
  decide:
    'a card a lock keeps is left, as Word leaves it (`client/editor-decide-all-locked.test.ts`)',
  retract: 'a take-back of your own suggestion: it gives words back, adds none',
  revert: "the organiser's undo of a direct edit",
  comment: 'a comment adds no words (the host makes discussions on any text)',
  'comment-reply': "a reply sits beside its thread's range and adds no words",
  'comment-resolve': 'a resolve changes no words',
};

/** Kinds the tests after the case table check: objects, notes and links inside a control. */
const PLACED_ABOVE = [
  'object-property',
  'object-resource',
  'convert-note',
  'field-code',
  'form-text',
  'equation',
  'link-target',
  'remove-link',
];

test('every change kind is placed: a lock case, or a reason it cannot reach the content', () => {
  const placed = new Set([...CASES.map((c) => c.kind), ...PLACED_ABOVE, ...Object.keys(NO_REACH)]);
  expect(ORDERED_CHANGE_KINDS.filter((kind) => !placed.has(kind))).toEqual([]);
});
