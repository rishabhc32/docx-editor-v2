// The change validator: every kind in a minimal valid form, every cap, unknown fields, wrong
// types, forged server fields, prototype keys, link targets, ids, and random input that must be
// refused by code, never thrown.

import { describe, expect, test } from 'bun:test';
import { isValidParaId } from '../../../store/package/para-id.ts';
import { orderedChangeId, orderedIdMinter, parseOrderedChangeId } from '../change/ids.ts';
import { CONTROL_KINDS } from '../change/kinds/controls.ts';
import { LINK_FIELD_KINDS } from '../change/kinds/links-fields.ts';
import { NOTE_KINDS } from '../change/kinds/notes.ts';
import { PASTE_BLOCK_KINDS } from '../change/kinds/paste-blocks.ts';
import { PICTURE_KINDS } from '../change/kinds/pictures.ts';
import { ORDERED_CHANGE_KINDS, orderedKindSpecOf } from '../change/kinds/registry.ts';
import { SECTION_KINDS } from '../change/kinds/sections.ts';
import { TABLE_KINDS } from '../change/kinds/tables.ts';
import { ORDERED_CHANGE_LIMITS as LIMITS, ORDERED_CORE_CHANGE_KINDS } from '../change/types.ts';
import {
  validateOrderedBatch,
  validateOrderedChange,
  validateOrderedStampedChange,
} from '../change/validate.ts';
import { decodeOrderedClientMessage } from '../protocol/codec.ts';
import { seededRandom } from './random.ts';
import { testPerson } from './test-persons.ts';

const NS = 'ana-1';
const SPACE = 'b:/word/document.xml#0.0.0';
const A = { space: SPACE, offset: 0 };
const B = { space: SPACE, offset: 5 };
const MARK = { space: 'm:/word/document.xml#0.0.0', offset: 0 };
const TOKEN = { space: 't:/word/document.xml#0.0.1', offset: 0 };
const EDIT = { part: '/word/document.xml', mode: 'suggest' };

/** One minimal valid change of every kind. */
const MINIMAL: Record<string, Record<string, unknown>> = {
  insert: { kind: 'insert', ...EDIT, place: A, text: 'x' },
  'insert-tab': { kind: 'insert-tab', ...EDIT, place: A },
  'insert-break': { kind: 'insert-break', ...EDIT, place: A, break: 'line', clear: 'all' },
  delete: { kind: 'delete', ...EDIT, from: A, to: B },
  format: { kind: 'format', ...EDIT, from: A, to: B, properties: [{ localName: 'b' }] },
  'mark-format': { kind: 'mark-format', ...EDIT, marks: [MARK], properties: [{ localName: 'b' }] },
  split: { kind: 'split', ...EDIT, place: { mark: MARK } },
  join: { kind: 'join', ...EDIT, mark: MARK },
  'paragraph-property': {
    kind: 'paragraph-property',
    ...EDIT,
    marks: [MARK],
    properties: [{ localName: 'jc', attributes: { val: 'center' } }],
  },
  'table-property': {
    kind: 'table-property',
    ...EDIT,
    on: [TOKEN],
    op: {
      op: 'setTableCellBorders',
      scope: 'outside',
      spec: { style: 'single', size: 4, color: { kind: 'hex', value: '1F3864' } },
    },
  },
  'insert-row': { kind: 'insert-row', ...EDIT, row: TOKEN, where: 'below' },
  'delete-row': { kind: 'delete-row', ...EDIT, row: TOKEN },
  'insert-column': { kind: 'insert-column', ...EDIT, cell: TOKEN, where: 'left' },
  'delete-column': { kind: 'delete-column', ...EDIT, cell: TOKEN },
  'insert-table': {
    kind: 'insert-table',
    ...EDIT,
    place: A,
    rows: 2,
    columns: 3,
    columnWidth: 2880,
  },
  'control-value': {
    kind: 'control-value',
    ...EDIT,
    control: TOKEN,
    value: { kind: 'checkbox', checked: true },
  },
  'link-target': { kind: 'link-target', ...EDIT, link: TOKEN, href: 'https://example.com/a?b=c#d' },
  'insert-control': {
    kind: 'insert-control',
    ...EDIT,
    place: A,
    control: {
      type: 'custom',
      tag: 'cite:1',
      text: '[1]',
      lock: 'sdtContentLocked',
      binding: {
        prefixMappings: "xmlns:c='urn:c'",
        xpath: '/c:root/c:item[1]',
        storeItemId: '{00000000-0000-0000-0000-000000000001}',
      },
    },
  },
  'wrap-control': { kind: 'wrap-control', ...EDIT, from: A, to: B, control: { type: 'date' } },
  'remove-control': { kind: 'remove-control', ...EDIT, control: TOKEN },
  'repeating-item': { kind: 'repeating-item', ...EDIT, control: TOKEN, action: 'remove', index: 1 },
  'building-block': { kind: 'building-block', ...EDIT, control: TOKEN, name: 'Party block' },
  'form-text': {
    kind: 'form-text',
    ...EDIT,
    field: A,
    op: { op: 'insertText', offset: 1, text: 'x' },
  },
  'form-field': {
    kind: 'form-field',
    ...EDIT,
    field: A,
    op: {
      op: 'setTextFormFieldDefault',
      text: 'Name',
      options: { type: 'regular', maxLength: 20, format: 'Uppercase', enabled: true },
      locale: 'en-US',
    },
  },
  'delete-control': { kind: 'delete-control', ...EDIT, control: TOKEN },
  'control-property': {
    kind: 'control-property',
    ...EDIT,
    control: TOKEN,
    tag: null,
    lock: 'sdtLocked',
  },
  'insert-link': {
    kind: 'insert-link',
    ...EDIT,
    from: A,
    to: B,
    anchor: '_Toc1',
    style: 'Hyperlink',
  },
  'remove-link': { kind: 'remove-link', ...EDIT, link: TOKEN, from: A, to: B, style: 'Hyperlink' },
  'insert-field': { kind: 'insert-field', ...EDIT, place: A, field: 'PAGE_X_OF_Y' },
  'field-code': { kind: 'field-code', ...EDIT, field: A, code: 'NUMPAGES' },
  'field-result': {
    kind: 'field-result',
    ...EDIT,
    fields: [
      { field: A, text: '4\t5' },
      { field: B, text: '', last: true },
    ],
  },
  'toc-pages': { kind: 'toc-pages', ...EDIT, pages: [{ from: A, to: B, text: '12' }] },
  bookmark: { kind: 'bookmark', ...EDIT, name: 'Terms_1', from: A, to: A, remove: [A] },
  'insert-toc': {
    kind: 'insert-toc',
    ...EDIT,
    before: MARK,
    alias: 'Table of Contents',
    instruction: 'TOC \\o "1-3" \\h \\z \\u',
    rows: [{ level: 0, text: 'Terms', heading: MARK, bookmark: '_Toc1', page: '2' }],
    bookmarks: [{ heading: MARK, name: '_Toc1' }],
  },
  'toc-result': {
    kind: 'toc-result',
    ...EDIT,
    separate: A,
    end: B,
    instruction: 'TOC \\o "1-3"',
    rows: [],
    bookmarks: [],
  },
  move: { kind: 'move', ...EDIT, from: A, to: B, place: { mark: MARK } },
  'insert-note': { kind: 'insert-note', ...EDIT, place: A, note: 'footnote', text: '' },
  equation: { kind: 'equation', ...EDIT, object: A, equation: 0, linear: 'x^2' },
  'convert-note': { kind: 'convert-note', ...EDIT, mode: 'edit', notes: [TOKEN] },
  'note-properties': {
    kind: 'note-properties',
    ...EDIT,
    mode: 'edit',
    section: { mark: MARK },
    footnote: {
      position: 'beneathText',
      numFmt: 'lowerRoman',
      numStart: 2,
      numRestart: 'eachSect',
    },
    endnote: { position: 'docEnd' },
  },
  'insert-object': {
    kind: 'insert-object',
    ...EDIT,
    place: A,
    object: {
      type: 'picture',
      media: { sha256: 'a'.repeat(64), contentType: 'image/png', size: 70 },
      extent: { cx: 914_400, cy: 457_200 },
    },
  },
  'object-property': {
    kind: 'object-property',
    ...EDIT,
    object: A,
    ops: [{ op: 'resize', cx: 457_200, cy: 457_200 }],
  },
  'object-resource': {
    kind: 'object-resource',
    ...EDIT,
    object: A,
    media: { sha256: 'b'.repeat(64), contentType: 'image/gif', size: 43 },
  },
  decide: {
    kind: 'decide',
    cards: [
      {
        suggestions: [
          { suggestion: 'S3', version: 2 },
          { suggestion: 'S:ben-2:4.1', version: 0 },
        ],
      },
    ],
    status: 'rejected',
  },
  retract: { kind: 'retract', suggestions: ['S:ana-1:1.1'] },
  revert: { kind: 'revert', suggestions: ['S:olga-1:1.1'], status: 'rejected' },
  comment: { kind: 'comment', part: '/word/document.xml', from: A, to: B, text: 'Why?\nSee 2.' },
  'comment-reply': { kind: 'comment-reply', comment: '1000000001', text: 'Agreed.' },
  'comment-resolve': { kind: 'comment-resolve', comment: '7' },
  'section-property': {
    kind: 'section-property',
    ...EDIT,
    sections: [],
    page: { orientation: 'landscape' },
    pageAt: MARK,
  },
  'header-footer': {
    kind: 'header-footer',
    ...EDIT,
    action: 'create',
    section: MARK,
    story: 'footer',
    variant: 'default',
  },
  'section-break': { kind: 'section-break', ...EDIT, mark: MARK, type: 'continuous' },
  watermark: { kind: 'watermark', ...EDIT, watermark: null },
  setting: { kind: 'setting', ...EDIT, setting: { op: 'setDocumentProtection', enforce: true } },
  'delete-paragraph': { kind: 'delete-paragraph', ...EDIT, mark: MARK },
  'insert-fragment': {
    kind: 'insert-fragment',
    ...EDIT,
    place: A,
    blocks: [
      {
        e: 'w:p',
        c: [
          {
            e: 'w:pPr',
            c: [{ e: 'w:numPr', c: [{ e: 'w:numId', a: { 'w:val': '#0' } }] }],
          },
          { e: 'w:hyperlink', a: { 'r:id': '#0' }, c: [{ e: 'w:r', c: [{ e: 'w:t', c: ['x'] }] }] },
        ],
      },
    ],
    whole: true,
    lists: [{ abstract: { e: 'w:abstractNum', c: [{ e: 'w:lvl', a: { 'w:ilvl': '0' } }] } }],
    styles: [{ e: 'w:style', a: { 'w:type': 'paragraph', 'w:styleId': 'Clause' } }],
    links: ['https://example.com/'],
  },
  'new-list': {
    kind: 'new-list',
    ...EDIT,
    list: { abstract: { e: 'w:abstractNum', c: [{ e: 'w:lvl', a: { 'w:ilvl': '0' } }] } },
    paragraphs: [{ mark: MARK, level: 0 }],
  },
  'list-level': {
    kind: 'list-level',
    ...EDIT,
    numId: '1',
    level: 1,
    format: 'ordered',
    marks: [MARK],
  },
  'replace-story': { kind: 'replace-story', ...EDIT, paragraphs: ['One.', ''] },
  attribution: {
    kind: 'attribution',
    ...EDIT,
    suggestions: ['S3'],
    author: 'Ana Legal',
    date: '2026-10-29T10:00:00Z',
  },
};

const change = (kind: string, extra: Record<string, unknown> = {}, id = `${NS}:1`) => ({
  v: 1,
  id,
  ...MINIMAL[kind],
  ...extra,
});

/** Cards of two suggestions each (the last of one when `count` is odd), `count` in all. */
const seenCards = (count: number) =>
  Array.from({ length: Math.ceil(count / 2) }, (_, k) => ({
    suggestions: Array.from({ length: Math.min(2, count - 2 * k) }, (_, j) => ({
      suggestion: `S${2 * k + j + 1}`,
      version: 0,
    })),
  }));

/** The refusal code of a change, or 'ok'. */
function codeOf(value: unknown, options = { namespace: NS }): string {
  const result = validateOrderedChange(value, options);
  return result.ok ? 'ok' : result.error.code;
}

describe('the change validator', () => {
  test('accepts every kind in a minimal form, and returns the same change', () => {
    expect(Object.keys(MINIMAL).sort()).toEqual([...ORDERED_CHANGE_KINDS].sort());
    for (const kind of ORDERED_CHANGE_KINDS) {
      const input = change(kind);
      const result = validateOrderedChange(input, {
        namespace: NS,
        parts: new Set(['/word/document.xml']),
      });
      expect({ kind, ok: result.ok }).toEqual({ kind, ok: true });
      if (result.ok) expect(result.value).toEqual(input as never);
    }
  });

  test('takes a property off in a format, a mark format and a paragraph press', () => {
    const removal = { localName: 'sz', remove: true };
    for (const [kind, properties] of [
      ['format', [{ localName: 'b' }, removal]],
      ['mark-format', [removal]],
      ['paragraph-property', [{ localName: 'jc', remove: true }]],
    ] as const) {
      const input = change(kind, { properties });
      const valid = validateOrderedChange(input, { namespace: NS });
      expect(valid.ok && valid.value).toEqual(input as never);
    }
  });

  test('returns a new object with known fields only, in a fixed order', () => {
    const input = change('insert', { into: 'S:ana-1:9.1' });
    const result = validateOrderedChange(input, { namespace: NS });
    if (!result.ok) throw new Error(result.error.code);
    expect(result.value).not.toBe(input);
    expect(Object.keys(result.value)).toEqual([
      'v',
      'id',
      'kind',
      'part',
      'mode',
      'into',
      'place',
      'text',
    ]);
    expect(validateOrderedChange(change('insert', { offset: 0 }), { namespace: NS }).ok).toBe(
      false
    );
  });

  test('enforces every cap', () => {
    const cases: [string, unknown, unknown][] = [
      [
        'text',
        change('insert', { text: 'a'.repeat(LIMITS.text) }),
        change('insert', { text: 'a'.repeat(LIMITS.text + 1) }),
      ],
      [
        'space',
        change('insert', { place: { space: `b:${'x'.repeat(LIMITS.space - 2)}`, offset: 0 } }),
        change('insert', { place: { space: `b:${'x'.repeat(LIMITS.space - 1)}`, offset: 0 } }),
      ],
      [
        'targets',
        change('join', {}),
        change('paragraph-property', {
          marks: Array.from({ length: LIMITS.targets + 1 }, () => MARK),
        }),
      ],
      [
        'properties',
        change('format', {}),
        change('format', {
          properties: Array.from({ length: LIMITS.properties + 1 }, () => ({ localName: 'b' })),
        }),
      ],
      [
        'attributes',
        change('format', {
          properties: [
            {
              localName: 'rFonts',
              attributes: Object.fromEntries(
                Array.from({ length: LIMITS.attributes }, (_, k) => [`a${k}`, 'x'])
              ),
            },
          ],
        }),
        change('format', {
          properties: [
            {
              localName: 'rFonts',
              attributes: Object.fromEntries(
                Array.from({ length: LIMITS.attributes + 1 }, (_, k) => [`a${k}`, 'x'])
              ),
            },
          ],
        }),
      ],
      [
        'attribute value',
        change('format', {
          properties: [
            { localName: 'color', attributes: { val: 'f'.repeat(LIMITS.attributeValue) } },
          ],
        }),
        change('format', {
          properties: [
            { localName: 'color', attributes: { val: 'f'.repeat(LIMITS.attributeValue + 1) } },
          ],
        }),
      ],
      [
        'children',
        change('paragraph-property', {
          properties: [
            {
              localName: 'tabs',
              children: Array.from({ length: LIMITS.children }, () => ({ localName: 'tab' })),
            },
          ],
        }),
        change('paragraph-property', {
          properties: [
            {
              localName: 'tabs',
              children: Array.from({ length: LIMITS.children + 1 }, () => ({ localName: 'tab' })),
            },
          ],
        }),
      ],
      [
        'href',
        change('link-target', { href: `https://e.com/${'a'.repeat(LIMITS.href - 14)}` }),
        change('link-target', { href: `https://e.com/${'a'.repeat(LIMITS.href - 13)}` }),
      ],
      [
        'suggestion',
        change('retract', { suggestions: ['S'.repeat(LIMITS.suggestion)] }),
        change('retract', { suggestions: ['S'.repeat(LIMITS.suggestion + 1)] }),
      ],
      [
        // Every card of an Accept All in one decision, counted over its cards
        'decisions',
        change('decide', { cards: seenCards(LIMITS.decisions) }),
        change('decide', { cards: seenCards(LIMITS.decisions + 1) }),
      ],
      [
        'rows',
        change('insert-table', { rows: LIMITS.tableRows }),
        change('insert-table', { rows: LIMITS.tableRows + 1 }),
      ],
      [
        'columns',
        change('insert-table', { columns: LIMITS.tableColumns }),
        change('insert-table', { columns: LIMITS.tableColumns + 1 }),
      ],
      [
        'column width',
        change('insert-table', { columnWidth: LIMITS.columnWidth }),
        change('insert-table', { columnWidth: LIMITS.columnWidth + 1 }),
      ],
      [
        'list item',
        change('control-value', {
          value: { kind: 'listItem', value: 'v'.repeat(LIMITS.listItem) },
        }),
        change('control-value', {
          value: { kind: 'listItem', value: 'v'.repeat(LIMITS.listItem + 1) },
        }),
      ],
    ];
    for (const [name, atCap, overCap] of cases) {
      expect({ name, code: codeOf(atCap) }).toEqual({ name, code: 'ok' });
      expect({ name, refused: codeOf(overCap) !== 'ok' }).toEqual({ name, refused: true });
    }
    const batch = (n: number) =>
      Array.from({ length: n }, (_, k) => change('insert', {}, `${NS}:${k + 1}`));
    expect(validateOrderedBatch(batch(LIMITS.batch), { namespace: NS }).ok).toBe(true);
    const over = validateOrderedBatch(batch(LIMITS.batch + 1), { namespace: NS });
    expect(over.ok ? 'ok' : over.error.code).toBe('too-many');
    const empty = validateOrderedBatch([], { namespace: NS });
    expect(empty.ok ? 'ok' : empty.error.code).toBe('empty');
  });

  test('refuses unknown fields, wrong types and malformed values by code', () => {
    const cases: [unknown, string][] = [
      [{ ...change('insert'), extra: 1 }, 'unknown-field'],
      [change('insert', { place: { ...A, side: 'left' } }), 'unknown-field'],
      [change('format', { properties: [{ localName: 'b', value: 1 }] }), 'unknown-field'],
      [change('format', { properties: [{ localName: 'b', children: [] }] }), 'unknown-field'],
      [
        change('format', {
          properties: [{ localName: 'b', remove: true, attributes: { val: '0' } }],
        }),
        'unknown-field',
      ],
      [change('format', { properties: [{ localName: 'b', remove: false }] }), 'invalid-value'],
      [change('list-level', { numId: 'x1' }), 'invalid-value'],
      [change('list-level', { level: 9 }), 'invalid-number'],
      [change('list-level', { format: 'roman' }), 'invalid-value'],
      [
        change('format', { properties: [{ localName: 'rStyle', remove: true }] }),
        'unknown-property',
      ],
      [change('decide', { mode: 'suggest' }), 'unknown-field'],
      [change('insert', { id: 7 }), 'invalid-type'],
      [change('insert', { place: 'here' }), 'invalid-type'],
      [change('insert', { text: 5 }), 'invalid-type'],
      [change('insert', { place: { space: SPACE, offset: '1' } }), 'invalid-type'],
      [change('insert', { place: { space: SPACE, offset: -1 } }), 'invalid-number'],
      [change('insert', { place: { space: SPACE, offset: 1.5 } }), 'invalid-number'],
      [change('insert', { place: { space: SPACE, offset: 2 ** 53 } }), 'invalid-number'],
      [change('insert', { place: { space: 'nospace', offset: 0 } }), 'invalid-space'],
      [change('insert', { place: { space: 'b:a b', offset: 0 } }), 'invalid-space'],
      [change('insert', { mode: 'maybe' }), 'invalid-value'],
      [change('insert', { kind: 7 }), 'invalid-type'],
      [change('insert', { kind: 'paste' }), 'unknown-kind'],
      [change('insert', { v: '1' }), 'invalid-version'],
      [change('insert', { v: 2 }), 'invalid-version'],
      [
        { v: 1, id: `${NS}:1`, kind: 'insert', mode: 'suggest', place: A, text: 'x' },
        'missing-field',
      ],
      [change('insert', { text: '' }), 'empty'],
      [change('insert', { text: 'a\tb' }), 'invalid-text'],
      [change('insert', { text: 'a\nb' }), 'invalid-text'],
      [change('insert', { text: '\u0000' }), 'invalid-text'],
      [change('insert', { text: '\u0085' }), 'invalid-text'],
      [change('insert', { text: '\ud800x' }), 'invalid-text'],
      [change('insert', { text: '\uffff' }), 'invalid-text'],
      [change('insert', { part: 'word/document.xml' }), 'invalid-part'],
      [change('format', { properties: [{ localName: 'rStyle' }] }), 'unknown-property'],
      [
        change('format', { properties: [{ localName: 'b' }, { localName: 'b' }] }),
        'repeated-property',
      ],
      [
        change('format', { properties: [{ localName: 'b', attributes: { 'w:val': '1' } }] }),
        'invalid-attribute',
      ],
      [
        change('format', { properties: [{ localName: 'b', attributes: { val: 1 } }] }),
        'invalid-type',
      ],
      [
        change('format', { properties: [{ localName: 'b', attributes: { val: '<x/>\u0001' } }] }),
        'invalid-text',
      ],
      [change('table-property', { op: { op: 'mergeCells' } }), 'invalid-value'],
      [change('table-property', { op: { op: 'setTableProperties' } }), 'empty'],
      [
        change('table-property', { op: { op: 'setTableRowHeight', heightTwips: 0 } }),
        'invalid-number',
      ],
      [
        change('table-property', {
          op: { op: 'setTableCellFill', color: { kind: 'hex', value: 'red' } },
        }),
        'invalid-value',
      ],
      [
        change('table-property', {
          op: { op: 'setTableCellBorders', scope: 'none', target: 'all', spec: {} },
        }),
        'unknown-field',
      ],
      [
        change('table-property', {
          op: { op: 'setTableProperties', justification: 'both' },
        }),
        'invalid-value',
      ],
      [change('insert-break', { break: 'section' }), 'invalid-value'],
      [change('insert-break', { break: 'page', clear: 'all' }), 'unknown-field'],
      [change('insert-break', { clear: 'none' }), 'invalid-value'],
      [change('insert-tab', { text: '\t' }), 'unknown-field'],
      [
        change('paragraph-property', {
          properties: [{ localName: 'numPr', children: [{ localName: 'tab' }] }],
        }),
        'unknown-property',
      ],
      [change('link-target', { anchor: 'top' }), 'missing-field'],
      [{ v: 1, id: `${NS}:1`, kind: 'link-target', ...EDIT, link: TOKEN }, 'missing-field'],
      [change('link-target', { href: undefined, anchor: 'a b' }), 'invalid-value'],
      [change('control-value', { value: { kind: 'date', iso: '28/10/2026' } }), 'invalid-date'],
      [
        change('control-value', { value: { kind: 'checkbox', checked: true, text: 'x' } }),
        'unknown-field',
      ],
      [change('decide', { status: 'maybe' }), 'invalid-value'],
      [
        change('decide', {
          cards: [
            {
              suggestions: [
                { suggestion: 'S1', version: 0 },
                { suggestion: 'S1', version: 1 },
              ],
            },
          ],
        }),
        'invalid-suggestion',
      ],
      // One suggestion on two cards
      [
        change('decide', {
          cards: [
            { suggestions: [{ suggestion: 'S1', version: 0 }] },
            { suggestions: [{ suggestion: 'S1', version: 0 }] },
          ],
        }),
        'invalid-suggestion',
      ],
      [change('decide', { cards: [] }), 'empty'],
      [change('decide', { cards: [{ suggestions: [] }] }), 'empty'],
      [change('decide', { cards: [{}] }), 'missing-field'],
      [change('decide', { suggestions: [{ suggestion: 'S1', version: 0 }] }), 'unknown-field'],
      [change('decide', { cards: [{ suggestions: ['S1'] }] }), 'invalid-type'],
      [change('decide', { cards: [{ suggestions: [{ suggestion: 'S1' }] }] }), 'missing-field'],
      [
        change('decide', { cards: [{ suggestions: [{ suggestion: 'S1', version: -1 }] }] }),
        'invalid-number',
      ],
      [
        change('decide', {
          cards: [{ suggestions: [{ suggestion: 'S1', version: 0, seen: true }] }],
        }),
        'unknown-field',
      ],
      [change('retract', { suggestions: ['S 1'] }), 'invalid-suggestion'],
      [change('retract', { suggestions: [] }), 'empty'],
      [change('retract', { suggestions: [], undo: [] }), 'empty'],
      [change('retract', { undo: ['ana-1'] }), 'invalid-id'],
      [change('retract', { undo: ['ben-1:2'] }), 'foreign-id'],
      [change('retract', { undo: [`${NS}:2`, `${NS}:2`] }), 'repeated-id'],
      [change('revert', { status: 'pending' }), 'invalid-value'],
      [change('revert', { suggestions: [] }), 'empty'],
      [change('revert', { part: '/word/document.xml' }), 'unknown-field'],
      [change('comment', { text: '' }), 'empty'],
      [change('comment', { text: 'a'.repeat(10_001) }), 'too-long'],
      [change('comment', { text: 'Tab\tin' }), 'invalid-text'],
      [change('comment', { text: 'Line\r\nend' }), 'invalid-text'],
      [change('comment', { mode: 'suggest' }), 'unknown-field'],
      [change('comment-reply', { comment: '' }), 'empty'],
      [change('comment-reply', { comment: 7 }), 'invalid-type'],
      [change('comment-reply', { text: 'a'.repeat(10_001) }), 'too-long'],
      [change('comment-resolve', { note: '' }), 'empty'],
      [change('comment-resolve', { comment: 'x'.repeat(65) }), 'too-long'],
      [change('section-property', { sections: null }), 'invalid-type'],
      [null, 'invalid-type'],
      [[change('insert')], 'invalid-type'],
      ['insert', 'invalid-type'],
    ];
    for (const [value, code] of cases) {
      const parsed = JSON.parse(JSON.stringify(value) ?? 'null') as unknown;
      expect({ value, code: codeOf(parsed) }).toEqual({ value, code });
    }
  });

  test('refuses n, by and at from a client, and checks them on an ordered change', () => {
    const ana = testPerson('Ana');
    for (const forged of [{ n: 1 }, { by: testPerson('Ben') }, { at: '2026-10-28T00:00:00Z' }])
      expect(codeOf(change('insert', forged))).toBe('server-field');
    const stamped = { ...change('insert'), n: 3, by: ana, at: '2026-10-28T09:30:00Z' };
    expect(validateOrderedStampedChange(stamped).ok).toBe(true);
    const unknown = validateOrderedStampedChange({ ...stamped, kind: 'paste' });
    expect(unknown.ok ? 'ok' : unknown.error.code).toBe('unknown-kind');
    const codes = [
      { ...stamped, n: 0 },
      { ...stamped, by: 'Ana' },
      { ...stamped, by: { ...ana, name: '' } },
      { ...stamped, by: { ...ana, name: ' Ana' } },
      { ...stamped, by: { ...ana, role: 'x' } },
      { ...stamped, at: '2026-10-28 09:30' },
      { ...change('insert'), by: ana, at: '2026-10-28T09:30:00Z' },
    ].map((value) => {
      const result = validateOrderedStampedChange(value);
      return result.ok ? 'ok' : result.error.code;
    });
    expect(codes).toEqual([
      'invalid-number',
      'invalid-type',
      'empty',
      'invalid-value',
      'unknown-field',
      'invalid-date',
      'missing-field',
    ]);
  });

  test('refuses prototype keys anywhere and pollutes nothing', () => {
    const frames = [
      `{"__proto__":{"polluted":1},"v":1,"id":"${NS}:1","kind":"insert"}`,
      `{"v":1,"id":"${NS}:1","kind":"insert","part":"/word/document.xml","mode":"suggest","place":{"space":"${SPACE}","offset":0,"constructor":{}},"text":"x"}`,
      `{"v":1,"id":"${NS}:1","kind":"format","part":"/word/document.xml","mode":"suggest","from":{"space":"${SPACE}","offset":0},"to":{"space":"${SPACE}","offset":1},"properties":[{"localName":"b","attributes":{"__proto__":"x"}}]}`,
      `{"v":1,"id":"${NS}:1","kind":"format","part":"/word/document.xml","mode":"suggest","from":{"space":"${SPACE}","offset":0},"to":{"space":"${SPACE}","offset":1},"properties":[{"localName":"b","prototype":{}}]}`,
      `{"v":1,"id":"${NS}:1","kind":"control-value","part":"/word/document.xml","mode":"suggest","control":{"space":"${SPACE}","offset":0},"value":{"kind":"text","text":"x","__proto__":{}}}`,
    ];
    for (const frame of frames) expect(codeOf(JSON.parse(frame))).toBe('forbidden-key');
    expect((Object.prototype as Record<string, unknown>).polluted).toBeUndefined();
    expect(codeOf(Object.create({ v: 1, id: `${NS}:1`, kind: 'insert' }))).toBe('invalid-type');
  });

  test('labels an unknown key the same at the top of a change and inside a field', () => {
    const key = `x${'\u0001'.repeat(3)}${'k'.repeat(200)}`;
    const label = `x???${'k'.repeat(60)}...`;
    const top = validateOrderedChange({ ...change('insert'), [key]: 1 });
    const nested = validateOrderedChange(change('insert', { place: { ...A, [key]: 1 } }));
    expect(top.ok ? 'ok' : top.error.path).toBe(label);
    expect(nested.ok ? 'ok' : nested.error.path).toBe(`place.${label}`);
  });

  test('passes link targets through sanitizeHref, unchanged, or refuses them', () => {
    for (const href of [
      'https://example.com/a?b=c#d',
      'http://x.y',
      'mailto:a@b.c',
      'tel:+15551234',
      '/relative/path',
      '#section',
    ])
      expect({ href, code: codeOf(change('link-target', { href })) }).toEqual({ href, code: 'ok' });
    for (const href of [
      'javascript:alert(1)',
      'JaVaScRiPt:alert(1)',
      'java\nscript:alert(1)',
      ' https://example.com',
      'data:text/html,<script>',
      'vbscript:msgbox',
      'file:///etc/passwd',
      'https://exa mple.com',
      'https://例え.jp',
    ])
      expect({ href, code: codeOf(change('link-target', { href })) }).toEqual({
        href,
        code: 'invalid-href',
      });
  });

  test("a take-back may only show items again, named by the sender's own earlier changes", () => {
    const undoing = change('retract', { suggestions: [], undo: [`${NS}:2`, `${NS}:3`] });
    const read = validateOrderedChange(undoing, { namespace: NS });
    expect(read.ok ? (read.value as unknown) : read.error).toEqual(undoing);
  });

  test('checks ids against the pattern, the sender namespace, and repeats in a batch', () => {
    for (const id of [
      'ana',
      'ana:0',
      'ana:01',
      'ana:1234567890123456',
      'a b:1',
      'ana:-1',
      ':1',
      `${'n'.repeat(65)}:1`,
    ])
      expect({ id, code: codeOf(change('insert', {}, id), { namespace: 'ana' }) }).toEqual({
        id,
        code: 'invalid-id',
      });
    expect(codeOf(change('insert', {}, 'ben-1:1'))).toBe('foreign-id');
    expect(validateOrderedChange(change('insert', {}, 'ben-1:1')).ok).toBe(true);
    const twice = validateOrderedBatch([change('insert'), change('split')], { namespace: NS });
    expect(twice.ok ? 'ok' : `${twice.error.code} ${twice.error.path}`).toBe('repeated-id [1].id');
    const parts = { namespace: NS, parts: new Set(['/word/document.xml']) };
    expect(codeOf(change('insert', { part: '/word/header9.xml' }), parts)).toBe('unknown-part');
  });

  test('random input is refused by code, never thrown', () => {
    const rand = seededRandom(20261028);
    const words = [
      'v',
      'id',
      'kind',
      'part',
      'mode',
      'into',
      'place',
      'from',
      'to',
      'mark',
      'space',
      'offset',
      'text',
      'properties',
      'localName',
      'attributes',
      'children',
      'value',
      'href',
      'anchor',
      'suggestions',
      'status',
      'n',
      'by',
      'at',
      '__proto__',
      'constructor',
      'prototype',
      'x',
    ];
    const leaf = (): unknown => {
      const pick = Math.floor(rand() * 9);
      if (pick === 0) return null;
      if (pick === 1) return rand() < 0.5;
      if (pick === 2) return Math.floor(rand() * 2000) - 1000;
      if (pick === 3) return rand() * 1e20;
      if (pick === 4) return ORDERED_CHANGE_KINDS[Math.floor(rand() * ORDERED_CHANGE_KINDS.length)];
      if (pick === 5) return `${NS}:${Math.floor(rand() * 50)}`;
      if (pick === 6) return SPACE;
      if (pick === 7) return String.fromCharCode(Math.floor(rand() * 0xffff));
      return words[Math.floor(rand() * words.length)];
    };
    const value = (depth: number): unknown => {
      const pick = rand();
      if (depth > 3 || pick < 0.4) return leaf();
      if (pick < 0.6) return Array.from({ length: Math.floor(rand() * 4) }, () => value(depth + 1));
      const out: Record<string, unknown> = {};
      for (let k = Math.floor(rand() * 6); k > 0; k -= 1)
        out[words[Math.floor(rand() * words.length)] ?? 'x'] = value(depth + 1);
      return out;
    };
    const mutate = (input: Record<string, unknown>): unknown => {
      const copy: Record<string, unknown> = JSON.parse(JSON.stringify(input)) as Record<
        string,
        unknown
      >;
      const keys = Object.keys(copy);
      const key = keys[Math.floor(rand() * keys.length)] ?? 'v';
      const pick = rand();
      if (pick < 0.3) delete copy[key];
      else if (pick < 0.7) copy[key] = value(2);
      else copy[words[Math.floor(rand() * words.length)] ?? 'x'] = value(2);
      return JSON.parse(JSON.stringify(copy)) as unknown;
    };
    let refused = 0;
    let accepted = 0;
    for (let k = 0; k < 4000; k += 1) {
      const kind = ORDERED_CHANGE_KINDS[k % ORDERED_CHANGE_KINDS.length] ?? 'insert';
      const input =
        k % 2 === 0 ? mutate(change(kind)) : JSON.parse(JSON.stringify(value(0)) ?? 'null');
      for (const result of [
        validateOrderedChange(input, { namespace: NS }),
        validateOrderedStampedChange(input),
        validateOrderedBatch([input], { namespace: NS }),
      ]) {
        if (result.ok) accepted += 1;
        else {
          refused += 1;
          expect(typeof result.error.code).toBe('string');
          expect(result.error.path.length).toBeLessThan(200);
        }
      }
    }
    expect(refused).toBeGreaterThan(8000);
    expect(accepted).toBeGreaterThan(0);
  });
});

/** Every copy of a value with one node made wrong: another JSON type, or an unknown field. */
function* corrupted(value: unknown): Generator<unknown> {
  yield typeof value === 'string' ? 7 : 'x';
  if (Array.isArray(value)) {
    for (const [k, item] of value.entries())
      for (const wrong of corrupted(item)) yield value.map((other, j) => (j === k ? wrong : other));
  } else if (typeof value === 'object' && value !== null) {
    yield { ...value, zz: 1 };
    for (const [key, item] of Object.entries(value))
      for (const wrong of corrupted(item)) yield { ...value, [key]: wrong };
  }
}

/** Each family's kinds: the format's first kinds (tables among them), then each group's. */
const FAMILIES: Record<string, readonly string[]> = {
  first: ORDERED_CORE_CHANGE_KINDS,
  pictures: Object.keys(PICTURE_KINDS),
  'links and fields': Object.keys(LINK_FIELD_KINDS),
  controls: Object.keys(CONTROL_KINDS),
  tables: Object.keys(TABLE_KINDS),
  sections: Object.keys(SECTION_KINDS),
  notes: Object.keys(NOTE_KINDS),
  paste: Object.keys(PASTE_BLOCK_KINDS),
};

// A fault path that does not start at its change reads, on a server, as a bad frame (a reload)
// instead of a refused change by name. Found by a host's media store.
describe('a refused change names its place in the message', () => {
  test('the families hold every kind once', () => {
    expect(Object.values(FAMILIES).flat().sort()).toEqual([...ORDERED_CHANGE_KINDS].sort());
  });

  for (const [family, kinds] of Object.entries(FAMILIES))
    test(`${family}: every fault in a change has a path inside that change`, () => {
      const wrong: string[] = [];
      let refused = 0;
      for (const kind of kinds) {
        const minimal = MINIMAL[kind] ?? {};
        // The optional fields a minimal form leaves out, each with a wrong value
        const absent = (orderedKindSpecOf(kind)?.fields ?? []).filter((key) => !(key in minimal));
        const cases = [
          ...Object.entries(minimal).filter(([key]) => key !== 'kind'),
          ...absent.flatMap((key) => [7, 'x', [], {}].map((bad) => [key, [bad]] as const)),
        ];
        for (const [key, good] of cases) {
          for (const bad of absent.includes(key) ? (good as unknown[]) : corrupted(good)) {
            const frame = {
              v: 3,
              type: 'changes',
              changes: [change(kind), change(kind, { [key]: bad }, `${NS}:2`)],
            };
            const result = decodeOrderedClientMessage(JSON.stringify(frame), { namespace: NS });
            if (result.ok) continue;
            refused += 1;
            // At the field, or at the change for a rule between two fields
            const { path } = result.error;
            if (path !== 'changes[1]' && !path.startsWith(`changes[1].${key}`))
              wrong.push(`${kind} ${key}: ${result.error.code} at ${result.error.path}`);
          }
        }
      }
      expect(wrong).toEqual([]);
      if (kinds.length > 0) expect(refused).toBeGreaterThan(kinds.length);
    });
});

describe('change ids', () => {
  test('parse, build and refuse', () => {
    expect(parseOrderedChangeId('ana-1:42')).toEqual({ namespace: 'ana-1', counter: 42 });
    expect(orderedChangeId('ana-1', 42)).toBe('ana-1:42');
    expect(() => orderedChangeId('a b', 1)).toThrow(RangeError);
    expect(() => orderedChangeId('ana', 0)).toThrow(RangeError);
    expect(() => orderedIdMinter('nope')).toThrow(RangeError);
  });

  test('every copy mints the same ids, each kind from its own counter', () => {
    const run = (order: readonly string[]) => {
      const ids = orderedIdMinter('ana-1:7');
      const out: Record<string, unknown[]> = {};
      for (const kind of order) {
        const minted =
          kind === 'paragraph'
            ? ids.paragraph()
            : kind === 'decimal'
              ? ids.decimal()
              : ids[kind as 'suggestion']();
        (out[kind] ??= []).push(minted);
      }
      return out;
    };
    const kinds = [
      'suggestion',
      'units',
      'paragraph',
      'container',
      'marker',
      'node',
      'record',
      'move',
      'relationship',
      'decimal',
    ];
    const a = run([...kinds, ...kinds]);
    const b = run([...kinds].reverse().concat([...kinds].reverse()));
    expect(b).toEqual(a);
    expect(a.suggestion).toEqual(['S:ana-1:7.1', 'S:ana-1:7.2']);
    expect(a.units).toEqual(['i:ana-1:7.1', 'i:ana-1:7.2']);
    const [paragraph] = a.paragraph as { space: string; paraId: string; textId: string }[];
    expect(paragraph?.space).toBe('k:ana-1:7.1');
    expect(isValidParaId(paragraph?.paraId ?? '')).toBe(true);
    expect(paragraph?.textId).toMatch(/^[0-9A-F]{8}$/);
    for (const rel of a.relationship as string[]) expect(rel).toMatch(/^rId[1-9]\d{8}$/);
    for (const n of a.decimal as number[]) expect(n >= 1e9 && n < 2 ** 31 - 1).toBe(true);
    const other = orderedIdMinter('ana-1:8');
    expect(other.suggestion()).not.toBe(orderedIdMinter('ana-1:7').suggestion());
    const used = new Set([orderedIdMinter('ana-1:7').paragraph().paraId]);
    expect(orderedIdMinter('ana-1:7').paragraph((id) => used.has(id)).paraId).not.toBe(
      [...used][0]
    );
  });
});
