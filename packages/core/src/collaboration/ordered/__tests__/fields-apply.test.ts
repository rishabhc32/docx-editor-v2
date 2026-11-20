// The field, bookmark and table of contents kinds (group 9.5) on the model, as the room applies
// them: the validator's bounds, two copies that apply the same changes in either order, refusals
// that change nothing, and a take-back that leaves the document as it was.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { inverseOfDirect } from '../client/undo-direct.ts';
import { decideOrdered, orderedCards } from '../decide.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedItem } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  DATE,
  DIRECT,
  EDIT,
  load,
  markOf,
  opens,
  partsOf,
  refusal,
  textOf,
} from './change-apply-fixture.ts';
import { docx, p, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const NS = 'ana-1';
const PLACE = { space: 'b:/word/document.xml#0.0.0', offset: 0 };

/** The validator's verdict on one change of a kind, with these fields. */
function verdict(kind: string, fields: Record<string, unknown>): string {
  const result = validateOrderedChange(
    { v: 1, id: `${NS}:1`, kind, part: BODY, mode: 'edit', ...fields },
    { namespace: NS, parts: new Set([BODY]) }
  );
  return result.ok ? 'ok' : result.error.code;
}

describe('the validator bounds the group 9.5 kinds', () => {
  test("a new bookmark's name follows Word's rules; a removed one is the file's", () => {
    const range = { from: PLACE, to: PLACE, remove: [] };
    expect(verdict('bookmark', { name: 'Terms_1', ...range })).toBe('ok');
    expect(verdict('bookmark', { name: '1st', ...range })).toBe('invalid-value');
    expect(verdict('bookmark', { name: 'a b', ...range })).toBe('invalid-value');
    expect(verdict('bookmark', { name: 'Old name 1', remove: [PLACE] })).toBe('ok');
    expect(verdict('bookmark', { name: 'x'.repeat(257), remove: [PLACE] })).toBe('too-long');
    expect(verdict('bookmark', { name: 'A', from: PLACE, remove: [] })).toBe('missing-field');
    expect(verdict('bookmark', { name: 'A', remove: [] })).toBe('empty');
  });

  test('a field result writes tabs and line breaks, no other control character', () => {
    const fields = (text: string) => ({ fields: [{ field: PLACE, text }] });
    expect(verdict('field-result', fields('4\t5\n6'))).toBe('ok');
    expect(verdict('field-result', fields('4\r5'))).toBe('invalid-text');
    expect(verdict('field-result', fields('x'.repeat(2049)))).toBe('too-long');
    expect(verdict('field-result', { fields: [] })).toBe('empty');
    expect(verdict('field-code', { field: PLACE, code: 'DATE' })).toBe('invalid-value');
    expect(verdict('insert-field', { place: PLACE, field: 'INCLUDETEXT' })).toBe('invalid-value');
  });

  test('a table of contents names a TOC field and bounded rows', () => {
    const toc = (extra: Record<string, unknown>) => ({
      before: PLACE,
      alias: 'Contents',
      instruction: 'TOC \\o "1-3" \\h',
      rows: [],
      bookmarks: [],
      ...extra,
    });
    expect(verdict('insert-toc', toc({}))).toBe('ok');
    expect(verdict('insert-toc', toc({ instruction: 'INCLUDETEXT "x"' }))).toBe('invalid-value');
    expect(verdict('insert-toc', toc({ alias: '' }))).toBe('empty');
    const row = { level: 9, text: 'A', heading: PLACE, bookmark: '_Toc1', page: '1' };
    expect(verdict('insert-toc', toc({ rows: [row] }))).toBe('invalid-number');
    expect(verdict('insert-toc', toc({ rows: Array(513).fill({ ...row, level: 0 }) }))).toBe(
      'too-many'
    );
  });
});

type Body = Parameters<typeof apply>[2];

describe('the group 9.5 kinds on the model', () => {
  test("a copy that applied its own changes first, then the room's order, writes the room's parts", () => {
    const fresh = load();
    const bookmark = (name: string, from: number, to: number): Body => ({
      kind: 'bookmark',
      ...DIRECT,
      name,
      from: at(fresh, 0, from),
      to: at(fresh, 0, to),
      remove: [],
    });
    const page: Body = { kind: 'insert-field', ...DIRECT, place: at(fresh, 0, 4), field: 'PAGE' };
    const pages: Body = {
      kind: 'insert-field',
      ...EDIT,
      place: at(fresh, 8, 0),
      field: 'PAGE_X_OF_Y',
    };
    const room: [string, Body, string][] = [
      ['olga-1:1', bookmark('Here', 4, 4), 'Olga'],
      ['olga-1:2', page, 'Olga'],
      ['ana-1:1', bookmark('Fee', 4, 7), 'Ana'],
      ['ana-1:2', pages, 'Ana'],
    ];
    const roomCopy = load();
    for (const [id, body, author] of room) apply(roomCopy, id, body, author);
    // Ana's copy: her changes first, taken back when the room's order arrives, then again
    const anaCopy = load();
    const own = room.filter(([, , author]) => author === 'Ana');
    const records = own.map(
      ([id, body, author]) =>
        withUndoRecord(anaCopy, () =>
          applyOrderedChange(anaCopy, { v: 1, id, ...body } as OrderedChange, {
            person: testPerson(author),
            date: DATE,
            ids: orderedIdMinter(id),
          })
        ).record
    );
    for (const record of [...records].reverse()) takeBack(anaCopy, record);
    for (const [id, body, author] of room) apply(anaCopy, id, body, author);
    const parts = partsOf(roomCopy);
    expect(partsOf(anaCopy)).toEqual(parts);
    expect(parts.join()).toContain('w:name="Fee"');
    expect(parts.join()).toContain('NUMPAGES');
  });

  test('a bookmark, a field code and a table of contents are direct edits only', () => {
    const doc = load();
    const bookmark = {
      kind: 'bookmark',
      ...EDIT,
      name: 'A',
      from: at(doc, 0, 0),
      to: at(doc, 0, 3),
      remove: [],
    } as const;
    expect(refusal(doc, 'ana-1:1', bookmark)).toBe('invalid-edit');
    const toc = {
      kind: 'insert-toc',
      ...EDIT,
      before: markOf(doc, 0),
      alias: 'Contents',
      instruction: 'TOC \\o "1-3"',
      rows: [],
      bookmarks: [],
    } as const;
    expect(refusal(doc, 'ana-1:2', toc)).toBe('invalid-edit');
    // A field's address must name a whole field
    expect(
      refusal(doc, 'ana-1:3', { kind: 'field-code', ...DIRECT, field: at(doc, 0, 1), code: 'PAGE' })
    ).toBe('unknown-address');
  });

  test('a table of contents goes between body blocks, and its take-back leaves nothing', () => {
    const doc = load();
    const before = partsOf(doc);
    const change = {
      v: 1,
      id: 'olga-1:1',
      kind: 'insert-toc',
      ...DIRECT,
      before: markOf(doc, 8),
      alias: 'Contents',
      instruction: 'TOC \\o "1-3" \\h',
      rows: [
        {
          level: 0,
          text: 'Due in 30 days.',
          heading: markOf(doc, 8),
          bookmark: '_Toc9',
          page: '1',
        },
      ],
      bookmarks: [{ heading: markOf(doc, 8), name: '_Toc9' }],
    } as OrderedChange;
    const taken = withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson('Olga'),
        date: DATE,
        ids: orderedIdMinter('olga-1:1'),
      })
    ).record;
    const story = partsOf(doc).join();
    expect(story).toContain('w:name="_Toc9"');
    expect(story).toContain('w:anchor="_Toc9"');
    expect(textOf(doc, 10)).toBe('Due in 30 days.');
    takeBack(doc, taken);
    expect(partsOf(doc)).toEqual(before);
    // Not inside a table or a control: the engine refuses it there too
    const inCell = { ...change, id: 'olga-1:2', before: markOf(doc, 4) } as OrderedChange;
    expect(refusal(doc, 'olga-1:2', inCell as never)).toBe('invalid-edit');
  });
});

describe('a link that Enter split', () => {
  test('goes whole with its own suggestion: a rejected proposal leaves no half (random session, seed 23)', () => {
    const doc = load();
    const [link] = opens(doc, 'hyperlink');
    if (!link) throw new Error('fixture');
    // Ana proposes a new target: "the terms" struck, a copy in a new link after it
    const proposed = apply(doc, 'ana-1:1', {
      kind: 'link-target',
      ...EDIT,
      link: { space: link.space, offset: link.offset },
      href: 'https://example.com/new',
    });
    expect(textOf(doc, 1)).toBe('See the termsthe terms.');
    // The organiser presses Enter inside the new link: it splits in two
    apply(doc, 'olga-1:1', { kind: 'split', ...DIRECT, place: at(doc, 1, 16) }, 'Olga');
    const suggestion = proposed.suggestions[0] ?? '';
    withUndoRecord(doc, () => setStatus(doc, suggestion, 'rejected'));
    // Both halves of the link go with it; the struck words show again
    expect(() => partsOf(doc)).not.toThrow();
    expect([textOf(doc, 1), textOf(doc, 2)]).toEqual(['See the terms', '.']);
  });

  test("is one link again when its author's deletion takes the Enter back over another person's words at the first half's end", () => {
    const doc = load();
    apply(doc, 'ana-1:1', { kind: 'split', ...EDIT, place: at(doc, 1, 7) });
    expect([textOf(doc, 1), textOf(doc, 2)]).toEqual(['See the', ' terms.']);

    // Bob's words at the first half's end go after its link; Ana's deletion takes her Enter back
    apply(
      doc,
      'bob-1:1',
      { kind: 'insert', ...EDIT, place: { mark: markOf(doc, 1) }, text: 'Z' },
      'Bob'
    );
    apply(doc, 'ana-1:2', { kind: 'delete', ...EDIT, from: at(doc, 1, 5), to: at(doc, 2, 2) });
    expect(textOf(doc, 1)).toBe('See theZ terms.');
    const body = partsOf(doc).join();
    expect(body.match(/<w:hyperlink /g)?.length).toBe(1);
  });
});

const BOB = 'w:author="Bob" w:date="2026-10-28T12:00:00Z"';

/** A complex field's runs; struck, as a move's source holds them. */
const fieldRuns = (code: string, struck = false) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:${struck ? 'delInstrText' : 'instrText'} xml:space="preserve"> ${code} ` +
  `</w:${struck ? 'delInstrText' : 'instrText'}></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  (struck ? '<w:r><w:delText>1</w:delText></w:r>' : t('1')) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** Bob's tracked move of a PAGE field, from the first paragraph into the second. */
const MOVED_FIELD = docx(
  p(
    t('Page ') +
      `<w:moveFromRangeStart w:id="1" ${BOB} w:name="m1"/><w:moveFrom w:id="2" ${BOB}>` +
      `${fieldRuns('PAGE', true)}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` +
      t(' of the terms.')
  ) +
    p(
      t('See ') +
        `<w:moveToRangeStart w:id="3" ${BOB} w:name="m1"/><w:moveTo w:id="4" ${BOB}>` +
        `${fieldRuns('PAGE')}</w:moveTo><w:moveToRangeEnd w:id="3"/>` +
        t('here.')
    )
);

/** The body after every card is decided one way; the document is restored. */
function decidedBody(doc: ReturnType<typeof load>, decision: 'accept' | 'reject'): string {
  const { record } = withUndoRecord(doc, () => {
    for (const card of orderedCards(doc)) decideOrdered(doc, card.suggestions[0] ?? '', decision);
  });
  const body = partsOf(doc).join();
  takeBack(doc, record);
  return body;
}

describe("a field in a pending move's destination", () => {
  test('is rewritten in the move, as the engine rewrites it in place, and goes with the move', () => {
    const doc = openOrderedWordFile(MOVED_FIELD, { file: 1 }).document;
    expect([textOf(doc, 0), textOf(doc, 1)]).toEqual([
      'Page \ufffc of the terms.',
      'See \ufffchere.',
    ]);
    const before = partsOf(doc);
    // The move's source is struck: the engine refuses a new instruction there too
    const source: Body = { kind: 'field-code', ...DIRECT, field: at(doc, 0, 5), code: 'NUMPAGES' };
    expect(refusal(doc, 'olga-1:1', source, 'Olga')).toBe('invalid-edit');
    const change = {
      v: 1,
      id: 'olga-1:2',
      kind: 'field-code',
      ...DIRECT,
      field: at(doc, 1, 4),
      code: 'NUMPAGES',
    } as OrderedChange;
    apply(doc, change.id, change as Body, 'Olga');
    const body = partsOf(doc).join();
    const moveTo = /<w:moveTo (?:(?!<\/w:moveTo>).)*<\/w:moveTo>/.exec(body)?.[0] ?? '';
    expect(moveTo).toContain(' NUMPAGES ');
    expect(body.match(/<w:moveTo /g)?.length).toBe(1);
    expect(/<w:moveFrom (?:(?!<\/w:moveFrom>).)*<\/w:moveFrom>/.exec(body)?.[0]).toContain(
      ' PAGE '
    );
    // Rejected, the move takes the new instruction with it; accepted, it keeps it
    expect(decidedBody(doc, 'reject')).not.toContain('NUMPAGES');
    expect(decidedBody(doc, 'reject')).toContain(' PAGE ');
    expect(decidedBody(doc, 'accept')).toContain(' NUMPAGES ');
    expect(decidedBody(doc, 'accept')).not.toContain(' PAGE ');
    // The organiser's undo: the new field deleted, the old one shown again
    inverseOfDirect(doc, [change]).forEach((inverse, k) =>
      apply(doc, `olga-1:${3 + k}`, { ...inverse } as Body, 'Olga')
    );
    expect(partsOf(doc)).toEqual(before);
  });
});

const fldChar = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (code: string) =>
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>`;

/**
 * A table of contents whose one row is a link holding the row's `PAGEREF` field, as Word writes
 * it, the TOC field's end in a paragraph of its own, the heading the row points to, and a `PAGE`
 * field over two paragraphs (a field in one paragraph loads whole, with no markers).
 */
const TOC_FILE = docx(
  p(
    fldChar('begin') +
      instr('TOC \\o "1-3" \\h \\z \\u') +
      fldChar('separate') +
      '<w:hyperlink w:anchor="_Toc1" w:history="1">' +
      t('Clause one') +
      '<w:r><w:tab/></w:r>' +
      fldChar('begin') +
      instr('PAGEREF _Toc1 \\h') +
      fldChar('separate') +
      t('1') +
      fldChar('end') +
      '</w:hyperlink>',
    '<w:pStyle w:val="TOC1"/>'
  ) +
    p(fldChar('end')) +
    p(
      '<w:bookmarkStart w:id="0" w:name="_Toc1"/>' + t('Clause one') + '<w:bookmarkEnd w:id="0"/>',
      '<w:pStyle w:val="Heading1"/>'
    ) +
    p(fldChar('begin') + instr('PAGE') + fldChar('separate') + t('1')) +
    p(fldChar('end'))
);

/** The `w:fldChar` items of the body, as `[type, item]`, in story order. */
function fieldChars(doc: ReturnType<typeof load>): [string, OrderedItem][] {
  const out: [string, OrderedItem][] = [];
  for (const item of doc.stories.get(BODY)?.tree.items() ?? []) {
    const node = item.node;
    if (item.kind !== 'zero' || !node || node.kind === 'textValue') continue;
    if (node.localName !== 'fldChar') continue;
    const type = node.attributes.find((a) => a.localName === 'fldCharType')?.value;
    if (type) out.push([type, item]);
  }
  return out;
}

describe("a table of contents' result", () => {
  const tocResult = (doc: ReturnType<typeof load>, separate: number, end: number): Body => {
    const chars = fieldChars(doc);
    const [, s] = chars[separate] ?? [];
    const [, e] = chars[end] ?? [];
    if (!s || !e) throw new Error('fixture');
    return {
      kind: 'toc-result',
      ...DIRECT,
      separate: addressOf(s),
      end: addressOf(e),
      instruction: 'TOC \\o "1-3" \\h \\z \\u',
      rows: [
        { level: 0, text: 'Clause one', heading: markOf(doc, 2), bookmark: '_Toc2', page: '2' },
      ],
      bookmarks: [{ heading: markOf(doc, 2), name: '_Toc2' }],
    };
  };
  // The fields' markers: the TOC's begin and separator (0, 1), the row's PAGEREF (2, 3, 4), the
  // TOC's end (5), and the PAGE field after the heading (6, 7, 8)
  const open = () => openOrderedWordFile(TOC_FILE, { file: 1 }).document;

  test("is refused when its end is a field's inside the result, not the TOC field's (x68-short seed 1012)", () => {
    const doc = open();
    expect(fieldChars(doc).map(([type]) => type)).toEqual([
      'begin',
      'separate',
      'begin',
      'separate',
      'end',
      'end',
      'begin',
      'separate',
      'end',
    ]);
    const before = partsOf(doc);
    // The row's PAGEREF end, inside the row's link
    expect(refusal(doc, 'olga-1:1', tocResult(doc, 1, 4), 'Olga')).toBe('not-one-field');
    expect(partsOf(doc)).toEqual(before);
  });

  test("is refused when its end is a later field's, past the TOC field's own end", () => {
    const doc = open();
    expect(refusal(doc, 'olga-1:1', tocResult(doc, 1, 8), 'Olga')).toBe('not-one-field');
  });

  test("is refused when its separator and end are a field's inside the result, not the TOC's", () => {
    const doc = open();
    // The PAGEREF's own pair: the rows would go inside the row's link
    expect(refusal(doc, 'olga-1:1', tocResult(doc, 3, 4), 'Olga')).toBe('not-one-field');
    // A PAGE field's own pair is no table of contents
    expect(refusal(doc, 'olga-1:2', tocResult(doc, 7, 8), 'Olga')).toBe('not-one-field');
  });

  test('is refused when its separator or its end is inside a link', () => {
    const toc = fldChar('begin') + instr('TOC \\o "1-3" \\h') + fldChar('separate');
    const link = (runs: string) => `<w:hyperlink w:anchor="_Toc1">${runs}</w:hyperlink>`;
    const heading = p(t('Clause one'), '<w:pStyle w:val="Heading1"/>');
    for (const body of [
      // The rows would go inside the link
      p(link(toc + t('Old row'))) + p(fldChar('end')) + heading,
      // The end's link would lose its open item
      p(toc + t('Old row')) + p(link(t('Old row') + fldChar('end'))) + heading,
    ]) {
      const doc = openOrderedWordFile(docx(body), { file: 1 }).document;
      const chars = fieldChars(doc);
      const separate = chars.find(([type]) => type === 'separate')?.[1];
      const end = chars.find(([type]) => type === 'end')?.[1];
      if (!separate || !end) throw new Error('fixture');
      const change: Body = {
        kind: 'toc-result',
        ...DIRECT,
        separate: addressOf(separate),
        end: addressOf(end),
        instruction: 'TOC \\o "1-3" \\h',
        rows: [
          { level: 0, text: 'Clause one', heading: markOf(doc, 2), bookmark: '_Toc2', page: '1' },
        ],
        bookmarks: [{ heading: markOf(doc, 2), name: '_Toc2' }],
      };
      expect(refusal(doc, 'olga-1:1', change, 'Olga')).toBe('not-one-field');
    }
  });

  test("with the TOC field's own separator and end, applies and composes", () => {
    const doc = open();
    apply(doc, 'olga-1:1', tocResult(doc, 1, 5), 'Olga');
    const body = partsOf(doc).join();
    expect(body).toContain('w:anchor="_Toc2"');
    expect(body).toContain('w:name="_Toc2"');
    // The row's tab is a positional tab, a marker with no character
    expect(textOf(doc, 0)).toBe('Clause one2');
  });
});

describe("a bookmark change's markers", () => {
  /** The shown markers of the body's `_Toc1` bookmark: its start and its end. */
  const markers = (doc: ReturnType<typeof load>): OrderedItem[] =>
    [...(doc.stories.get(BODY)?.tree.items() ?? [])].filter((item) => {
      const node = item.node;
      return (
        item.kind === 'zero' &&
        node !== null &&
        node.kind !== 'textValue' &&
        (node.localName === 'bookmarkStart' || node.localName === 'bookmarkEnd')
      );
    });

  test("are refused when an end is another bookmark's", () => {
    const doc = openOrderedWordFile(TOC_FILE, { file: 1 }).document;
    const [start, end] = markers(doc);
    if (!start || !end) throw new Error('fixture');
    // `_Toc1`'s end, named as a marker of another bookmark: the TOC row's target would break
    const other: Body = { kind: 'bookmark', ...DIRECT, name: 'Other', remove: [addressOf(end)] };
    expect(refusal(doc, 'olga-1:1', other, 'Olga')).toBe('invalid-edit');
    // Its own pair goes
    apply(
      doc,
      'olga-1:2',
      { kind: 'bookmark', ...DIRECT, name: '_Toc1', remove: [addressOf(start), addressOf(end)] },
      'Olga'
    );
    expect(partsOf(doc).join()).not.toContain('w:bookmarkEnd w:id="0"');
  });
});
