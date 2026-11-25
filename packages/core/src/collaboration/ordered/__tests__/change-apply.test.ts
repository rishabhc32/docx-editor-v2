// The change apply (task 4.2): every kind of change maps to the model's own operations, with
// ids from the change id, the author and the date from the stamp, and `mode: 'edit'` as a
// direct edit. Two copies that apply the same changes make the same document; a refused change
// is refused by name and changes nothing.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart, validateOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { applyOrderedEdit } from '../edits/apply.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import { snapshot } from './snapshot.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DATE,
  DIRECT,
  EDIT,
  FOOTNOTES,
  HEADER,
  load,
  markOf,
  opens,
  partsOf,
  refusal,
  textOf,
} from './change-apply-fixture.ts';
import { testAttribution, testPerson } from './test-persons.ts';

const xml = (doc: ReturnType<typeof load>, part = BODY) =>
  serializeOoxmlPart(composeOrderedPart(doc, part).part);

describe('the change apply', () => {
  test('the fixture reads as its paragraphs', () => {
    const doc = load();
    expect(Array.from({ length: 10 }, (_, k) => textOf(doc, k))).toEqual([
      'The fee is 100 USD.',
      'See the terms.',
      'Party: Acme',
      'First item',
      'A1',
      'B1',
      'A2',
      'B2',
      'Due in 30 days.',
      'Signed here.',
    ]);
  });

  test('every kind applies, and two copies that apply the same changes are the same', () => {
    const run = () => {
      const doc = load();
      const cells = opens(doc, 'tc');
      const rows = opens(doc, 'tr');
      const [table] = opens(doc, 'tbl');
      const [link] = opens(doc, 'hyperlink');
      const [control] = opens(doc, 'sdt');
      if (!table || !link || !control || !cells[0] || !cells[1] || !rows[0])
        throw new Error('fixture');
      const steps: [string, Parameters<typeof apply>[2]][] = [
        ['ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' }],
        ['ana:2', { kind: 'insert-tab', ...EDIT, place: at(doc, 0, 0) }],
        ['ana:3', { kind: 'insert-break', ...EDIT, place: at(doc, 8, 3), break: 'page' }],
        ['ana:4', { kind: 'insert-break', ...EDIT, place: at(doc, 8, 6), break: 'column' }],
        [
          'ana:5',
          { kind: 'insert-break', ...EDIT, place: at(doc, 9, 6), break: 'line', clear: 'all' },
        ],
        ['ana:6', { kind: 'delete', ...EDIT, from: at(doc, 8, 0), to: at(doc, 8, 2) }],
        [
          'ana:7',
          {
            kind: 'format',
            ...EDIT,
            from: at(doc, 9, 0),
            to: at(doc, 9, 6),
            properties: [{ localName: 'b' }],
          },
        ],
        ['ana:8', { kind: 'split', ...EDIT, place: at(doc, 9, 7) }],
        ['ana:9', { kind: 'join', ...EDIT, mark: markOf(doc, 0) }],
        [
          'ana:10',
          {
            kind: 'paragraph-property',
            ...EDIT,
            marks: [markOf(doc, 3)],
            properties: [
              { localName: 'jc', attributes: { val: 'center' } },
              {
                localName: 'tabs',
                children: [{ localName: 'tab', attributes: { val: 'right', pos: '9000' } }],
              },
              { localName: 'numPr', children: [{ localName: 'ilvl', attributes: { val: '1' } }] },
            ],
          },
        ],
        [
          'ana:12',
          {
            kind: 'table-property',
            ...EDIT,
            on: [addressOf(cells[0]), addressOf(cells[1])],
            op: { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } },
          },
        ],
        [
          'ana:15',
          {
            kind: 'control-value',
            ...EDIT,
            control: addressOf(control),
            value: { kind: 'text', text: 'Beta Ltd' },
          },
        ],
        [
          'ana:16',
          { kind: 'link-target', ...EDIT, link: addressOf(link), href: 'https://e.com/n' },
        ],
        [
          'ana:17',
          {
            kind: 'insert-note',
            ...EDIT,
            place: at(doc, 2, 6),
            note: 'footnote',
            text: ' Paid in full.',
          },
        ],
        [
          'ana:18',
          {
            kind: 'comment',
            part: BODY,
            from: at(doc, 2, 0),
            to: at(doc, 2, 5),
            text: 'Check the date.',
          },
        ],
        [
          'ana:19',
          {
            kind: 'insert-table',
            ...EDIT,
            place: at(doc, 9, 0),
            rows: 1,
            columns: 2,
            columnWidth: 2400,
          },
        ],
        ['ana:20', { kind: 'insert', ...DIRECT, place: at(doc, 2, 0), text: 'Direct ' }],
      ];
      const outcomes = steps.map(([id, body]) => apply(doc, id, body));
      return { doc, outcomes };
    };
    const one = run();
    const two = run();
    expect(snapshot(one.doc, false)).toEqual(snapshot(two.doc, false));
    expect(partsOf(one.doc)).toEqual(partsOf(two.doc));
    expect(one.outcomes.map((o) => o.suggestions.length > 0)).toEqual([
      ...Array.from({ length: 14 }, () => true),
      false,
      true,
      true,
    ]);
    for (const name of one.doc.stories.keys()) {
      const valid = validateOoxmlPart(composeOrderedPart(one.doc, name).part);
      expect({ name, ok: valid.ok }).toEqual({ name, ok: true });
    }
    const body = xml(one.doc);
    expect(body).toContain('<w:tab/>');
    expect(body).toContain('<w:br w:type="page"/>');
    expect(body).toContain('<w:br w:type="column"/>');
    expect(body).toContain('<w:br w:clear="all" w:type="textWrapping"/>');
    expect(body).toContain('<w:commentRangeStart');
    expect(body).toContain('<w:commentReference');
    expect(xml(one.doc, FOOTNOTES)).toContain('Paid in full.');
    expect(one.doc.suggestions.get('S:ana:20.1')?.status).toBe('accepted');
    expect(one.doc.suggestions.get('S:ana:1.1')?.attribution).toEqual(testAttribution('Ana', DATE));
    const added = [...one.doc.comments.values()].filter((comment) => comment.person !== null);
    expect(
      added.map(({ author, date, person, parent, done }) => ({
        author,
        date,
        person,
        parent,
        done,
      }))
    ).toEqual([
      { author: 'Ana', date: DATE, person: testPerson('Ana'), parent: null, done: false },
    ]);
  });

  test('each table change applies the same on two copies', () => {
    type Make = (doc: ReturnType<typeof load>) => Parameters<typeof apply>[2];
    const cell = (doc: ReturnType<typeof load>, k: number) => {
      const found = opens(doc, 'tc')[k];
      if (!found) throw new Error('no cell');
      return addressOf(found);
    };
    const row = (doc: ReturnType<typeof load>, k: number) => {
      const found = opens(doc, 'tr')[k];
      if (!found) throw new Error('no row');
      return addressOf(found);
    };
    const changes: [string, Make][] = [
      ['insert row', (doc) => ({ kind: 'insert-row', ...EDIT, row: row(doc, 0), where: 'below' })],
      ['delete row', (doc) => ({ kind: 'delete-row', ...EDIT, row: row(doc, 1) })],
      [
        'insert column',
        (doc) => ({ kind: 'insert-column', ...EDIT, cell: cell(doc, 1), where: 'right' }),
      ],
      ['delete column', (doc) => ({ kind: 'delete-column', ...EDIT, cell: cell(doc, 0) })],
      [
        'row height',
        (doc) => ({
          kind: 'table-property',
          ...EDIT,
          on: [row(doc, 0), row(doc, 1)],
          op: { op: 'setTableRowHeight', heightTwips: 600 },
        }),
      ],
      [
        'cell borders over a selection',
        (doc) => ({
          kind: 'table-property',
          ...EDIT,
          on: [cell(doc, 0), cell(doc, 1), cell(doc, 2), cell(doc, 3)],
          op: {
            op: 'setTableCellBorders',
            scope: 'outside',
            spec: { style: 'single', size: 8, color: { kind: 'hex', value: 'FF0000' } },
          },
        }),
      ],
      [
        'table width',
        (doc) => {
          const [table] = opens(doc, 'tbl');
          if (!table) throw new Error('no table');
          return {
            kind: 'table-property',
            ...EDIT,
            on: [addressOf(table)],
            op: { op: 'setTableProperties', preferredWidth: { type: 'pct', value: 5000 } },
          };
        },
      ],
    ];
    for (const [name, make] of changes) {
      const run = () => {
        const doc = load();
        const outcome = apply(doc, 'ana:1', make(doc));
        return { doc, outcome };
      };
      const one = run();
      const two = run();
      expect({ name, suggestions: [...new Set(one.outcome.suggestions)] }).toEqual({
        name,
        suggestions: ['S:ana:1.1'],
      });
      expect({ name, same: partsOf(one.doc).join() === partsOf(two.doc).join() }).toEqual({
        name,
        same: true,
      });
      expect(validateOoxmlPart(composeOrderedPart(one.doc, BODY).part).ok).toBe(true);
    }
  });

  test("a new column leaves the cells of another person's struck row with that one deletion", () => {
    // Policy 1, as Word keeps one `w:cellDel`: the view shows the row's cells struck, and the
    // engine's new column keeps those marks; they are the row's deletion, not the column's
    const doc = load();
    const row = opens(doc, 'tr')[1];
    if (!row) throw new Error('fixture');
    apply(doc, 'ana:1', { kind: 'delete-row', ...EDIT, row: addressOf(row) });
    const [cell] = opens(doc, 'tc');
    if (!cell) throw new Error('fixture');
    apply(
      doc,
      'ben:1',
      { kind: 'insert-column', ...EDIT, cell: addressOf(cell), where: 'right' },
      'Ben'
    );
    const struck = opens(doc, 'tc').filter((c) => c.by === null && c.dels.length > 0);
    expect(struck.map((c) => c.dels)).toEqual([['S:ana:1.1'], ['S:ana:1.1']]);
  });

  test('tabs and breaks read as the engine reads them', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert-tab', ...EDIT, place: at(doc, 0, 3) });
    apply(doc, 'ana:2', { kind: 'insert-break', ...EDIT, place: at(doc, 0, 5), break: 'line' });
    apply(doc, 'ana:3', { kind: 'insert-break', ...EDIT, place: at(doc, 0, 8), break: 'page' });
    apply(doc, 'ana:4', { kind: 'insert-break', ...EDIT, place: at(doc, 0, 10), break: 'column' });
    expect(textOf(doc, 0)).toBe('The\t \nfe\fe\n is 100 USD.');
  });

  test('a change is the edits it maps to, with the same ids', () => {
    const viaChange = load();
    apply(viaChange, 'ana:1', { kind: 'insert', ...EDIT, place: at(viaChange, 0, 4), text: 'x' });
    const viaEdit = load();
    applyOrderedEdit(
      {
        doc: viaEdit,
        story: BODY,
        attribution: testAttribution('Ana', DATE),
        ids: orderedIdMinter('ana:1'),
      },
      { kind: 'insert', at: at(viaEdit, 0, 4), text: 'x' }
    );
    expect(snapshot(viaChange, false)).toEqual(snapshot(viaEdit, false));
  });

  test('a header edit goes to the header story', () => {
    const doc = load();
    apply(doc, 'ana:1', {
      kind: 'insert',
      part: HEADER,
      mode: 'suggest',
      place: at(doc, 0, 0, HEADER),
      text: 'Draft ',
    });
    expect(textOf(doc, 0, HEADER)).toBe('Draft Header text');
  });

  test('every refusal is by name, and a refused change leaves nothing behind', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' });
    apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, 8, 0), text: 'Ben ' }, 'Ben');
    apply(doc, 'ana:2', { kind: 'delete', ...EDIT, from: at(doc, 9, 0), to: at(doc, 9, 6) });
    const cells = opens(doc, 'tc');
    const [first, , , last] = cells;
    if (!first || !last) throw new Error('fixture');
    const cases: [string, Parameters<typeof apply>[2], string?][] = [
      [
        'not-own-suggestion',
        { kind: 'insert', ...EDIT, into: 'S:ben:1.1', place: at(doc, 0, 0), text: 'x' },
      ],
      [
        'unknown-suggestion',
        { kind: 'insert', ...EDIT, into: 'S:none:1.1', place: at(doc, 0, 0), text: 'x' },
      ],
      [
        'invalid-edit',
        { kind: 'insert', ...DIRECT, into: 'S:ana:1.1', place: at(doc, 0, 0), text: 'x' },
      ],
      ['invalid-edit', { kind: 'insert', ...EDIT, place: at(doc, 0, 0), text: 'a\tb' }],
      ['no-paragraph-after', { kind: 'join', ...EDIT, mark: markOf(doc, 9) }],
      ['unknown-address', { kind: 'join', ...EDIT, mark: at(doc, 0, 2) }],
      [
        'invalid-edit',
        {
          kind: 'paragraph-property',
          ...EDIT,
          marks: [markOf(doc, 0)],
          properties: [
            { localName: 'tabs', children: [{ localName: 'tab', attributes: { val: 'num' } }] },
          ],
        },
      ],
      ...['constructor', '__proto__'].map((val): [string, Parameters<typeof apply>[2]] => [
        'invalid-edit',
        {
          kind: 'paragraph-property',
          ...EDIT,
          marks: [markOf(doc, 0)],
          properties: [
            {
              localName: 'tabs',
              children: [{ localName: 'tab', attributes: { val, pos: '720' } }],
            },
          ],
        },
      ]),
      [
        'not-in-table',
        {
          kind: 'table-property',
          ...EDIT,
          on: [addressOf(first), { space: 'b:none', offset: 0 }],
          op: { op: 'setTableCellVerticalAlignment', alignment: 'center' },
        },
      ],
      ['not-own-suggestion', { kind: 'retract', suggestions: ['S:ben:1.1'] }],
      ['unknown-suggestion', { kind: 'retract', suggestions: ['S:ana:1.1', 'S:none:1.1'] }],
      [
        'unknown-suggestion',
        {
          kind: 'decide',
          cards: [{ suggestions: [{ suggestion: 'S:none:1.1', version: 0 }] }],
          status: 'accepted',
        },
      ],
      [
        'empty-range',
        { kind: 'comment', part: BODY, from: at(doc, 0, 4), to: at(doc, 0, 3), text: 'Back' },
      ],
      [
        'invalid-edit',
        {
          kind: 'insert-break',
          ...EDIT,
          place: at(doc, 0, 0),
          break: 'page',
          clear: 'left',
        } as never,
      ],
    ];
    for (const [reason, body] of cases) {
      const label = JSON.stringify(body).slice(0, 90);
      // A cell the story does not hold is refused by the edit's own lookup
      const got = refusal(doc, 'ana:9', body);
      expect({ label, got }).toEqual({
        label,
        got: reason === 'not-in-table' ? 'unknown-address' : reason,
      });
    }
  });

  test('a change over several targets holds whole or not at all', () => {
    const doc = load();
    apply(doc, 'ana:1', { kind: 'join', ...EDIT, mark: markOf(doc, 1) });
    // The struck mark is accepted: the paragraph it ended no longer shows
    const struck = doc.stories
      .get(BODY)
      ?.tree.toArray()
      .find((item) => item.kind === 'mark' && item.dels.length > 0);
    if (!struck) throw new Error('no struck mark');
    apply(
      doc,
      'olga:1',
      {
        kind: 'decide',
        cards: [{ suggestions: orderedCardSeen(doc, 'S:ana:1.1') ?? [] }],
        status: 'accepted',
      },
      'Olga'
    );
    const hidden = addressOf(struck);
    expect(
      refusal(doc, 'ana:2', {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 0), hidden],
        properties: [{ localName: 'jc', attributes: { val: 'right' } }],
      })
    ).toBe('hidden-target');
    expect(refusal(doc, 'ana:3', { kind: 'join', ...EDIT, mark: hidden })).toBe('hidden-target');
    // Two take-backs: the second is decided, so the first stays too
    apply(doc, 'ana:4', { kind: 'insert', ...EDIT, place: at(doc, 0, 0), text: 'x' });
    expect(
      refusal(doc, 'ana:5', { kind: 'retract', suggestions: ['S:ana:4.1', 'S:ana:1.1'] })
    ).toBe('not-pending');
    expect(doc.suggestions.get('S:ana:4.1')?.status).toBe('pending');
  });

  test('a direct edit is accepted at once; a suggestion stays pending', () => {
    const doc = load();
    apply(
      doc,
      'olga:1',
      { kind: 'delete', ...DIRECT, from: at(doc, 0, 0), to: at(doc, 0, 4) },
      'Olga'
    );
    expect(textOf(doc, 0)).toBe('fee is 100 USD.');
    expect(xml(doc)).not.toContain('<w:del ');
    apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 0), to: at(doc, 0, 4) });
    expect(xml(doc)).toContain('<w:del ');
  });

  test("a mark format sets a paragraph mark's run properties, in the change's suggestion", () => {
    const doc = load();
    const last = 9;
    const outcome = apply(doc, 'ana:1', {
      kind: 'format',
      ...EDIT,
      from: at(doc, last, 0),
      to: { mark: markOf(doc, last) },
      properties: [{ localName: 'b' }],
    });
    expect(outcome.suggestions).toEqual(['S:ana:1.1']);
    apply(doc, 'ana:2', {
      kind: 'mark-format',
      ...EDIT,
      into: 'S:ana:1.1',
      marks: [markOf(doc, last)],
      properties: [{ localName: 'b' }],
    });
    const mark = /<w:pPr>(?:(?!<\/w:pPr>).)*<\/w:pPr><w:r(?:(?!<\/w:p>).)*Signed here/s.exec(
      xml(doc)
    )?.[0];
    expect(mark).toContain('<w:rPr><w:b/><w:rPrChange');
    // A direct mark format writes no record
    const direct = load();
    apply(
      direct,
      'olga:1',
      {
        kind: 'mark-format',
        ...DIRECT,
        marks: [markOf(direct, last)],
        properties: [{ localName: 'i' }],
      },
      'Olga'
    );
    expect(xml(direct)).toContain('<w:rPr><w:i/></w:rPr></w:pPr>');
    expect(xml(direct)).not.toContain('rPrChange');
    // A mark that no longer shows is refused by name
    const joined = load();
    apply(joined, 'olga:1', { kind: 'join', ...DIRECT, mark: markOf(joined, 8) }, 'Olga');
    expect(
      refusal(joined, 'ana:1', {
        kind: 'mark-format',
        ...EDIT,
        marks: [markOf(load(), 8)],
        properties: [{ localName: 'b' }],
      })
    ).toBe('hidden-target');
  });

  test('a link target can be a bookmark', () => {
    const doc = load();
    const [link] = opens(doc, 'hyperlink');
    if (!link) throw new Error('fixture');
    apply(doc, 'ana:1', { kind: 'link-target', ...EDIT, link: addressOf(link), anchor: 'Terms' });
    expect(xml(doc)).toContain('w:anchor="Terms"');
    expect([...doc.resources.added.values()].some((r) => r.kind === 'relationship')).toBe(false);
  });

  test('a break at a paragraph start leaves no node id twice in the part', () => {
    // The mark's face is copied under the edit's ids: the paragraph it splits keeps its own
    // nodes as the file had them, and a part holds each node id once
    const face = '<w:rPr><w:rFonts w:ascii="Times New Roman"/><w:b/><w:szCs w:val="24"/></w:rPr>';
    for (const pPr of [`<w:jc w:val="center"/>${face}`, face]) {
      const doc = loadOrderedPackage(packageOf(docx(p(t('Alpha beta.'), pPr) + p(t('Gamma.')))));
      apply(doc, 'ana:1', { kind: 'split', ...EDIT, place: at(doc, 0, 0) });
      expect(validateOoxmlPart(composeOrderedPart(doc, BODY).part).ok).toBe(true);
      expect(xml(doc)).toContain('Alpha beta.');
    }
  });

  test('a new note and a new table get Word paragraph ids, as the editor gives them on open', () => {
    // Opened as the room and every page open it: every paragraph has its ids
    const doc = openOrderedWordFile(APPLY_FILE, { file: 0 }).document;
    apply(doc, 'ana:1', {
      kind: 'insert-note',
      ...EDIT,
      place: at(doc, 0, 4),
      note: 'footnote',
      text: 'Paid in full.',
    });
    apply(doc, 'ana:2', {
      kind: 'insert-table',
      ...EDIT,
      place: at(doc, 1, 0),
      rows: 1,
      columns: 2,
      columnWidth: 2000,
    });
    const notes = xml(doc, FOOTNOTES);
    const note = notes.slice(notes.lastIndexOf('<w:footnote '));
    expect(note).toContain('Paid in full.');
    expect(note).toMatch(/<w:p [^>]*w14:paraId="[0-9A-F]{8}"/);
    const body = xml(doc);
    const table = body.slice(body.indexOf('<w:tbl>'), body.indexOf('</w:tbl>'));
    expect(table.match(/<w:p[ >]/g)?.length).toBe(2);
    expect(table.match(/<w:p [^>]*w14:paraId=/g)?.length).toBe(2);
  });

  test('a comment is taken back with its record', () => {
    const doc = load();
    const before = snapshot(doc, false);
    expect(
      refusal(doc, 'ana:1', {
        kind: 'comment',
        part: BODY,
        from: at(doc, 0, 0),
        to: { space: 'b:none', offset: 0 },
        text: 'Lost',
      })
    ).toBe('unknown-address');
    expect(snapshot(doc, false)).toEqual(before);
    expect([...doc.comments.values()].filter((comment) => comment.person !== null)).toEqual([]);
  });
});

describe("a suggested Enter inside another person's pending link", () => {
  /** Ana proposes a new target (a new link after the struck one); Olga suggests Enter in it. */
  function splitProposal() {
    const doc = load();
    const [link] = opens(doc, 'hyperlink');
    if (!link) throw new Error('fixture');
    const proposed = apply(doc, 'ana-1:1', {
      kind: 'link-target',
      ...EDIT,
      link: addressOf(link),
      href: 'https://example.com/new',
    });
    const split = apply(doc, 'olga-1:1', { kind: 'split', ...EDIT, place: at(doc, 1, 16) }, 'Olga');
    expect([textOf(doc, 1), textOf(doc, 2)]).toEqual(['See the termsthe', ' terms.']);
    return { doc, link: proposed.suggestions[0] ?? '', enter: split.suggestions[0] ?? '' };
  }

  test('both rejected: the paragraph reads as the file had it', () => {
    const { doc, link, enter } = splitProposal();
    withUndoRecord(doc, () => {
      setStatus(doc, link, 'rejected');
      setStatus(doc, enter, 'rejected');
    });
    expect(() => partsOf(doc)).not.toThrow();
    expect([textOf(doc, 1), textOf(doc, 2)]).toEqual(['See the terms.', 'Party: Acme']);
  });

  // The link rejected while the Enter stays (pending or accepted): both halves of the split link
  // go with it (the random session with pastes, seed 71, found the case)
  for (const enterStatus of ['pending', 'accepted'] as const)
    test(`the link rejected, the Enter ${enterStatus}: both halves of the link go`, () => {
      const { doc, link, enter } = splitProposal();
      withUndoRecord(doc, () => {
        if (enterStatus === 'accepted') setStatus(doc, enter, 'accepted');
        setStatus(doc, link, 'rejected');
      });
      expect(() => partsOf(doc)).not.toThrow();
      expect([textOf(doc, 1), textOf(doc, 2)]).toEqual(['See the terms', '.']);
    });
});
