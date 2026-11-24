// A party's format or property change on another person's tracked words, as the host's check
// judges it (design D9): a format record in a run of their `w:ins` or `w:del` is nested in it,
// unless the format changes nothing there; a mark's, a row's or a table's record sits beside
// their revision mark.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart, orderedViewParagraph } from '../compose/compose.ts';
import { orderedPlaceAt, type OrderedPlace } from '../edits/places.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { actor, BODY, realApply, snapshot } from './room-harness.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const CY = actor('Cy', 'cy-1', 'party');
const EDIT = { part: BODY, mode: 'suggest' } as const;

const SENTENCE = 'Clause 3. The Buyer shall pay the fee within thirty days.';
const BOLD = { localName: 'b' };
const COLOUR = { localName: 'color', attributes: { val: 'C00000' } };
const JOHN = 'w:id="1" w:author="John Doe" w:date="2026-05-04T09:43:00Z"';
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/>' +
  `</w:tblGrid><w:tr><w:tc>${p(t('Cell text.'))}</w:tc></w:tr></w:tbl>`;

/** A room on a file with this body. */
function setup(body = p(t(SENTENCE)) + p(t('Second paragraph.'))) {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadOrderedPackage(packageOf(docx(body))),
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 10, 24) + 1000 * tick++,
  });
  const doc = room.document;
  const counters = new Map<string, number>();
  const send = (who: typeof ANA, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change }]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };

  /** The place before the n-th shown character of paragraph `index`, now. */
  const at = (index: number, n: number): OrderedPlace => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) {
      throw new Error(`no paragraph ${index}`);
    }

    return orderedPlaceAt(view, n);
  };
  const markOf = (index: number): OrderedPlace => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) {
      throw new Error(`no paragraph ${index}`);
    }

    return { space: view.mark.space, offset: view.mark.offset };
  };

  /** The open item of the n-th container of a kind. */
  const open = (token: string, n = 0): OrderedPlace => {
    const items = doc.stories
      .get(BODY)
      ?.tree.toArray()
      .filter((item) => item.kind === 'open' && item.token === token);
    const item = items?.[n];
    if (!item) {
      throw new Error(`no ${token} ${n}`);
    }

    return { space: item.space, offset: item.offset };
  };
  const format = (from: OrderedPlace, to: OrderedPlace, property: object = BOLD) => ({
    kind: 'format',
    ...EDIT,
    from,
    to,
    properties: [property],
  });
  const xml = () => serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
  return { room, doc, send, at, markOf, open, format, xml };
}

type Room = ReturnType<typeof setup>;

/** Refused as nested, and the document is as it was. */
function refusedNested(room: Room, who: typeof ANA, change: Record<string, unknown>): void {
  const before = snapshot(room.doc);
  expect(room.send(who, change)).toBe('policy-nested/nested');
  expect(snapshot(room.doc)).toEqual(before);
}

describe('a party’s format on another person’s tracked words', () => {
  // Today: Word nests Ana's `w:rPrChange` in their `w:ins` or `w:del`, and the check cannot
  // take it back alone (`nested`), whoever made the revision mark and whatever the property
  const OTHERS = [
    { name: 'a party', who: CY },
    { name: 'the organiser', who: OLGA },
  ];
  const SHAPES = {
    inserted: [
      ['part', 21, 24],
      ['all', 14, 31],
      ['around', 12, 33],
    ],
    struck: [
      ['part', 10, 12],
      ['all', 10, 20],
      ['around', 8, 22],
    ],
  } as const;
  for (const { name, who } of OTHERS) {
    for (const [kind, shapes] of Object.entries(SHAPES)) {
      test(`on ${name}’s ${kind} words: nested, in every shape`, () => {
        for (const [, from, to] of shapes) {
          for (const property of [BOLD, COLOUR]) {
            const room = setup();
            const { send, at, format } = room;
            const theirs =
              kind === 'inserted'
                ? { kind: 'insert', ...EDIT, place: at(0, 14), text: 'unless the Buyer ' }
                : { kind: 'delete', ...EDIT, from: at(0, 10), to: at(0, 20) };
            expect(send(who, theirs)).toBe('accepted');
            refusedNested(room, ANA, format(at(0, from), at(0, to), property));
          }
        }
      });
    }
  }

  test('on a Word file’s own insertion or deletion: nested', () => {
    // Their words start at 14 and at 10
    const files = [
      [p(t('Clause 3. The ') + `<w:ins ${JOHN}>${t('unless the Buyer ')}</w:ins>` + t('pay.')), 14],
      [
        p(
          t('Clause 3. ') +
            `<w:del ${JOHN}><w:r><w:delText xml:space="preserve">The Buyer </w:delText></w:r></w:del>` +
            t('shall pay.')
        ),
        10,
      ],
    ] as const;
    for (const [body, start] of files) {
      for (const property of [BOLD, COLOUR]) {
        const room = setup(body);
        refusedNested(
          room,
          ANA,
          room.format(room.at(0, start + 1), room.at(0, start + 4), property)
        );
      }
    }
  });

  // Today: a format that changes nothing writes no record, so the check finds nothing of hers
  test('that changes nothing there: allowed, with a record only beside their words', () => {
    // John Doe's inserted or struck words, already bold
    const files = [
      [p(t('Clause 3. The ') + `<w:ins ${JOHN}>${t('unless the Buyer ', '<w:b/>')}</w:ins>`), 14],
      [
        p(
          t('Clause 3. ') +
            `<w:del ${JOHN}><w:r><w:rPr><w:b/></w:rPr>` +
            '<w:delText xml:space="preserve">The Buyer </w:delText></w:r></w:del>' +
            t('shall pay.')
        ),
        10,
      ],
    ] as const;
    for (const [body, start] of files) {
      const { send, at, format, xml } = setup(body);
      expect(send(ANA, format(at(0, start + 1), at(0, start + 4)))).toBe('accepted');
      expect(xml()).not.toContain('w:rPrChange');
    }

    // Cy's " unless" typed at the end of a bold run, and the plain " shall" after it
    const room = setup(p(t('Clause 3. ') + t('The Buyer', '<w:b/>') + t(' shall pay.')));
    const { send, at, format, xml } = room;
    expect(send(CY, { kind: 'insert', ...EDIT, place: at(0, 19), text: ' unless' })).toBe(
      'accepted'
    );
    expect(send(ANA, format(at(0, 19), at(0, 26)))).toBe('accepted');
    expect(xml()).not.toContain('w:rPrChange');
    expect(send(ANA, format(at(0, 19), at(0, 32)))).toBe('accepted');
    expect(xml()).toMatch(/<\/w:ins><w:r><w:rPr><w:b\/><w:rPrChange [^>]*w:author="Ana"/);
  });

  test('in a run another person’s container holds: nested in their control, move and note', () => {
    const control = setup();
    expect(
      control.send(CY, {
        kind: 'wrap-control',
        ...EDIT,
        from: control.at(0, 4),
        to: control.at(0, 7),
        control: { type: 'richText' },
      })
    ).toBe('accepted');
    refusedNested(control, ANA, control.format(control.at(0, 8), control.at(0, 10)));

    const move = setup();
    const moved = {
      kind: 'move',
      ...EDIT,
      from: move.at(0, 10),
      to: move.at(0, 20),
      place: move.at(1, 0),
    };
    expect(move.send(CY, moved)).toBe('accepted');
    refusedNested(move, ANA, move.format(move.at(0, 11), move.at(0, 14)));
    refusedNested(move, ANA, move.format(move.at(1, 1), move.at(1, 4)));

    const note = setup();
    expect(
      note.send(CY, {
        kind: 'insert-note',
        ...EDIT,
        place: note.at(0, 8),
        note: 'footnote',
        text: 'Note words',
      })
    ).toBe('accepted');
    refusedNested(note, ANA, note.format(note.at(0, 8), note.at(0, 9)));
  });

  // Today: a new row's marker is the row's (`w:trPr/w:ins`): the organiser's words typed in it
  // directly have no `w:ins` of their own, and a format there sits beside Cy's row. Cy's own
  // words in it are in her `w:ins`
  test('in another person’s new row: nested in their words, not in the row', () => {
    for (const [who, mode, expected] of [
      [CY, 'suggest', 'policy-nested/nested'],
      [OLGA, 'edit', 'accepted'],
    ] as const) {
      const room = setup(p(t(SENTENCE)) + TABLE + p(t('After.')));
      const { send, open, format, doc } = room;
      expect(send(CY, { kind: 'insert-row', ...EDIT, row: open('tr'), where: 'below' })).toBe(
        'accepted'
      );
      const mark = doc.stories
        .get(BODY)
        ?.tree.toArray()
        .find((item) => item.kind === 'mark' && item.within.by.length > 0);
      const place = { mark: { space: mark?.space ?? '', offset: mark?.offset ?? 0 } };
      expect(send(who, { kind: 'insert', part: BODY, mode, place, text: 'Typed' })).toBe(
        'accepted'
      );
      const typed = doc.stories
        .get(BODY)
        ?.tree.toArray()
        .find((item) => item.kind === 'text' && item.text === 'Typed');
      const space = typed?.space ?? '';
      const offset = typed?.offset ?? 0;
      const change = format({ space, offset: offset + 1 }, { space, offset: offset + 3 });
      expect(send(ANA, change)).toBe(expected);
    }
  });

  // Today: Word writes a mark's, a paragraph's, a row's and a table's change record beside the
  // revision mark (`w:pPr/w:rPr`, `w:pPr`, `w:trPr`, `w:tblPr`), and takes it back alone
  test('a mark’s, a paragraph’s, a row’s or a table’s property change: allowed', () => {
    const struckMark = setup();
    expect(struckMark.send(CY, { kind: 'join', ...EDIT, mark: struckMark.markOf(0) })).toBe(
      'accepted'
    );
    const { send, markOf, at, format } = struckMark;
    expect(
      send(ANA, { kind: 'mark-format', ...EDIT, marks: [markOf(0)], properties: [BOLD] })
    ).toBe('accepted');
    const centre = { localName: 'jc', attributes: { val: 'center' } };
    expect(
      send(ANA, { kind: 'paragraph-property', ...EDIT, marks: [markOf(0)], properties: [centre] })
    ).toBe('accepted');

    // Words around Cy's struck mark: the mark's record and records on plain words
    expect(send(ANA, format(at(0, 50), at(1, 6), COLOUR))).toBe('accepted');

    const newParagraph = setup();
    expect(newParagraph.send(CY, { kind: 'split', ...EDIT, place: newParagraph.at(0, 9) })).toBe(
      'accepted'
    );
    expect(
      newParagraph.send(ANA, {
        kind: 'mark-format',
        ...EDIT,
        marks: [newParagraph.markOf(0)],
        properties: [BOLD],
      })
    ).toBe('accepted');

    const table = setup(p(t(SENTENCE)) + TABLE + p(t('After.')));
    expect(
      table.send(CY, { kind: 'insert-row', ...EDIT, row: table.open('tr'), where: 'below' })
    ).toBe('accepted');
    const height = { op: 'setTableRowHeight', heightTwips: 600 };
    expect(
      table.send(ANA, { kind: 'table-property', ...EDIT, on: [table.open('tr', 1)], op: height })
    ).toBe('accepted');
    expect(table.send(CY, { kind: 'delete-row', ...EDIT, row: table.open('tr') })).toBe('accepted');
    expect(
      table.send(ANA, { kind: 'table-property', ...EDIT, on: [table.open('tr')], op: height })
    ).toBe('accepted');

    // Words in Cy's struck row: the row's marker is the row's, not a `w:del` around them
    expect(table.send(ANA, table.format(table.at(1, 0), table.at(1, 4)))).toBe('accepted');
    const inserted = setup();
    expect(
      inserted.send(CY, {
        kind: 'insert-table',
        ...EDIT,
        place: inserted.at(1, 0),
        rows: 1,
        columns: 2,
        columnWidth: 2000,
      })
    ).toBe('accepted');
    const centred = { op: 'setTableProperties', justification: 'center' };
    expect(
      inserted.send(ANA, {
        kind: 'table-property',
        ...EDIT,
        on: [inserted.open('tbl')],
        op: centred,
      })
    ).toBe('accepted');
  });
});
