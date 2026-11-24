// The suggestions-only policy against the host's earlier check (task 5.3): the verdicts the
// multi-party acceptance pass compared with that check, one rule each. The room judges where a
// change's words land, a control value by the words it strikes, a caret by the Word run it
// would split, and a property change on a struck paragraph mark as beside the deletion.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedCardOf } from '../card-of.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart, orderedViewParagraph } from '../compose/compose.ts';
import type { OrderedReplacedFormat } from '../edits/context.ts';
import { orderedPlaceAt, type OrderedPlace } from '../edits/places.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import type { OrderedDocument } from '../model/types.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { BODY, load, opens, addressOf } from './change-apply-fixture.ts';
import { actor } from './room-harness.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const BEN = actor('Ben', 'ben-1', 'party');
const OTTO = actor('Otto', 'otto-1', 'organiser');
const EDIT = { part: BODY, mode: 'suggest' } as const;

/** A room on the change apply fixture: paragraph 0 reads "The fee is 100 USD.", 2 holds a control. */
function setup(document: OrderedDocument = load()) {
  const replaced: OrderedReplacedFormat[] = [];
  let tick = 0;
  const room = new OrderedRoom({
    document,
    applyChange: (doc, change, ctx) => {
      replaced.push(...applyOrderedChange(doc, change, ctx).replaced);
    },
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
  });
  const counters = new Map<string, number>();
  const send = (who: typeof ANA, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change }]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  const doc = room.document;
  /** The place before the n-th character of shown paragraph `index`, now. */
  const at = (index: number, n: number): OrderedPlace => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) throw new Error(`no paragraph ${index}`);
    return orderedPlaceAt(view, n);
  };
  const markOf = (index: number): OrderedPlace => {
    const view = orderedViewParagraph(doc, BODY, index);
    if (!view) throw new Error(`no paragraph ${index}`);
    return { space: view.mark.space, offset: view.mark.offset };
  };
  const textOf = (index: number) =>
    orderedViewParagraph(doc, BODY, index)
      ?.chars.map((c) => c.ch)
      .join('');
  return { room, doc, send, at, markOf, textOf, replaced };
}

/** The text of paragraph 0 with each unit's suggestion: what lands where. */
function labelled(doc: OrderedDocument): string {
  const view = orderedViewParagraph(doc, BODY, 0);
  return (view?.chars ?? [])
    .map((c) => (c.item.by === null ? c.ch : `[${c.ch}:${c.item.by.split(':')[1]}]`))
    .join('');
}

describe('the policy and today’s check', () => {
  // Today's check reads no cards: Ben's run between Ana's `w:del` and `w:ins` is taken back
  // alone and leaves the file as it was (allowed)
  test('typing between another person’s deletion and their insertion is beside both', () => {
    const { doc, send, at } = setup();
    // Ana replaces "100" with "200": a deletion next to her insertion, one review card
    expect(send(ANA, { kind: 'delete', ...EDIT, from: at(0, 11), to: at(0, 14) })).toBe('accepted');
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, 14), text: '200' })).toBe('accepted');
    expect(orderedCardOf(doc, 'S:ana-1:1.1')?.suggestions).toEqual(['S:ana-1:1.1', 'S:ana-1:2.1']);
    // Between the two: Ben's own words, and Ana's deletion and insertion are two cards now
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 14), text: 'x' })).toBe('accepted');
    expect(orderedCardOf(doc, 'S:ana-1:1.1')?.suggestions).toEqual(['S:ana-1:1.1']);
    expect(orderedCardOf(doc, 'S:ben-1:1.1')?.suggestions).toEqual(['S:ben-1:1.1']);
    // Inside Ana's new "200" the words would split her run: inside her insertion
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 16), text: 'y' })).toBe(
      'policy-nested/nested'
    );
    // Ana may type in her own replacement, and the organiser anywhere
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, 16), text: 'z' })).toBe('accepted');
    expect(send(OLGA, { kind: 'insert', ...EDIT, place: at(0, 16), text: 'o' })).toBe('accepted');
  });

  test('two cards side by side are not one card', () => {
    const { send, at } = setup();
    // Ana strikes "The" and Ben strikes "fee": two people, two cards
    expect(send(ANA, { kind: 'delete', ...EDIT, from: at(0, 0), to: at(0, 3) })).toBe('accepted');
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 7) })).toBe('accepted');
    // Ana's two deletions, separated by a space only Ben struck, are her two cards
    expect(send(ANA, { kind: 'delete', ...EDIT, from: at(0, 8), to: at(0, 10) })).toBe('accepted');
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 8), text: 'q' })).toBe('accepted');
  });

  test('typing in struck text is judged where the words land, past it', () => {
    const { doc, send, at } = setup();
    // Ana inserts "new ", and the organiser strikes all of it
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'new ' })).toBe('accepted');
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 8) })).toBe('accepted');
    // Ben's caret is inside Ana's struck words; his words go after them, outside her insertion
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 5), text: 'X' })).toBe('accepted');
    expect(labelled(doc)).toBe('The [n:ana-1][e:ana-1][w:ana-1][ :ana-1][X:ben-1]fee is 100 USD.');
  });

  test('struck text that ends inside another person’s insertion still puts the words in it', () => {
    const { send, at } = setup();
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'abcd' })).toBe('accepted');
    // The organiser strikes "ab" of Ana's "abcd"
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 6) })).toBe('accepted');
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 5), text: 'X' })).toBe(
      'policy-nested/nested'
    );
  });

  // Today: Word nests the new run in the `w:sdt` or `w:hyperlink` that Ben's `w:ins` holds, so
  // taking back Ana's words alone does not give Ben's container back as it was. Whatever the
  // words around the caret are: Ben's other suggestion (r5 s25-long:387), or words a decision
  // accepted (r4 s2-short:1923). Past the container's end and at its start her words sit beside it
  const CONTAINERS = [
    {
      name: 'control',
      // "fee" struck, and a copy of it in Ben's new control at 7
      make: [{ kind: 'wrap-control', from: 4, to: 7, control: { type: 'date' } }],
      xy: 8,
      edges: [12, 7],
    },
    {
      name: 'link',
      // A suggested link holds its suggestion's own words: "abcd" at 4
      make: [
        { kind: 'insert', place: 4, text: 'abcd' },
        { kind: 'insert-link', from: 4, to: 8, href: 'https://example.com/' },
      ],
      xy: 6,
      edges: [10, 4],
    },
  ] as const;
  for (const { name, make, xy, edges } of CONTAINERS) {
    test(`typing in another person’s new ${name} around other words is inside it`, () => {
      const { doc, send, at } = setup();
      const places = (change: Record<string, unknown>) =>
        Object.fromEntries(
          Object.entries(change).map(([k, v]) => [k, typeof v === 'number' ? at(0, v) : v])
        );

      // Ben strikes "USD", makes the container, and adds "XY" in it to his deletion
      expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 15), to: at(0, 18) })).toBe(
        'accepted'
      );
      for (const change of make) {
        expect(send(BEN, { ...EDIT, ...places(change) })).toBe('accepted');
      }

      expect(
        send(BEN, { kind: 'insert', ...EDIT, place: at(0, xy), text: 'XY', into: 'S:ben-1:1.1' })
      ).toBe('accepted');
      const typed = (n: number) =>
        send(ANA, { kind: 'insert', ...EDIT, place: at(0, n), text: 'a' });

      // Between Ben's container's words and "XY", and inside "XY" once accepted
      expect(typed(xy)).toBe('policy-nested/nested');
      const card = orderedCardOf(doc, 'S:ben-1:1.1')?.suggestions ?? [];
      expect(card).toEqual(['S:ben-1:1.1']);
      const seen = card.map((id) => ({
        suggestion: id,
        version: doc.suggestions.get(id)?.version,
      }));
      expect(
        send(OLGA, { kind: 'decide', cards: [{ suggestions: seen }], status: 'accepted' })
      ).toBe('accepted');
      expect(typed(xy + 1)).toBe('policy-nested/nested');

      // Past its end and at its start the engine types beside it (r7 quick-short:273)
      const [end, start] = edges;
      expect({ end: typed(end), start: typed(start) }).toEqual({
        end: 'accepted',
        start: 'accepted',
      });
    });
  }

  // Today: at a link's or a control's start the engine puts the new items beside it, so the host's
  // check finds Ana's `w:ins` next to Ben's container (r7 quick-short:273)
  test('every edit at the start of another person’s new link is beside it', () => {
    const { doc, send, at, textOf } = setup();
    const first = () =>
      orderedViewParagraph(doc, BODY, 0)?.chars.findIndex((c) => c.containers !== '') ?? -1;
    const start = () => ({ ...EDIT, place: at(0, first()) });

    // The organiser strikes "fee", and Ben types "abcd" after it and links them
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 7) })).toBe('accepted');
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 7), text: 'abcd' })).toBe('accepted');
    expect(
      send(BEN, {
        kind: 'insert-link',
        ...EDIT,
        from: at(0, 7),
        to: at(0, 11),
        href: 'https://example.com/',
      })
    ).toBe('accepted');
    expect(send(ANA, { kind: 'insert', ...start(), text: 'X' })).toBe('accepted');
    expect(send(ANA, { kind: 'insert-tab', ...start() })).toBe('accepted');
    expect(send(ANA, { kind: 'insert-break', ...start(), break: 'line' })).toBe('accepted');
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    const link = /<w:hyperlink\b.*?<\/w:hyperlink>/s.exec(xml)?.[0] ?? '';
    expect({
      ana: xml.includes('w:author="Ana"'),
      inLink: link.includes('w:author="Ana"'),
    }).toEqual({ ana: true, inLink: false });

    // Between Ben's words it is inside his insertion
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, first() + 1), text: 'Y' })).toBe(
      'policy-nested/nested'
    );
    expect(send(ANA, { kind: 'split', ...start() })).toBe('accepted');
    expect(textOf(1)?.startsWith('abcd')).toBe(true);
  });

  // Today: the keyboard's `inside` types at a control's edge in it, inside Ben's insertion; else
  // the words sit beside his control. In a control, a link's start is in the control
  test('at the edges of another person’s new control and link, the words are judged where they land', () => {
    const { doc, send, at } = setup();

    // "fee" struck, a copy of it in Ben's new control at 7, and "ee" of the copy in his new link
    const wrap = {
      kind: 'wrap-control',
      from: at(0, 4),
      to: at(0, 7),
      control: { type: 'richText' },
    };
    expect(send(BEN, { ...EDIT, ...wrap })).toBe('accepted');
    const link = {
      kind: 'insert-link',
      from: at(0, 8),
      to: at(0, 10),
      href: 'https://example.com/',
    };
    expect(send(BEN, { ...EDIT, ...link })).toBe('accepted');
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const typed = (n: number, inside?: OrderedPlace) =>
      send(ANA, { kind: 'insert', ...EDIT, place: at(0, n), text: 'a', ...(inside && { inside }) });

    expect({ link: typed(8), start: typed(7, control), end: typed(10, control) }).toEqual({
      link: 'policy-nested/nested',
      start: 'policy-nested/nested',
      end: 'policy-nested/nested',
    });
    expect({ end: typed(10), start: typed(7) }).toEqual({ end: 'accepted', start: 'accepted' });
  });

  test('typing at the start of another person’s new link in the file’s control is in the control', () => {
    const file = docx(
      p(
        t('Party: ') +
          '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>' +
          `${t('Acme')}</w:sdtContent></w:sdt>`
      )
    );
    const { doc, send, at } = setup(loadOrderedPackage(packageOf(file)));

    // Ben's "abc" in the control, after "Ac", and his link over them
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 9), text: 'abc' })).toBe('accepted');
    const link = {
      kind: 'insert-link',
      from: at(0, 9),
      to: at(0, 12),
      href: 'https://example.com/',
    };
    expect(send(BEN, { ...EDIT, ...link })).toBe('accepted');
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const typed = (inside?: OrderedPlace) =>
      send(ANA, { kind: 'insert', ...EDIT, place: at(0, 9), text: 'a', ...(inside && { inside }) });
    expect({ plain: typed(), inside: typed(control) }).toEqual({
      plain: 'accepted',
      inside: 'accepted',
    });
  });

  test('a deletion of another person’s words that someone struck already does nothing there', () => {
    const { send, at, doc } = setup();
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'abcd' })).toBe('accepted');
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 7) })).toBe('accepted');
    const struck = () =>
      orderedViewParagraph(doc, BODY, 0)
        ?.chars.map((c) => c.item.dels.join(','))
        .join('|');
    const before = struck();
    // "abc" is struck by the organiser already: Ben's deletion leaves it as it is (policy 1)
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 5), to: at(0, 7) })).toBe('accepted');
    expect(struck()).toBe(before);
    // Reaching the "d" no one struck is inside Ana's insertion
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 6), to: at(0, 8) })).toBe(
      'policy-nested/nested'
    );
  });

  test('a control value that strikes another person’s pending words is inside their suggestion', () => {
    const value = (doc: OrderedDocument, text: string) => {
      const [control] = opens(doc, 'sdt');
      if (!control) throw new Error('fixture');
      return {
        kind: 'control-value',
        ...EDIT,
        control: addressOf(control),
        value: { kind: 'text', text },
      };
    };
    // On the file's words a party's value is a suggestion
    const first = setup();
    expect(first.textOf(2)).toBe('Party: Acme');
    expect(first.send(BEN, value(first.doc, 'Beta Ltd'))).toBe('accepted');

    const { doc, send, at, textOf } = setup();
    // Ana types in the control: a value now would strike her pending words
    expect(send(ANA, { kind: 'insert', ...EDIT, place: at(2, 9), text: 'X' })).toBe('accepted');
    expect(textOf(2)).toBe('Party: AcXme');
    expect(send(BEN, value(doc, 'Gamma'))).toBe('policy-nested/nested');
    // Her own value takes her words back and strikes the file's
    expect(send(ANA, value(doc, 'Delta'))).toBe('accepted');
    // Ben's value would strike Ana's value; the organiser's may
    expect(send(BEN, value(doc, 'Gamma'))).toBe('policy-nested/nested');
    expect(send(OLGA, value(doc, 'Epsilon'))).toBe('accepted');
  });

  test('a new row beside a row with another person’s pending cell format copies no record of it', () => {
    const { doc, send } = setup();
    const rows = opens(doc, 'tr');
    const cells = opens(doc, 'tc');
    const [first, second] = rows;
    if (!first || !second || !cells[0] || !cells[1]) throw new Error('fixture');
    // Ben shades the first row's cells: a new row beside it copies the shown fill and no record
    // of his, as the engine's `insertTableRow` does (today's check allows it)
    const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } };
    expect(
      send(BEN, {
        kind: 'table-property',
        ...EDIT,
        on: [addressOf(cells[0]), addressOf(cells[1])],
        op: fill,
      })
    ).toBe('accepted');
    expect(send(ANA, { kind: 'insert-row', ...EDIT, row: addressOf(first), where: 'below' })).toBe(
      'accepted'
    );
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    const [, added] = xml.match(/<w:tr[ >].*?<\/w:tr>/gs) ?? [];
    expect(added).toContain('w:fill="FFCC00"');
    expect(added).not.toContain('w:author="Ben"');
    expect(send(ANA, { kind: 'insert-row', ...EDIT, row: addressOf(second), where: 'below' })).toBe(
      'accepted'
    );
  });

  // r5 s2-short:2750. Ana's new row copies the shown cell properties of the row it goes above,
  // Ben's struck one: its first cell continues the merge that starts above it, in the file's
  // row. An accept and a reject of Ben's deletion leave her row as it was, so it is not inside
  // his suggestion (today's check calls it nested: a fault of that check)
  test('a new row in a vertical merge, above another person’s struck row, is beside it', () => {
    const merged = (merge: string, text: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${merge}</w:tcPr>${p(t(text))}</w:tc>`;
    const row = (merge: string, text: string) =>
      `<w:tr>${merged(merge, text)}${merged('', `${text} fee`)}</w:tr>`;
    const file = docx(
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
        row('<w:vMerge w:val="restart"/>', 'Licence') +
        row('<w:vMerge/>', 'Two') +
        row('<w:vMerge/>', 'Three') +
        '</w:tbl>' +
        p(t('After.'))
    );
    const anasRow = (doc: OrderedDocument) =>
      /<w:tr><w:trPr><w:ins [^>]*w:author="Ana"[^>]*\/><\/w:trPr>.*?<\/w:tr>/
        .exec(serializeOoxmlPart(composeOrderedPart(doc, BODY).part))?.[0]
        .replace(/ w:(id|date|paraId|textId)="[^"]*"/g, '');
    for (const status of ['accepted', 'rejected'] as const) {
      const { doc, send } = setup(loadOrderedPackage(packageOf(file)));
      const second = () => addressOf(opens(doc, 'tr')[1]!);
      expect(send(BEN, { kind: 'delete-row', ...EDIT, row: second() })).toBe('accepted');
      expect(send(ANA, { kind: 'insert-row', ...EDIT, row: second(), where: 'above' })).toBe(
        'accepted'
      );
      const before = anasRow(doc);
      expect(before).toContain('<w:vMerge/>');
      const card = orderedCardOf(doc, 'S:ben-1:1.1')?.suggestions ?? [];
      expect(card).toEqual(['S:ben-1:1.1']);
      const seen = card.map((id) => ({
        suggestion: id,
        version: doc.suggestions.get(id)?.version,
      }));
      expect(send(OLGA, { kind: 'decide', cards: [{ suggestions: seen }], status })).toBe(
        'accepted'
      );
      expect({ status, row: anasRow(doc) }).toEqual({ status, row: before });
    }
  });

  test('a column change cuts through another person’s new row', () => {
    const tc = (doc: OrderedDocument, k = 0) => {
      const item = opens(doc, 'tc')[k];
      if (!item) throw new Error('fixture');
      return addressOf(item);
    };
    // With only the file's rows, a party's column is a suggestion
    const first = setup();
    expect(
      first.send(ANA, { kind: 'insert-column', ...EDIT, cell: tc(first.doc), where: 'left' })
    ).toBe('accepted');

    const { doc, send } = setup();
    const [row] = opens(doc, 'tr');
    if (!row) throw new Error('fixture');
    // Ben adds a row: a new column, or a column struck, would change a cell in his row
    expect(send(BEN, { kind: 'insert-row', ...EDIT, row: addressOf(row), where: 'below' })).toBe(
      'accepted'
    );
    expect(send(ANA, { kind: 'insert-column', ...EDIT, cell: tc(doc, 1), where: 'right' })).toBe(
      'policy-nested/nested'
    );
    expect(send(ANA, { kind: 'delete-column', ...EDIT, cell: tc(doc, 1) })).toBe(
      'policy-nested/nested'
    );
    // A row beside his row is beside it
    expect(send(ANA, { kind: 'insert-row', ...EDIT, row: addressOf(row), where: 'above' })).toBe(
      'accepted'
    );
    expect(send(OLGA, { kind: 'insert-column', ...EDIT, cell: tc(doc, 1), where: 'right' })).toBe(
      'accepted'
    );
  });

  test('a column that takes its width from another person’s new cell is inside their column', () => {
    const { doc, send } = setup();
    const tc = (k: number) => {
      const item = opens(doc, 'tc')[k];
      if (!item) throw new Error('fixture');
      return addressOf(item);
    };
    expect(send(ANA, { kind: 'insert-column', ...EDIT, cell: tc(0), where: 'right' })).toBe(
      'accepted'
    );
    // Right of Ana's new cell, the new column halves it (today's check: her group changed)
    expect(send(BEN, { kind: 'insert-column', ...EDIT, cell: tc(1), where: 'right' })).toBe(
      'policy-nested/nested'
    );
    // Left of it, the width comes from A1, the file's cell (today's check allows it)
    expect(send(BEN, { kind: 'insert-column', ...EDIT, cell: tc(1), where: 'left' })).toBe(
      'accepted'
    );
    // The organiser is not held by the party rules, suggesting or editing directly
    const anas = () => opens(doc, 'tc').findIndex((item) => item.by === 'S:ana-1:1.1');
    expect(send(OLGA, { kind: 'insert-column', ...EDIT, cell: tc(anas()), where: 'right' })).toBe(
      'accepted'
    );
    expect(
      send(OLGA, {
        kind: 'insert-column',
        part: BODY,
        mode: 'edit',
        cell: tc(anas()),
        where: 'right',
      })
    ).toBe('accepted');
  });

  test('a column over another person’s pending row height changes their records', () => {
    const { doc, send } = setup();
    const tc = (k: number) => addressOf(opens(doc, 'tc')[k]!);
    const [first] = opens(doc, 'tr');
    if (!first) throw new Error('fixture');
    const height = { op: 'setTableRowHeight', heightTwips: 480 };
    expect(send(BEN, { kind: 'table-property', ...EDIT, on: [addressOf(first)], op: height })).toBe(
      'accepted'
    );
    // Today's check: rejecting Ana's column alone would leave Ben's records with the lent
    // width, and the record on her new cell ties his row height to her column
    expect(send(ANA, { kind: 'insert-column', ...EDIT, cell: tc(0), where: 'right' })).toBe(
      'policy-nested/nested'
    );
    expect(send(ANA, { kind: 'insert-column', ...EDIT, cell: tc(2), where: 'right' })).toBe(
      'policy-nested/nested'
    );
    // The organiser's column keeps both changes whole, whichever is decided first
    expect(send(OLGA, { kind: 'insert-column', ...EDIT, cell: tc(0), where: 'right' })).toBe(
      'accepted'
    );
    // His row height on the other row: a column puts a cell in every row
    const { doc: other, send: again } = setup();
    expect(
      again(BEN, {
        kind: 'table-property',
        ...EDIT,
        on: [addressOf(opens(other, 'tr')[1]!)],
        op: height,
      })
    ).toBe('accepted');
    expect(
      again(ANA, {
        kind: 'insert-column',
        ...EDIT,
        cell: addressOf(opens(other, 'tc')[0]!),
        where: 'right',
      })
    ).toBe('policy-nested/nested');
  });

  // Word's rule (Word for Mac, 2026-10-30, `~$c2-as.docx`): a second author's table style takes
  // the table's one `w:tblPrChange` over, and the first author's name is gone from the file.
  // The room takes the table's record over too, but keeps the other person's pending row height
  // on its row's own `w:tblPrExChange` (Word's row form), so no person's change loses its
  // record (D17). Today's check finds the height's own record gone (`others_suggestions`), so a
  // party's style there is refused; the organiser's takes the table's record over.
  test('a table style over another person’s pending row height keeps their height recorded', () => {
    const { doc, send } = setup();
    const [first] = opens(doc, 'tr');
    if (!first) throw new Error('fixture');
    const height = { op: 'setTableRowHeight', heightTwips: 480 };
    expect(send(BEN, { kind: 'table-property', ...EDIT, on: [addressOf(first)], op: height })).toBe(
      'accepted'
    );
    const style = { op: 'setTableProperties', preferredWidth: { type: 'pct', value: 5000 } };
    const on = [addressOf(opens(doc, 'tbl')[0]!)];
    expect(send(ANA, { kind: 'table-property', ...EDIT, on, op: style })).toBe(
      'policy-not-own/others_suggestions'
    );
    expect(send(OLGA, { kind: 'table-property', ...EDIT, on, op: style })).toBe('accepted');
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    const authors = (name: string) =>
      [...xml.matchAll(new RegExp(`<w:${name} [^>]*w:author="(\\w+)"`, 'g'))].map((m) => m[1]);
    expect(authors('tblPrChange')).toEqual(['Olga']);
    expect(authors('tblPrExChange')).toEqual(['Ben']);
  });

  test('a new column through a row another person struck is inside their deletion', () => {
    const column = (where: 'left' | 'right', k: number) => {
      const { doc, send } = setup();
      const second = opens(doc, 'tr')[1];
      if (!second) throw new Error('fixture');
      expect(send(BEN, { kind: 'delete-row', ...EDIT, row: addressOf(second) })).toBe('accepted');
      const cell = addressOf(opens(doc, 'tc')[k]!);
      return [
        send(ANA, { kind: 'insert-column', ...EDIT, cell, where }),
        send(OLGA, { kind: 'insert-column', ...EDIT, cell, where }),
      ];
    };
    // Its cell in Ben's struck row sits inside his deletion: today's check cannot take the
    // column back alone. At the left edge the check passes it by accident (the cell comes
    // before the deletion's range); the room refuses it in every place
    for (const [where, k] of [
      ['left', 0],
      ['right', 0],
      ['left', 1],
      ['right', 1],
    ] as const)
      expect(column(where, k)).toEqual(['policy-nested/nested', 'accepted']);
    // Her own struck row is hers to change
    const { doc, send } = setup();
    expect(send(ANA, { kind: 'delete-row', ...EDIT, row: addressOf(opens(doc, 'tr')[1]!) })).toBe(
      'accepted'
    );
    expect(
      send(ANA, {
        kind: 'insert-column',
        ...EDIT,
        cell: addressOf(opens(doc, 'tc')[0]!),
        where: 'right',
      })
    ).toBe('accepted');
  });

  test('a format on the words of a row another person struck sits beside their deletion', () => {
    const { doc, send, textOf } = setup();
    const [row] = opens(doc, 'tr');
    if (!row) throw new Error('fixture');
    expect(send(BEN, { kind: 'delete-row', ...EDIT, row: addressOf(row) })).toBe('accepted');
    // "A1" in the struck row: Word marks the row, not a deletion around the words
    expect(textOf(4)).toBe('A1');
    const view = orderedViewParagraph(doc, BODY, 4);
    if (!view) throw new Error('fixture');
    const words = { from: orderedPlaceAt(view, 0), to: orderedPlaceAt(view, 2) };
    expect(send(ANA, { kind: 'format', ...EDIT, ...words, properties: [{ localName: 'b' }] })).toBe(
      'accepted'
    );
  });

  test('a new link target strikes the link’s words: over another person’s words it is inside', () => {
    const { doc, send, at, textOf } = setup();
    const target = () => {
      const [link] = opens(doc, 'hyperlink');
      if (!link) throw new Error('fixture');
      return { kind: 'link-target', ...EDIT, link: addressOf(link), href: 'https://e.com/n' };
    };
    expect(textOf(1)).toBe('See the terms.');
    // Ben types in the link's words; Ana's new target would strike them
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(1, 7), text: 'x' })).toBe('accepted');
    expect(send(ANA, target())).toBe('policy-nested/nested');
    expect(send(OLGA, target())).toBe('accepted');
  });

  test('Enter in a paragraph with another person’s pending paragraph format is inside it', () => {
    const { send, at, markOf } = setup();
    const jc = { localName: 'jc', attributes: { val: 'center' } };
    // "Due in 30 days.": no one's pending format
    const due = at(8, 3);
    expect(
      send(ANA, { kind: 'paragraph-property', ...EDIT, marks: [markOf(0)], properties: [jc] })
    ).toBe('accepted');
    // The new paragraph would carry Ana's pending format under her name
    expect(send(BEN, { kind: 'split', ...EDIT, place: at(0, 4) })).toBe('policy-nested/nested');
    expect(send(ANA, { kind: 'split', ...EDIT, place: at(0, 4) })).toBe('accepted');
    expect(send(BEN, { kind: 'split', ...EDIT, place: due })).toBe('accepted');
  });

  test('a paragraph property on a mark another person struck is beside the deletion', () => {
    const { send, at, markOf } = setup();
    // The organiser proposes joining paragraphs 0 and 1: she strikes paragraph 0's mark
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 19), to: at(1, 0) })).toBe('accepted');
    const style = { localName: 'jc', attributes: { val: 'center' } };
    expect(
      send(BEN, { kind: 'paragraph-property', ...EDIT, marks: [markOf(0)], properties: [style] })
    ).toBe('accepted');
    // A format on struck text goes in the deletion (Word nests its record there), and today's
    // check cannot take it back alone (`room-policy-formats.test.ts`)
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 0), to: at(0, 3) })).toBe('accepted');
    expect(
      send(BEN, {
        kind: 'format',
        ...EDIT,
        from: at(0, 1),
        to: at(0, 2),
        properties: [{ localName: 'b' }],
      })
    ).toBe('policy-nested/nested');
  });

  // Word writes Ben's `w:pPrChange` beside Ana's `w:ins` on the mark, and today's check takes
  // it back alone (allowed). It joins Ana's card: one decision for her paragraph
  test('a paragraph property on another person’s new paragraph is beside their insertion', () => {
    const { doc, send, at, markOf } = setup();
    // Ana breaks paragraph 0 after "The fee": the first part ends with her new mark
    expect(send(ANA, { kind: 'split', ...EDIT, place: at(0, 7) })).toBe('accepted');
    const style = { localName: 'pStyle', attributes: { val: 'Heading1' } };
    expect(
      send(BEN, { kind: 'paragraph-property', ...EDIT, marks: [markOf(0)], properties: [style] })
    ).toBe('accepted');
    expect(orderedCardOf(doc, 'S:ana-1:1.1')?.suggestions).toEqual(['S:ana-1:1.1', 'S:ben-1:1.1']);
  });

  // Word keeps one `w:pPrChange`, so the second takes the first over (policy 2). Today's check
  // finds the first person's suggestion gone (`others_suggestions`): a party may not do it
  test('a paragraph format over another person’s pending one replaces it and reports it', () => {
    const { doc, send, markOf, replaced } = setup();
    const jc = (val: string) => ({ localName: 'jc', attributes: { val } });
    expect(
      send(ANA, {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(0)],
        properties: [jc('center')],
      })
    ).toBe('accepted');
    expect(replaced).toEqual([]);
    const right = {
      kind: 'paragraph-property',
      ...EDIT,
      marks: [markOf(0)],
      properties: [jc('right')],
    };
    // Refused once applied, the take-over taken back with it: Ana's card is as it was
    expect(send(BEN, right)).toBe('policy-not-own/others_suggestions');
    expect(orderedCardOf(doc, 'S:ana-1:1.1')?.suggestions).toEqual(['S:ana-1:1.1']);
    expect(send(OLGA, right)).toBe('accepted');
    expect(replaced.at(-1)).toEqual({ suggestion: 'S:ana-1:1.1', by: 'S:olga-1:1.1' });
  });
});

describe('a take-back of one’s own words under another person’s deletion (s2-short:1000)', () => {
  /** The words a suggestion struck, in paragraph 0. */
  const struckBy = (doc: OrderedDocument, id: string) =>
    (orderedViewParagraph(doc, BODY, 0)?.chars ?? [])
      .filter((c) => c.item.dels.includes(id))
      .map((c) => c.ch)
      .join('');

  test('a party may not shrink the organiser’s deletion over their right-to-left words', () => {
    const { doc, send, at, textOf } = setup();
    // Ben types "עד " (step 1488), and the organiser strikes it with "fee " (step 1489)
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'עד ' })).toBe('accepted');
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 11) })).toBe('accepted');
    expect(struckBy(doc, 'S:olga-1:1.1')).toBe('עד fee ');
    // Ben's deletion over it takes his words back, and her deletion would lose them (step 1492)
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 3), to: at(0, 13) })).toBe(
      'policy-not-own/others_suggestions'
    );
    expect(struckBy(doc, 'S:olga-1:1.1')).toBe('עד fee ');
    expect(textOf(0)).toBe('The עד fee is 100 USD.');
  });

  test('a party may not take back their left-to-right words from under it either', () => {
    const { doc, send, at } = setup();
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'new ' })).toBe('accepted');
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 6), to: at(0, 12) })).toBe('accepted');
    expect(struckBy(doc, 'S:olga-1:1.1')).toBe('w fee ');
    // Only the struck "w" of Ben's words is under her deletion
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 6), to: at(0, 7) })).toBe(
      'policy-not-own/others_suggestions'
    );
    // The words of his no one struck are his to take back
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 6) })).toBe('accepted');
    expect(struckBy(doc, 'S:olga-1:1.1')).toBe('w fee ');
  });

  test('a party takes back their own words that no one struck', () => {
    const { send, at, textOf } = setup();
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'עד ' })).toBe('accepted');
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 7) })).toBe('accepted');
    expect(textOf(0)).toBe('The fee is 100 USD.');
  });

  test('the organiser may take back her own words under another person’s deletion', () => {
    const { doc, send, at, textOf } = setup();
    expect(send(OLGA, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'עד ' })).toBe('accepted');
    expect(send(OTTO, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 11) })).toBe('accepted');
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 3), to: at(0, 13) })).toBe('accepted');
    // Her words go; the struck text stays in view
    expect(struckBy(doc, 'S:otto-1:1.1')).toBe('fee ');
    expect(textOf(0)).toBe('The fee is 100 USD.');
  });

  test('nor take back their words or their mark with another person’s format or deletion on it', () => {
    const { send, at, markOf } = setup();
    const bold = { properties: [{ localName: 'b' }] };
    const jc = { localName: 'jc', attributes: { val: 'center' } };
    // Ben types "new " and breaks paragraph 0 after "The fee"; the organiser formats his words
    expect(send(BEN, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'new ' })).toBe('accepted');
    expect(send(BEN, { kind: 'split', ...EDIT, place: at(0, 11) })).toBe('accepted');
    expect(send(OLGA, { kind: 'format', ...EDIT, from: at(0, 4), to: at(0, 7), ...bold })).toBe(
      'accepted'
    );
    // Her format on his "new" would go with them, by a deletion or a take-back
    const notOwn = 'policy-not-own/others_suggestions';
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 8) })).toBe(notOwn);
    expect(send(BEN, { kind: 'retract', suggestions: ['S:ben-1:1.1'] })).toBe(notOwn);
    // The " " she left alone is his to take back
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 7), to: at(0, 8) })).toBe('accepted');
    // Her paragraph property on his new mark would go with the join
    expect(
      send(OLGA, { kind: 'paragraph-property', ...EDIT, marks: [markOf(0)], properties: [jc] })
    ).toBe('accepted');
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 10), to: at(1, 0) })).toBe(notOwn);
    expect(send(BEN, { kind: 'retract', suggestions: ['S:ben-1:2.1'] })).toBe(notOwn);
  });

  // Word nests the organiser's `w:rPrChange` in Ben's deleted run: today's take-back of the
  // deletion takes her format too, so its two views differ (`package`)
  test('nor take back their deletion with another person’s format nested in it', () => {
    const { send, at } = setup();
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 4), to: at(0, 7) })).toBe('accepted');
    const bold = [{ localName: 'b' }];
    expect(
      send(OLGA, { kind: 'format', ...EDIT, from: at(0, 4), to: at(0, 7), properties: bold })
    ).toBe('accepted');
    expect(send(BEN, { kind: 'retract', suggestions: ['S:ben-1:1.1'] })).toBe(
      'policy-not-own/others_suggestions'
    );
    expect(send(OLGA, { kind: 'retract', suggestions: ['S:olga-1:1.1'] })).toBe('accepted');
    expect(send(BEN, { kind: 'retract', suggestions: ['S:ben-1:1.1'] })).toBe('accepted');
  });

  test('nor join over a new mark of theirs that another person struck', () => {
    const { send, at } = setup();
    expect(send(BEN, { kind: 'split', ...EDIT, place: at(0, 7) })).toBe('accepted');
    // The organiser proposes the join back: she strikes his new mark
    expect(send(OLGA, { kind: 'delete', ...EDIT, from: at(0, 7), to: at(1, 0) })).toBe('accepted');
    const notOwn = 'policy-not-own/others_suggestions';
    expect(send(BEN, { kind: 'delete', ...EDIT, from: at(0, 7), to: at(1, 0) })).toBe(notOwn);
    expect(send(BEN, { kind: 'retract', suggestions: ['S:ben-1:1.1'] })).toBe(notOwn);
  });

  test('the organiser may take back her own words under another person’s format', () => {
    const { send, at, textOf } = setup();
    expect(send(OLGA, { kind: 'insert', ...EDIT, place: at(0, 4), text: 'new ' })).toBe('accepted');
    expect(
      send(OTTO, {
        kind: 'format',
        ...EDIT,
        from: at(0, 4),
        to: at(0, 7),
        properties: [{ localName: 'i' }],
      })
    ).toBe('accepted');
    expect(send(OLGA, { kind: 'retract', suggestions: ['S:olga-1:1.1'] })).toBe('accepted');
    expect(textOf(0)).toBe('The fee is 100 USD.');
  });
});
