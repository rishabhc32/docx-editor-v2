// The suggestions-only policy at a control's prompt that carries another person's pending
// suggestion (task 9.6; acceptance quick pass 101, cy-1:78). Content that takes the prompt's place
// drops the prompt's runs, as Word does, and the save keeps the prompt as words only: the other
// person's suggestion would be lost, and today's check refuses it (`others_suggestions`). The
// room refuses it as nested. For each kind of edit at each place by a prompt, the verdict is
// nested exactly when the edit, applied, loses that suggestion from the saved file.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedChangeTargets } from '../change/targets.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedCards } from '../decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { loadOrderedPackage } from '../index.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { addressOf, apply, at, markOf, opens, textOf } from './change-apply-fixture.ts';
import { reopened, saved } from './matrix/word.ts';
import { actor } from './room-harness.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

const BODY = '/word/document.xml';
const EDIT = { part: BODY, mode: 'suggest' } as const;
const FACE = '<w:rStyle w:val="PlaceholderText"/>';
const PROMPT = 'Click or tap here to enter text.';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** The host's reproduction: a plain text control showing Word's default prompt. */
const PLAIN = docx(
  p(t('Prompt: ') + sdt('<w:showingPlcHdr/><w:text/>', t(PROMPT, FACE))) + p(t('after'))
);
const INLINE = docx(
  p(t('Name: ') + sdt('<w:showingPlcHdr/>', t(PROMPT, FACE)) + t(' end.')) +
    p(t('Moved words here.'))
);
const BLOCK = docx(
  p(t('Name:')) + sdt('<w:showingPlcHdr/>', p(t(PROMPT, FACE))) + p(t('Moved words here.'))
);

const BOB = actor('Bob', 'bob-1', 'party');
const CY = actor('Cy', 'cy-1', 'party');
const OLGA = actor('Olga', 'olga-1', 'organiser');

/** The prompt's first unit, and a format change on its words " here to " (12 to 21). */
function promptOf(doc: OrderedDocument) {
  const words = doc.stories
    .get(BODY)
    ?.tree.toArray()
    .find((item) => item.kind === 'text' && item.prompt === true);
  if (!words) throw new Error('fixture');
  const place = (n: number) => ({ space: words.space, offset: words.offset + n });
  const bold = {
    kind: 'format',
    ...EDIT,
    from: place(12),
    to: place(21),
    properties: [{ localName: 'b' }],
  };
  return { place, bold };
}

function roomOn(file: Uint8Array) {
  let tick = 0;
  const room = new OrderedRoom({
    document: loadOrderedPackage(packageOf(file)),
    applyChange: (doc, change, ctx) => void applyOrderedChange(doc, change, ctx),
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 30) + 1000 * tick++,
  });
  const counters = new Map<string, number>();
  const send = (who: typeof BOB, change: Record<string, unknown>) => {
    const k = (counters.get(who.namespace) ?? 0) + 1;
    counters.set(who.namespace, k);
    const result = room.submit(who, [{ v: 1, id: `${who.namespace}:${k}`, ...change } as never]);
    const refusal = result.refused[0];
    return refusal ? `${refusal.code}${refusal.reason ? `/${refusal.reason}` : ''}` : 'accepted';
  };
  return { doc: room.document, send };
}

const bobs = (doc: OrderedDocument) => orderedCards(doc).filter((c) => c.author === 'Bob').length;

describe("typing over a prompt with another person's pending format (the host's case)", () => {
  test("a party's typing on the format or beside it is nested; on a control without it, accepted", () => {
    const { doc, send } = roomOn(PLAIN);
    const { place, bold } = promptOf(doc);
    expect(send(BOB, bold)).toBe('accepted');
    for (const n of [21, 0, 15])
      expect([n, send(CY, { kind: 'insert', ...EDIT, place: place(n), text: 'Q ' })]).toEqual([
        n,
        'policy-nested/nested',
      ]);
    expect(bobs(reopened(saved(packageOf(PLAIN), doc)))).toBe(1);
    const other = roomOn(PLAIN);
    const free = promptOf(other.doc);
    expect(other.send(CY, { kind: 'insert', ...EDIT, place: free.place(0), text: 'Q ' })).toBe(
      'accepted'
    );
  });

  test('the format’s author and the organiser may type over it; the format goes as in Word', () => {
    for (const who of [BOB, OLGA]) {
      const { doc, send } = roomOn(PLAIN);
      const { place, bold } = promptOf(doc);
      expect(send(BOB, bold)).toBe('accepted');
      const formatCard = () => orderedCards(doc).some((c) => c.suggestions.includes('S:bob-1:1.1'));
      expect(formatCard()).toBe(true);
      expect(send(who, { kind: 'insert', ...EDIT, place: place(0), text: 'Q ' })).toBe('accepted');
      expect(textOf(doc, 0)).toBe('Prompt: Q ');

      // No card on a prompt that content replaced: the room, the save and the editor alone agree.
      // Bob's card is his typing now; the organiser's typing is direct
      const left = who === BOB ? 1 : 0;
      expect(formatCard()).toBe(false);
      expect(bobs(doc)).toBe(left);
      expect(bobs(reopened(saved(packageOf(PLAIN), doc)))).toBe(left);
    }
  });

  // A deletion leaves a prompt's units (a prompt is not content), so Bob's format on them stays
  // as it was, and today's check finds nothing of his changed (allowed)
  test('a deletion that reaches the prompt leaves it, with the format on it', () => {
    for (const [file, line, from, to] of [
      [PLAIN, 0, 2, 8 + PROMPT.length],
      [BLOCK, 1, 0, PROMPT.length],
    ] as const) {
      const { doc, send } = roomOn(file);
      expect(send(BOB, promptOf(doc).bold)).toBe('accepted');
      const deletion = {
        kind: 'delete',
        ...EDIT,
        from: at(doc, line, from),
        to: at(doc, line, to),
      };
      expect(send(CY, deletion)).toBe('accepted');
      expect(bobs(doc)).toBe(1);
      expect(bobs(reopened(saved(packageOf(file), doc)))).toBe(1);
    }
  });

  test('once content replaced the prompt, a party types in the control freely', () => {
    const { doc, send } = roomOn(PLAIN);
    const { place, bold } = promptOf(doc);
    expect(send(BOB, bold)).toBe('accepted');
    expect(
      send(OLGA, { kind: 'insert', part: BODY, mode: 'edit', place: place(0), text: 'Q' })
    ).toBe('accepted');
    expect(send(CY, { kind: 'insert', ...EDIT, place: at(doc, 0, 9), text: 'R' })).toBe('accepted');
    expect(textOf(doc, 0)).toBe('Prompt: QR');
  });
});

type Place = { readonly name: string; readonly file: Uint8Array; readonly line: number } & (
  | { readonly n: number }
  | { readonly mark: true }
);
const PLACES: readonly Place[] = [
  { name: 'inline, its start', file: INLINE, line: 0, n: 6 },
  { name: 'inline, inside', file: INLINE, line: 0, n: 9 },
  { name: "inline, in the other person's words", file: INLINE, line: 0, n: 21 },
  { name: 'inline, its end', file: INLINE, line: 0, n: 6 + PROMPT.length },
  { name: 'block, its start', file: BLOCK, line: 1, n: 0 },
  { name: 'block, inside', file: BLOCK, line: 1, n: 3 },
  { name: 'block, its mark', file: BLOCK, line: 1, mark: true },
];

const control = (doc: OrderedDocument) =>
  addressOf((opens(doc, 'sdt')[0] ?? opens(doc, 'sdtBlock')[0])!);
const one = (text: string) => ({ e: 'w:p', c: [{ e: 'w:r', c: [{ e: 'w:t', c: [text] }] }] });
type Make = (doc: OrderedDocument, place: unknown, line: number) => Record<string, unknown>;
const KINDS: Readonly<Record<string, Make>> = {
  words: (_, place) => ({ kind: 'insert', ...EDIT, place, text: 'Q' }),
  tab: (_, place) => ({ kind: 'insert-tab', ...EDIT, place }),
  'line break': (_, place) => ({ kind: 'insert-break', ...EDIT, place, break: 'line' }),
  Enter: (_, place) => ({ kind: 'split', ...EDIT, place }),
  note: (_, place) => ({ kind: 'insert-note', ...EDIT, place, note: 'footnote', text: ' N' }),
  table: (_, place) => ({
    kind: 'insert-table',
    ...EDIT,
    place,
    rows: 1,
    columns: 1,
    columnWidth: 2400,
  }),
  field: (_, place) => ({ kind: 'insert-field', ...EDIT, place, field: 'PAGE' }),
  move: (doc, place, line) => ({
    kind: 'move',
    ...EDIT,
    from: at(doc, line + 1, 0),
    to: at(doc, line + 1, 5),
    place,
  }),
  'paste, one paragraph': (_, place) => ({
    kind: 'insert-fragment',
    ...EDIT,
    place,
    blocks: [one('P')],
  }),
  'paste, two paragraphs': (_, place) => ({
    kind: 'insert-fragment',
    ...EDIT,
    place,
    blocks: [one('P'), one('R')],
  }),
  value: (doc) => ({
    kind: 'control-value',
    ...EDIT,
    control: control(doc),
    value: { kind: 'text', text: 'V' },
  }),
  deletion: (doc, _, line) => ({
    kind: 'delete',
    ...EDIT,
    from: at(doc, line, line === 0 ? 8 : 2),
    to: at(doc, line, line === 0 ? 30 : 25),
  }),
};

describe('every kind of edit at every place by a prompt: nested exactly when the save would lose the format', () => {
  for (const place of PLACES)
    test(place.name, () => {
      const rows: string[] = [];
      for (const [name, make] of Object.entries(KINDS)) {
        const doc = loadOrderedPackage(packageOf(place.file));
        apply(doc, 'bob-1:1', promptOf(doc).bold as never, 'Bob');
        const where = 'mark' in place ? markOf(doc, place.line) : at(doc, place.line, place.n);
        const change = { v: 1, id: 'cy-1:1', ...make(doc, where, place.line) } as never;
        const targets = orderedChangeTargets(doc, change);
        if (!targets.ok) throw new Error(`${name}: ${targets.code}`);
        const verdict = suggestionsOnlyPolicy().check({
          actor: CY,
          change,
          touched: targets.touched,
          document: doc,
        });
        const nested = !verdict.ok && verdict.code === 'policy-nested';
        try {
          apply(doc, 'cy-1:1', make(doc, where, place.line) as never, 'Cy');
        } catch (error) {
          if (!(error instanceof OrderedRefusal)) throw error;
        }

        const lost = bobs(doc) === 0 || bobs(reopened(saved(packageOf(place.file), doc))) === 0;
        rows.push(
          `${name}: ${nested ? 'nested' : 'allowed'}${nested === lost ? '' : ' (differs)'}`
        );
      }

      expect(rows.filter((row) => row.endsWith('(differs)'))).toEqual([]);
    });
});

/** Bob's pending insertion of runs, as Word writes a suggested control's content. */
const insertedByBob = (runs: string) =>
  `<w:ins w:id="70" w:author="Bob" w:date="2026-10-28T08:00:00Z">${runs}</w:ins>`;
const CITY =
  '<w:id w:val="9204"/><w:showingPlcHdr/><w:comboBox><w:listItem w:value=""/>' +
  '<w:listItem w:displayText="Mumbai" w:value="Mumbai"/></w:comboBox>';

/**
 * "Courts at [Choose a city.] have jurisdiction.": a control's prompt that holds a combo box
 * showing its own, each control the file's own or Bob's pending insertion, in a block control
 * after "Venue:" or in an inline one after "Venue: ".
 */
function venue(block: boolean, outerBob: boolean, innerBob: boolean): Uint8Array {
  const outer = (runs: string) => (outerBob ? insertedByBob(runs) : runs);
  const city = t('Choose a city.', FACE);
  const prompt =
    outer(t('Courts at ', FACE)) +
    sdt(CITY, innerBob ? insertedByBob(city) : city) +
    outer(t(' have jurisdiction.', FACE));
  const props = '<w:id w:val="9203"/><w:showingPlcHdr/>';
  const first = block
    ? p(t('Venue:')) + sdt(props, p(prompt))
    : p(t('Venue: ') + sdt(props, prompt));
  return docx(first + p(t('Moved words here.')));
}

/** Bob's suggestions as a save writes them: his insertions' words, then his formats' words. */
function bobsWords(doc: OrderedDocument): string {
  const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
  const words = (runs: string) =>
    [...runs.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('');
  const inserted = [...xml.matchAll(/<w:ins [^>]*w:author="Bob"[^>]*>(.*?)<\/w:ins>/gs)];
  const formatted = [
    ...xml.matchAll(/<w:r>(?:(?!<\/w:r>).)*?w:rPrChange [^>]*w:author="Bob".*?<\/w:r>/gs),
  ];
  const insertions = inserted.map((m) => words(m[1] ?? '')).join('');
  const formats = formatted.map((m) => words(m[0])).join('');
  return `${insertions}|${formats}`;
}

/** Bob's pending bold on the second to fifth characters of the prompt text `words` starts. */
function boldOn(doc: OrderedDocument, words: string) {
  const item = doc.stories
    .get(BODY)
    ?.tree.toArray()
    .find((each) => each.kind === 'text' && each.prompt === true && each.text?.startsWith(words));
  if (!item) {
    throw new Error('fixture');
  }

  const place = (n: number) => ({ space: item.space, offset: item.offset + n });
  return {
    kind: 'format',
    ...EDIT,
    from: place(1),
    to: place(5),
    properties: [{ localName: 'b' }],
  };
}

const VENUES = [
  { name: 'inline', block: false, line: 0, start: 7 },
  { name: 'block', block: true, line: 1, start: 0 },
] as const;

const BOBS = [
  { name: 'his format on the outer prompt', bold: 'Courts', outer: false, inner: false },
  { name: 'his format on the inner prompt', bold: 'Choose', outer: false, inner: false },
  { name: 'the outer control his insertion', bold: null, outer: true, inner: false },
  { name: 'the inner control his insertion', bold: null, outer: false, inner: true },
  { name: 'both controls his insertion', bold: null, outer: true, inner: true },
] as const;

/** Places from the outer prompt's start. */
const VENUE_PLACES = [
  ['the outer prompt, its start', 0],
  ['the outer prompt, its middle', 5],
  ['the outer prompt, its end', 43],
  ['the inner prompt, its start', 10],
  ['the inner prompt, its middle', 15],
  ['the inner prompt, its end', 24],
] as const;

const VENUE_KINDS: Readonly<Record<string, Make>> = {
  ...KINDS,
  'value of the combo box': (doc) => ({
    kind: 'control-value',
    ...EDIT,
    control: addressOf(opens(doc, 'sdt').at(-1)!),
    value: { kind: 'listItem', value: 'Mumbai' },
  }),
};

/**
 * Each kind of edit at each place of the two prompts: the edits that lose a suggestion of Bob's
 * from the room or the saved file and are allowed, and those refused as nested that lose none.
 */
function venueVerdicts(file: Uint8Array, line: number, start: number, bold: string | null) {
  const lost: string[] = [];
  const kept: string[] = [];
  for (const [place, n] of VENUE_PLACES) {
    for (const [name, make] of Object.entries(VENUE_KINDS)) {
      const doc = loadOrderedPackage(packageOf(file));
      if (bold) {
        apply(doc, 'bob-1:1', boldOn(doc, bold) as never, 'Bob');
      }

      const before = bobsWords(doc);
      const where = at(doc, line, start + n);
      const change = { v: 1, id: 'cy-1:1', ...make(doc, where, line) } as never;
      const targets = orderedChangeTargets(doc, change);
      if (!targets.ok) {
        throw new Error(`${name}: ${targets.code}`);
      }

      // As the room judges: before the apply, then on what it did
      const policy = suggestionsOnlyPolicy();
      const input = { actor: CY, change, touched: targets.touched, document: doc };
      let verdict = policy.check(input);
      try {
        const outcome = apply(doc, 'cy-1:1', make(doc, where, line) as never, 'Cy');
        if (verdict.ok) {
          verdict = policy.checkApplied?.(input, outcome) ?? verdict;
        }
      } catch (error) {
        if (!(error instanceof OrderedRefusal)) {
          throw error;
        }
      }

      const nested = !verdict.ok && verdict.code === 'policy-nested';
      const saves = bobsWords(reopened(saved(packageOf(file), doc)));
      const loses = bobsWords(doc) !== before || saves !== before;
      if (loses && !nested) {
        lost.push(`${place}, ${name}`);
      }

      if (nested && !loses) {
        kept.push(`${place}, ${name}`);
      }
    }
  }

  return { lost, kept };
}

describe('every kind of edit at every place of a prompt that holds a control with its own', () => {
  for (const { name: control, block, line, start } of VENUES) {
    for (const bob of BOBS) {
      test(`${control}: ${bob.name}`, () => {
        const file = venue(block, bob.outer, bob.inner);
        const { lost, kept } = venueVerdicts(file, line, start, bob.bold);
        expect({ lost, kept }).toEqual({ lost: [], kept: [] });
      });
    }
  }
});

describe("in another person's new control, the nested verdict follows what the edit does", () => {
  /** Bob's words in the room and in its saved file, after Cy's `change`, and the verdict. */
  function afterCy(file: Uint8Array, change: (doc: OrderedDocument) => Record<string, unknown>) {
    const { doc, send } = roomOn(file);
    const before = bobsWords(doc);
    const verdict = send(CY, change(doc));
    const kept = bobsWords(doc) === before;
    const saves = bobsWords(reopened(saved(packageOf(file), doc))) === before;
    return { verdict, kept, saves };
  }

  test('a deletion over the prompt of his new control strikes nothing: allowed', () => {
    for (const [outer, inner] of [
      [true, false],
      [false, true],
      [true, true],
    ] as const) {
      const file = venue(false, outer, inner);
      const deletion = (doc: OrderedDocument) => ({
        kind: 'delete',
        ...EDIT,
        from: at(doc, 0, 8),
        to: at(doc, 0, 30),
      });
      expect([outer, inner, afterCy(file, deletion)]).toEqual([
        outer,
        inner,
        { verdict: 'accepted', kept: true, saves: true },
      ]);
    }
  });

  test('typing in his new combo box goes to the end of the outer prompt: allowed', () => {
    for (const { block, line, start } of VENUES) {
      const typing = (doc: OrderedDocument) => ({
        kind: 'insert',
        ...EDIT,
        place: at(doc, line, start + 15),
        text: 'Q',
      });
      expect([block, afterCy(venue(block, false, true), typing)]).toEqual([
        block,
        { verdict: 'accepted', kept: true, saves: true },
      ]);
    }
  });

  test('a table at the prompt of his new control goes beside the paragraph: allowed', () => {
    const table = (doc: OrderedDocument) => ({
      kind: 'insert-table',
      ...EDIT,
      place: at(doc, 0, 12),
      rows: 1,
      columns: 1,
      columnWidth: 2400,
    });
    expect(afterCy(venue(false, true, false), table)).toEqual({
      verdict: 'accepted',
      kept: true,
      saves: true,
    });
  });

  test('Enter in his new combo box meets the edit’s own refusal', () => {
    const enter = (doc: OrderedDocument) => ({ kind: 'split', ...EDIT, place: at(doc, 0, 22) });
    expect(afterCy(venue(false, false, true), enter).verdict).toBe(
      'apply-refused/break-in-container'
    );
  });

  test('content that lands in his new combo box is inside his insertion', () => {
    const tab = (doc: OrderedDocument) => ({ kind: 'insert-tab', ...EDIT, place: at(doc, 0, 22) });
    expect(afterCy(venue(false, false, true), tab).verdict).toBe('policy-nested/nested');
  });
});
