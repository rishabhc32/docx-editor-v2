// A party's direct edit inside a pending insertion (design D9). The engine types a direct
// insertion there into that insertion, as Word extends a tracked insertion with untracked
// typing: for the organiser, as before; for a party, refused as the file's own text refuses it.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../change/apply.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedPlaceAt, type OrderedPlace } from '../edits/places.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { addressOf, BODY, FOOTNOTES, load, opens } from './change-apply-fixture.ts';
import { actor } from './room-harness.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS } from './word-fixtures.ts';

const OLGA = actor('Olga', 'olga-1', 'organiser');
const ANA = actor('Ana', 'ana-1', 'party');
const BEN = actor('Ben', 'ben-1', 'party');
const SUGGEST = { part: BODY, mode: 'suggest' } as const;
const DIRECT = { part: BODY, mode: 'edit' } as const;

/** Where Ben's insertion goes in the fixture: a shown paragraph and an offset in it. */
const PLACES = [
  { name: 'body text', index: 0, n: 4 },
  { name: 'a link', index: 1, n: 8 },
  { name: 'a control', index: 2, n: 9 },
  { name: 'a cell', index: 4, n: 1 },
] as const;

/** A room with the suggestions-only policy, on the change apply fixture by default. */
function setup(document: OrderedDocument = load()) {
  let tick = 0;
  const room = new OrderedRoom({
    document,
    applyChange: applyOrderedChange,
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
  const at = (index: number, n: number, story = BODY): OrderedPlace => {
    const view = orderedViewParagraph(doc, story, index);
    if (!view) {
      throw new Error(`no paragraph ${index}`);
    }

    return orderedPlaceAt(view, n);
  };

  /** The characters from `n` of paragraph `index`, each with its suggestion's author. */
  const authors = (index: number, n: number, length: number, story = BODY) =>
    (orderedViewParagraph(doc, story, index)?.chars ?? [])
      .slice(n, n + length)
      .map((c) => `${c.ch}:${doc.suggestions.get(c.item.by ?? '')?.attribution.author ?? '-'}`);
  return { doc, send, at, authors };
}

describe('a direct edit inside a pending insertion', () => {
  for (const { name, index, n } of PLACES) {
    test(`in ${name}: refused for a party, joins it for the organiser`, () => {
      const { doc, send, at, authors } = setup();
      expect(send(BEN, { kind: 'insert', ...SUGGEST, place: at(index, n), text: 'abc' })).toBe(
        'accepted'
      );
      const inside = { from: at(index, n + 1), to: at(index, n + 2) };
      const bold = [{ localName: 'b' }];
      expect({
        tab: send(ANA, { kind: 'insert-tab', ...DIRECT, place: at(index, n + 1) }),
        delete: send(ANA, { kind: 'delete', ...DIRECT, ...inside }),
        format: send(ANA, { kind: 'format', ...DIRECT, ...inside, properties: bold }),
        insert: send(ANA, { kind: 'insert', ...DIRECT, place: at(index, n + 1), text: 'X' }),
      }).toEqual({
        tab: 'policy-direct/text',
        delete: 'policy-direct/text',
        format: 'policy-direct/package',
        insert: 'policy-direct/text',
      });
      expect(authors(index, n, 3)).toEqual(['a:Ben', 'b:Ben', 'c:Ben']);
      expect(doc.suggestions.size).toBe(1);

      // The organiser's direct words join Ben's insertion
      expect(send(OLGA, { kind: 'insert', ...DIRECT, place: at(index, n + 1), text: 'O' })).toBe(
        'accepted'
      );
      expect(authors(index, n, 4)).toEqual(['a:Ben', 'O:Ben', 'b:Ben', 'c:Ben']);
    });
  }

  test('in a row or a note another person made: refused at its start too', () => {
    const { doc, send, at, authors } = setup();
    const row = opens(doc, 'tr')[0];
    if (!row) {
      throw new Error('fixture');
    }

    expect(send(BEN, { kind: 'insert-row', ...SUGGEST, row: addressOf(row), where: 'above' })).toBe(
      'accepted'
    );
    expect(send(ANA, { kind: 'insert', ...DIRECT, place: at(4, 0), text: 'X' })).toBe(
      'policy-direct/text'
    );
    expect(send(OLGA, { kind: 'insert', ...DIRECT, place: at(4, 0), text: 'O' })).toBe('accepted');
    expect(authors(4, 0, 1)).toEqual(['O:Ben']);

    // Ben's new note is the last paragraph of the notes: its reference mark, then his words
    const note = { kind: 'insert-note', ...SUGGEST, note: 'footnote', text: 'Note' };
    expect(send(BEN, { ...note, place: at(0, 4) })).toBe('accepted');
    const direct = { kind: 'insert', part: FOOTNOTES, mode: 'edit', text: 'X' };
    expect([0, 1, 3].map((n) => send(ANA, { ...direct, place: at(3, n, FOOTNOTES) }))).toEqual([
      'policy-direct/text',
      'policy-direct/text',
      'policy-direct/text',
    ]);
  });

  test('in a text box another person made: refused at its start too', () => {
    const textBox = MATRIX_ITEMS.find((item) => item.key === 'textbox');
    if (!textBox) {
      throw new Error('fixture');
    }

    const opened = openOrderedWordFile(docx(bodyOf(textBox), MATRIX_EXTRAS), { file: 1 });
    const { doc, send, at } = setup(opened.document);
    const before = new Set([...doc.stories.keys()].filter(isTextBoxStoryName));

    // Ben's copy of the text box is his insertion, and its text box a story of its own
    const object = orderedViewParagraph(doc, BODY, 1)?.chars.find((c) => c.item.kind === 'atom');
    if (!object) {
      throw new Error('fixture');
    }

    const from = { space: object.address.space, offset: object.address.offset };
    expect(send(BEN, { kind: 'move', ...SUGGEST, from, to: at(1, 7), place: at(2, 0) })).toBe(
      'accepted'
    );
    const copy = [...doc.stories.keys()].find((n) => isTextBoxStoryName(n) && !before.has(n));
    if (!copy) {
      throw new Error('no copy');
    }

    const direct = { kind: 'insert', part: copy, mode: 'edit', text: 'X' };
    expect([0, 4, 8].map((n) => send(ANA, { ...direct, place: at(0, n, copy) }))).toEqual([
      'policy-direct/text',
      'policy-direct/text',
      'policy-direct/text',
    ]);
  });

  test('in the party’s own insertion: refused too, as the page sends a suggestion there', () => {
    const { send, at, authors } = setup();
    expect(send(ANA, { kind: 'insert', ...SUGGEST, place: at(0, 4), text: 'abc' })).toBe(
      'accepted'
    );
    expect(send(ANA, { kind: 'insert', ...DIRECT, place: at(0, 5), text: 'X' })).toBe(
      'policy-direct/text'
    );
    expect(send(ANA, { kind: 'insert', ...SUGGEST, place: at(0, 5), text: 'Y' })).toBe('accepted');
    expect(authors(0, 4, 4)).toEqual(['a:Ana', 'Y:Ana', 'b:Ana', 'c:Ana']);
  });
});
