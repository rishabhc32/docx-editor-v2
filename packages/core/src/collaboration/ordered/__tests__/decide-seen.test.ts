// The decision seen check: a decision names every suggestion of the card the organiser saw,
// with the version they saw. An edit that touches the card in between (its items, its labels,
// or which suggestions it holds) refuses the decision with `stale-decision`, and nothing
// changes; an edit elsewhere does not.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { withUndoRecord, takeBack } from '../model/undo-record.ts';
import { encodeOrderedServerMessage, decodeOrderedServerMessage } from '../protocol/codec.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { apply, at, DATE, EDIT, load, refusal } from './change-apply-fixture.ts';
import { actor, realApply } from './room-harness.ts';
import { testPerson } from './test-persons.ts';

/** Ana replaces "100" with "250": one card of a deletion and an insertion. */
function replaced() {
  const doc = load();
  apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 11), to: at(doc, 0, 14) });
  apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, 0, 14), text: '250' });
  const seen = orderedCardSeen(doc, 'S:ana:2.1');
  if (!seen) throw new Error('no card');
  return { doc, seen };
}

describe('the decision seen check', () => {
  test('the card as the organiser saw it is decided whole', () => {
    const { doc, seen } = replaced();
    expect(seen.map((s) => s.suggestion)).toEqual(['S:ana:1.1', 'S:ana:2.1']);
    apply(
      doc,
      'olga:1',
      { kind: 'decide', cards: [{ suggestions: seen }], status: 'accepted' },
      'Olga'
    );
    expect(doc.suggestions.get('S:ana:1.1')?.status).toBe('accepted');
    expect(doc.suggestions.get('S:ana:2.1')?.status).toBe('accepted');
  });

  test('an edit of the card after the organiser saw it makes the decision stale', () => {
    const cases: [string, (doc: ReturnType<typeof load>) => void][] = [
      // Ana types on in her new word: the insertion's version moves
      [
        'the author extends it',
        (doc) => apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, 0, 17), text: '0' }),
      ],
      // Ben types inside Ana's new word: his insertion joins the card
      [
        'another person types inside it',
        (doc) =>
          apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 15), text: 'x' }, 'Ben'),
      ],
      // Ben formats Ana's new word: his format joins the card, and her units' labels change
      [
        'another person formats it',
        (doc) =>
          apply(
            doc,
            'ben:1',
            {
              kind: 'format',
              ...EDIT,
              from: at(doc, 0, 14),
              to: at(doc, 0, 17),
              properties: [{ localName: 'b' }],
            },
            'Ben'
          ),
      ],
      // Ana takes her deletion back: the card holds one suggestion now
      [
        'the author takes a part back',
        (doc) => apply(doc, 'ana:3', { kind: 'retract', suggestions: ['S:ana:1.1'] }),
      ],
    ];
    for (const [name, change] of cases) {
      const { doc, seen } = replaced();
      change(doc);
      const reason = refusal(
        doc,
        'olga:1',
        { kind: 'decide', cards: [{ suggestions: seen }], status: 'accepted' },
        'Olga'
      );
      expect({ name, reason }).toEqual({ name, reason: 'stale-decision' });
      expect(doc.suggestions.get('S:ana:2.1')?.status).toBe('pending');
    }
  });

  test('a card that grew is stale even when the versions the organiser saw still hold', () => {
    const { doc } = replaced();
    // The organiser saw only the insertion, before the deletion next to it joined its card
    const partial = (orderedCardSeen(doc, 'S:ana:2.1') ?? []).filter(
      (s) => s.suggestion === 'S:ana:2.1'
    );
    expect(
      refusal(
        doc,
        'olga:1',
        { kind: 'decide', cards: [{ suggestions: partial }], status: 'rejected' },
        'Olga'
      )
    ).toBe('stale-decision');
  });

  test('an edit elsewhere leaves the decision as it was', () => {
    const { doc, seen } = replaced();
    apply(doc, 'ben:1', { kind: 'insert', ...EDIT, place: at(doc, 8, 0), text: 'Ben ' }, 'Ben');
    apply(
      doc,
      'olga:1',
      { kind: 'decide', cards: [{ suggestions: seen }], status: 'rejected' },
      'Olga'
    );
    expect(doc.suggestions.get('S:ana:2.1')?.status).toBe('rejected');
  });

  test('versions go back with an undo record', () => {
    const { doc } = replaced();
    const version = () => doc.suggestions.get('S:ana:2.1')?.version;
    const before = version();
    const change = {
      v: 1,
      id: 'ana:3',
      kind: 'insert',
      ...EDIT,
      place: at(doc, 0, 17),
      text: '0',
    } as const;
    const { record } = withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson('Ana'),
        date: DATE,
        ids: orderedIdMinter('ana:3'),
      })
    );
    expect(version()).toBeGreaterThan(before ?? 0);
    takeBack(doc, record);
    expect(version()).toBe(before);
  });

  test('the room refuses a stale decision by name, and the reason reaches the organiser', () => {
    let tick = 0;
    const room = new OrderedRoom({
      document: load(),
      applyChange: realApply,
      policy: suggestionsOnlyPolicy(),
      now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
    });
    const ana = actor('Ana', 'ana-1', 'party');
    const olga = actor('Olga', 'olga-1', 'organiser');
    const doc = room.document;
    room.submit(ana, [
      { v: 1, id: 'ana-1:1', kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' },
    ]);
    const seen = orderedCardSeen(doc, 'S:ana-1:1.1');
    room.submit(ana, [
      { v: 1, id: 'ana-1:2', kind: 'insert', ...EDIT, place: at(doc, 0, 8), text: 'x' },
    ]);
    const result = room.submit(olga, [
      { v: 1, id: 'olga-1:1', kind: 'decide', cards: [{ suggestions: seen }], status: 'accepted' },
    ]);
    expect(result.refused.map((r) => [r.code, r.reason])).toEqual([
      ['apply-refused', 'stale-decision'],
    ]);
    const frame = encodeOrderedServerMessage({ v: 3, type: 'refuse', refusals: result.refused });
    const decoded = decodeOrderedServerMessage(frame);
    expect(decoded.ok && decoded.value.type === 'refuse' && decoded.value.refusals[0]?.reason).toBe(
      'stale-decision'
    );
    // The organiser looks again and decides the card as it is now
    const again = room.submit(olga, [
      {
        v: 1,
        id: 'olga-1:2',
        kind: 'decide',
        cards: [{ suggestions: orderedCardSeen(doc, 'S:ana-1:1.1') }],
        status: 'accepted',
      },
    ]);
    expect(again.refused).toEqual([]);
    expect(doc.suggestions.get('S:ana-1:1.1')?.status).toBe('accepted');
  });
});
