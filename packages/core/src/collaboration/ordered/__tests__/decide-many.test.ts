// A decision of many cards is one change: every card of an Accept All or a Reject All, each as
// the organiser saw it. The seen check is per card: a card that changed since stays as it is,
// the others are decided, and only a change that leaves every card is refused by name.

import { describe, expect, test } from 'bun:test';
import { orderedDecidedCards, orderedSeenCards } from '../change/apply-decide.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedCards } from '../decide.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedDecideChange } from '../change/types.ts';
import { loadOrderedPackage } from '../load/load.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import { apply, at, BODY, DATE, EDIT, refusal, textOf } from './change-apply-fixture.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { actor, realApply } from './room-harness.ts';
import { snapshot } from './snapshot.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const CLAUSES = 30;
const ORIGINAL = Array.from({ length: CLAUSES }, (_, k) => `Clause ${k} reads as agreed.`);
const FILE = docx(ORIGINAL.map((text) => p(t(text))).join(''));

const texts = (doc: OrderedDocument) => ORIGINAL.map((_, k) => textOf(doc, k));
/** The suggestions the cards hold, as the review shows them. */
const pending = (doc: OrderedDocument) => orderedCards(doc).flatMap((card) => card.suggestions);

/**
 * One card in each clause: Ana's new word, Ben's deletion, or Ana's replacement (a deletion and
 * an insertion, one card).
 */
function suggested(): OrderedDocument {
  const doc = loadOrderedPackage(packageOf(FILE));
  for (let k = 0; k < CLAUSES; k += 1) {
    const offset = `Clause ${k} `.length;
    if (k % 3 === 0)
      apply(doc, `ana:${100 + k}`, { kind: 'insert', ...EDIT, place: at(doc, k, 0), text: 'New ' });
    else if (k % 3 === 1)
      apply(
        doc,
        `ben:${100 + k}`,
        { kind: 'delete', ...EDIT, from: at(doc, k, offset), to: at(doc, k, offset + 6) },
        'Ben'
      );
    else {
      const end = ORIGINAL[k]!.length - 1;
      apply(doc, `ana:${100 + k}`, {
        kind: 'delete',
        ...EDIT,
        from: at(doc, k, end - 6),
        to: at(doc, k, end),
      });
      apply(doc, `ana:${200 + k}`, {
        kind: 'insert',
        ...EDIT,
        place: at(doc, k, end),
        text: 'signed',
      });
    }
  }
  return doc;
}

/** Every card the organiser's copy shows, as one decision. */
function every(doc: OrderedDocument, status: 'accepted' | 'rejected') {
  return { kind: 'decide', cards: orderedSeenCards(doc, pending(doc)), status } as const;
}

const ACCEPTED = ORIGINAL.map((text, k) =>
  k % 3 === 0
    ? `New ${text}`
    : k % 3 === 1
      ? text.replace('reads ', '')
      : text.replace('agreed', 'signed')
);

describe('a decision of many cards', () => {
  test('Accept All of every card is one change, and every card is accepted', () => {
    const doc = suggested();
    const decision = every(doc, 'accepted');
    expect(decision.cards).toHaveLength(CLAUSES);
    const outcome = apply(doc, 'olga:1', decision, 'Olga');
    expect(outcome.suggestions).toHaveLength(CLAUSES + CLAUSES / 3);
    expect(pending(doc)).toEqual([]);
    expect(texts(doc)).toEqual(ACCEPTED);
  });

  test('Reject All of every card is one change, and the original text is back', () => {
    const doc = suggested();
    apply(doc, 'olga:1', every(doc, 'rejected'), 'Olga');
    expect(pending(doc)).toEqual([]);
    expect(texts(doc)).toEqual(ORIGINAL);
  });

  for (const status of ['accepted', 'rejected'] as const)
    test(`a card that changed after the organiser saw it stays pending; the others are ${status}`, () => {
      const doc = suggested();
      const decision = every(doc, status);
      // Ana types on in her new word of clause 3, and takes back the deletion of her replacement
      // in clause 5: that card holds one suggestion now
      apply(doc, 'ana:900', { kind: 'insert', ...EDIT, place: at(doc, 3, 3), text: 'er' });
      apply(doc, 'ana:901', { kind: 'retract', suggestions: ['S:ana:105.1'] });
      expect(orderedDecidedCards(doc, { v: 1, id: 'olga:1', ...decision })).toHaveLength(
        CLAUSES - 2
      );
      apply(doc, 'olga:1', decision, 'Olga');
      expect(pending(doc).sort()).toEqual(['S:ana:103.1', 'S:ana:205.1']);
      const expected = (status === 'accepted' ? ACCEPTED : ORIGINAL).slice();
      expected[3] = `Newer ${ORIGINAL[3]}`;
      expected[5] = `${ORIGINAL[5]!.slice(0, -1)}signed.`;
      expect(texts(doc)).toEqual(expected);
    });

  test('a card another person typed inside is not decided', () => {
    const doc = suggested();
    const decision = every(doc, 'accepted');
    // Ben types inside Ana's new word: his insertion joins her card
    apply(doc, 'ben:900', { kind: 'insert', ...EDIT, place: at(doc, 0, 2), text: 'x' }, 'Ben');
    apply(doc, 'olga:1', decision, 'Olga');
    expect(pending(doc).sort()).toEqual(['S:ana:100.1', 'S:ben:900.1']);
  });

  test('when no card is as the organiser saw it, the change is refused by name', () => {
    const doc = suggested();
    const two = {
      kind: 'decide',
      cards: orderedSeenCards(doc, ['S:ana:100.1', 'S:ben:101.1']),
      status: 'accepted',
    } as const;
    apply(doc, 'ana:900', { kind: 'insert', ...EDIT, place: at(doc, 0, 2), text: 'x' });
    apply(doc, 'ben:900', { kind: 'insert', ...EDIT, place: at(doc, 1, 0), text: 'Old ' }, 'Ben');
    // Ben's new words join no card of his deletion, so that card is as it was: decided alone
    expect(orderedDecidedCards(doc, { v: 1, id: 'olga:1', ...two })).toHaveLength(1);
    apply(doc, 'ben:901', { kind: 'retract', suggestions: ['S:ben:101.1'] }, 'Ben');
    expect(orderedDecidedCards(doc, { v: 1, id: 'olga:1', ...two })).toEqual([]);
    expect(refusal(doc, 'olga:1', two, 'Olga')).toBe('stale-decision');
    // A decision whose first card is decided already says so
    const decided = {
      kind: 'decide',
      cards: orderedSeenCards(doc, ['S:ana:102.1']),
      status: 'accepted',
    } as const;
    apply(doc, 'olga:2', decided, 'Olga');
    expect(refusal(doc, 'olga:3', decided, 'Olga')).toBe('not-pending');
  });

  test('a suggestion no copy holds refuses the whole change', () => {
    const doc = suggested();
    const decision = every(doc, 'accepted');
    const forged = {
      ...decision,
      cards: [...decision.cards, { suggestions: [{ suggestion: 'S:none:1.1', version: 0 }] }],
    };
    expect(refusal(doc, 'olga:1', forged, 'Olga')).toBe('unknown-suggestion');
  });

  test('the undo record of the change gives every card back, as it was', () => {
    const doc = suggested();
    const before = JSON.stringify(snapshot(doc, false));
    const versions = pending(doc).map((id) => doc.suggestions.get(id)?.version);
    const change: OrderedDecideChange = { v: 1, id: 'olga:1', ...every(doc, 'rejected') };
    const { record } = withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson('Olga'),
        date: DATE,
        ids: orderedIdMinter('olga:1'),
      })
    );
    expect(pending(doc)).toEqual([]);
    takeBack(doc, record);
    expect(JSON.stringify(snapshot(doc, false))).toBe(before);
    expect(pending(doc).map((id) => doc.suggestions.get(id)?.version)).toEqual(versions);
    expect(texts(doc)).toEqual(texts(suggested()));
  });

  test('the room orders an Accept All as one change, and leaves the card that changed', () => {
    let tick = 0;
    const room = new OrderedRoom({
      document: suggested(),
      applyChange: realApply,
      policy: suggestionsOnlyPolicy(),
      now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
    });
    const ana = actor('Ana', 'ana-1', 'party');
    const olga = actor('Olga', 'olga-1', 'organiser');
    const doc = room.document;
    const decision = every(doc, 'accepted');
    room.submit(ana, [
      { v: 1, id: 'ana-1:1', kind: 'insert', ...EDIT, place: at(doc, 3, 3), text: 'er' },
    ]);
    const order = room.order;
    const result = room.submit(olga, [{ v: 1, id: 'olga-1:1', ...decision }]);
    expect(result.refused).toEqual([]);
    expect(result.accepted).toHaveLength(1);
    expect(room.order).toBe(order + 1);
    // Ana's words joined her own new word: its card changed
    expect(pending(doc)).toEqual(['S:ana:103.1']);
  });

  // Backlog row 56's live run (s2): a run with nothing in it but its properties, whose file
  // change of them shares its revision with a run of words. A decision reached the words and
  // left the empty run's `w:rPrChange` in the file, so a card stayed on every screen after
  // Reject All
  for (const status of ['accepted', 'rejected'] as const)
    test(`a decision reaches the format change of a run with no text (${status})`, () => {
      const change = (inner: string) =>
        `<w:rPrChange w:id="0" w:author="Author" w:date="2026-07-07T20:18:00Z"><w:rPr>${inner}</w:rPr></w:rPrChange>`;
      const file = docx(
        p(`<w:r><w:rPr><w:noProof/>${change('<w:sz w:val="22"/>')}</w:rPr></w:r>` + t('Words.')) +
          p(`<w:r><w:rPr><w:b/>${change('')}</w:rPr><w:t>Bold</w:t></w:r>`)
      );
      const doc = loadOrderedPackage(packageOf(file));
      expect(pending(doc)).toHaveLength(1);
      apply(doc, 'olga:1', every(doc, status), 'Olga');
      const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
      expect(xml).not.toContain('rPrChange');
      expect(xml).toContain(
        status === 'accepted'
          ? '<w:r><w:rPr><w:noProof/></w:rPr></w:r>'
          : '<w:r><w:rPr><w:sz w:val="22"/></w:rPr></w:r>'
      );
    });
});
