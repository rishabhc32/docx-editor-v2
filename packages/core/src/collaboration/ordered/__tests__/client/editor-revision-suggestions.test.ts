// A host's card to its suggestions (`orderedSuggestionsOfRevision`): a dashboard reads the
// editor's review items and asks which pending suggestions each mark names, with the index the
// editor's own decisions use. Two suggestions of one author in one second in one paragraph are
// told apart by their `w:id`; a decided one names none; a style's record and a body suggestion
// with the same `w:id`, author and date are told apart by their part.

import { afterEach, expect, test } from 'bun:test';
import type { ReviewRevisionItem } from '../../../../store/store/review-items.ts';
import { orderedSuggestionsOfRevision } from '../../client/ops/review.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { STYLED } from '../style-revision-fixture.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();
const BODY = '/word/document.xml';
const STYLES = '/word/styles.xml';

/** The editor's revision items, as a host reads them. */
const items = (person: Person): ReviewRevisionItem[] =>
  surface(person)
    .session.reviewItems()
    .filter((item): item is ReviewRevisionItem => item.kind === 'revision');

/** The suggestions one item's mark names, as a host asks. */
function suggestionsOf(doc: OrderedDocument, item: ReviewRevisionItem): readonly string[] {
  const { id, author, date } = item.address;
  const part = item.ranges[0]?.partName ?? STYLES;
  return orderedSuggestionsOfRevision(doc, {
    part,
    id,
    author,
    ...(date === undefined ? {} : { date }),
    ...(item.formattingKind === undefined ? {} : { element: item.formattingKind }),
  });
}

test('two suggestions of one author in one second in one paragraph are told apart', async () => {
  const bytes = docx(p(t('Alpha beta gamma.')));
  // One clock second for every batch: both suggestions have one date
  const room = localRoom(bytes, {
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 28, 9),
  });
  const ana = await join(room, 'Ana', { review, wordFile: bytes });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: bytes,
  });
  caretAt(ana, 0, 6);
  surface(ana).insertPlainText('new ');
  ana.handle.session.flushPendingJournals();
  await settle(16);
  caretAt(ana, 0, 0);
  surface(ana).insertPlainText('So ');
  ana.handle.session.flushPendingJournals();
  await settle(16);
  const doc = olive.handle.session.orderedDocument();
  const cards = orderedCards(doc);
  expect(cards).toHaveLength(2);
  const found = items(olive);
  expect(found).toHaveLength(2);
  expect(new Set(found.map((item) => item.date)).size).toBe(1);
  // Each item names exactly its own suggestion
  const named = found.map((item) => suggestionsOf(doc, item));
  expect(named.map((ids) => ids.length)).toEqual([1, 1]);
  expect(new Set(named.flat())).toEqual(new Set(cards.map((card) => card.id)));
  // A decided suggestion names none
  const first = found[0]!;
  const decided = named[0]![0]!;
  expect(olive.handle.session.requestDecisions([decided], 'accepted')).not.toBeNull();
  await settle(16);
  expect(suggestionsOf(olive.handle.session.orderedDocument(), first)).toEqual([]);
  expect(suggestionsOf(olive.handle.session.orderedDocument(), found[1]!)).toEqual(named[1]);
});

test("a style's record and a body suggestion with the same three values are told apart by part", async () => {
  const room = localRoom(STYLED);
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: STYLED,
  });
  const doc = olive.handle.session.orderedDocument();
  const same = { id: '1', author: 'Ana', date: '2026-10-01T00:00:00Z' };
  const body = orderedSuggestionsOfRevision(doc, { part: BODY, ...same });
  const style = orderedSuggestionsOfRevision(doc, { part: STYLES, ...same, element: 'rPrChange' });
  expect(body).toHaveLength(1);
  expect(style).toHaveLength(1);
  expect(style[0]).not.toBe(body[0]);
  expect(doc.partRevisions.get(style[0]!)?.styleId).toBe('Normal');
  // Outside the stories the element is part of the name
  expect(orderedSuggestionsOfRevision(doc, { part: STYLES, ...same })).toEqual([]);
  expect(
    orderedSuggestionsOfRevision(doc, { part: STYLES, ...same, element: 'pPrChange' })
  ).toEqual([]);
  // Every item the editor shows names exactly one pending suggestion, and all differ
  const named = items(olive).map((item) => suggestionsOf(doc, item));
  expect(named.every((ids) => ids.length === 1)).toBe(true);
  expect(new Set(named.flat()).size).toBe(7);
});
