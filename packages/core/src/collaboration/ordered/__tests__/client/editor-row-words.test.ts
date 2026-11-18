// A party's new row and the words they type in it later, with real editors and one room (the
// live run of the 2.x flows, T step t15). The words are a suggestion of their own, made in a
// later batch, but they sit in the party's own new row: the row's card holds them (join 4), so
// one decision decides both, as the editor alone shows them as the row's one card and as 2.x
// decided a row with what it holds. Before, the host's accept of the row's card left the words
// a suggestion with a card of their own, in the saved file too.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { orderedSuggestionsOfRevision } from '../../client/ops/review.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import type { ReviewRevisionItem } from '../../../../store/store/review-items.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { caretIn, TABLE_FILE, tableCommand } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** The editor's revision items. */
const items = (person: Person): ReviewRevisionItem[] =>
  surface(person)
    .session.reviewItems()
    .filter((item): item is ReviewRevisionItem => item.kind === 'revision');

/** The suggestions a host decides for the editor's one card, by its lead mark, as a host does. */
function suggestionsOfCard(person: Person): readonly string[] {
  const [card] = items(person);
  if (!card) throw new Error('no card');
  const { id, author, date } = card.address;
  return orderedSuggestionsOfRevision(person.handle.session.orderedDocument(), {
    part: BODY,
    id,
    author,
    ...(date === undefined ? {} : { date }),
  });
}

/** Ana's new row below "Fee", with "ROWA" typed in its last cell in a later batch. */
async function anaRowWithWords() {
  const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
  const olive = await join(room, 'Olive', {
    wordFile: TABLE_FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  caretIn(ana, 'Fee');
  expect(tableCommand(ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
  await settle();
  caretAt(ana, editorTexts(ana).indexOf('Total') + 3, 0);
  surface(ana).insertPlainText('ROWA');
  await settle();
  return { room, ana, olive };
}

describe("a party's new row holds the words they type in it", () => {
  for (const [status, rows] of [
    ['accepted', ['Fee', 'Rate', 'Total', '', '', 'ROWA', '100', '5', '105']],
    ['rejected', ['Fee', 'Rate', 'Total', '100', '5', '105']],
  ] as const)
    test(`one card; the host's decision of it is ${status} for the row and its words`, async () => {
      const { room, ana, olive } = await anaRowWithWords();
      const [card, ...others] = orderedCards(room.room.document);
      expect(others).toEqual([]);
      expect(card?.suggestions).toHaveLength(2);
      expect(items(olive)).toHaveLength(1);
      expect(
        olive.handle.session.requestDecisions(suggestionsOfCard(olive), status)
      ).not.toBeNull();
      await settle();
      for (const person of [ana, olive]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        expect(editorTexts(person).slice(1, -1)).toEqual([...rows]);
        expect(items(person)).toEqual([]);
      }
      expect(orderedCards(room.room.document)).toEqual([]);
      // The saved file holds no change: the row and its words were decided as one
      const saved = strFromU8(
        unzipSync(new Uint8Array(await olive.editor.save()))['word/document.xml']!
      );
      expect(saved).not.toMatch(/<w:(ins|del) /);
      expect(saved.includes('ROWA')).toBe(status === 'accepted');
    });
});
