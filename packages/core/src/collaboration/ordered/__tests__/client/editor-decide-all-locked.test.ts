// Accept All and Reject All next to a content control's lock, with real editors and one ordered
// room. Word leaves the changes a lock keeps and decides the others (desktop Word for Mac,
// 2026-11-26: a deletion in a group and one outside it; Accept All accepted the one outside and
// left the one in the group). The organiser's Accept All does the same in the room: the card in
// the group stays, the other is decided, and every copy and a reopened file agree. The engine
// once refused the whole Accept All.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();
const del = (id: number, text: string) =>
  `<w:del w:id="${id}" w:author="Bea" w:date="2026-11-26T00:00:00Z"><w:r>` +
  `<w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

/** Bea's deletion of "ab" in a Word-saved group, and of "." after "Signed". */
const FILE = docx(
  p(
    t('A: ') +
      '<w:sdt><w:sdtPr><w:tag w:val="group"/><w:lock w:val="contentLocked"/><w:group/>' +
      `</w:sdtPr><w:sdtContent>${t('G ')}${del(6, 'ab')}</w:sdtContent></w:sdt>`
  ) + p(t('Signed') + del(7, '.'))
);

async function seated() {
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: FILE,
  });
  const bob = await join(room, 'Bob', { review, wordFile: FILE });
  return { room, olive, bob };
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

/** The words a document's main part holds as deleted, pending a decision. */
const deletedWords = (doc: OrderedDocument): string[] =>
  [
    ...serializeOoxmlPart(composeOrderedPart(doc, BODY).part).matchAll(
      /<w:delText(?: [^>]*)?>([^<]*)<\/w:delText>/g
    ),
  ].map((match) => match[1]!);

describe('Accept All and Reject All leave a change a lock keeps (Word)', () => {
  for (const action of ['accept', 'reject'] as const) {
    test(`the organiser's ${action} all decides the change outside the group only`, async () => {
      const { room, olive, bob } = await seated();
      expect(orderedCards(room.room.document)).toHaveLength(2);
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle(16);
      expect(olive.refusals).toEqual([]);
      converged(room, olive, bob);
      expect(orderedCards(room.room.document)).toHaveLength(1);
      expect(editorTexts(bob)[1]).toBe(action === 'accept' ? 'Signed' : 'Signed.');

      expect(deletedWords(room.room.document)).toEqual(['ab']);
      const saved = new Uint8Array(await olive.editor.save());
      expect(deletedWords(openOrderedWordFile(saved, { file: 2 }).document)).toEqual(['ab']);
    });
  }

  test('the room leaves the card a lock keeps, whatever a client asks', async () => {
    const { room, olive, bob } = await seated();
    const all = orderedCards(room.room.document).flatMap((card) => card.suggestions);
    expect(olive.handle.session.requestDecisions(all, 'accepted')).not.toBeNull();
    await settle(16);
    converged(room, olive, bob);
    expect(orderedCards(room.room.document)).toHaveLength(1);
    expect(deletedWords(room.room.document)).toEqual(['ab']);
  });
});
