// A file's format change on a run with no text, with a revision of its own (found
// in a live run on s2). The list keeps such a run whole as one item with no width; the
// cards read only units with a width, so the change made no card: the host's Accept all and
// Reject all did not name it, and it stayed pending in the save. The editor alone and Word
// decide it. It is a card of its own now, decided with every other card in the one change.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
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
const DATE = '2026-10-01T00:00:00Z';

// Paragraph 0: an empty run whose size Ana changed (its own `w:id` 5), then words. Paragraph 1:
// Ana's new word, a card of its own
const FILE = docx(
  p(
    '<w:r><w:rPr><w:noProof/>' +
      `<w:rPrChange w:id="5" w:author="Ana" w:date="${DATE}"><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrChange>` +
      '</w:rPr></w:r>' +
      t('Words.')
  ) +
    p(
      t('Clause ') +
        `<w:ins w:id="6" w:author="Ana" w:date="${DATE}">${t('new ')}</w:ins>` +
        t('end.')
    )
);

const shown = (doc: OrderedDocument): string[] =>
  orderedCards(doc).flatMap((card) => card.suggestions);

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

/** The saved main part of an editor, reopened. */
async function savedBody(person: Person) {
  const saved = openOrderedWordFile(new Uint8Array(await person.editor.save()), { file: 2 });
  return {
    pending: shown(saved.document),
    xml: serializeOoxmlPart(saved.base.parts.get('/word/document.xml')!),
  };
}

describe('a format change on a run with no text in a room', () => {
  test('is a card of its own', async () => {
    const room = localRoom(FILE);
    expect(orderedCards(room.room.document)).toHaveLength(2);
  });

  for (const status of ['accepted', 'rejected'] as const)
    test(`the host's request for every card decides it (${status})`, async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review, wordFile: FILE });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        review,
        wordFile: FILE,
      });
      const every = shown(olive.handle.session.orderedDocument());
      expect(olive.handle.session.requestDecisions(every, status)).not.toBeNull();
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      for (const person of [ana, olive]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        expect(editorTexts(person)).toEqual([
          'Words.',
          status === 'accepted' ? 'Clause new end.' : 'Clause end.',
        ]);
        const saved = await savedBody(person);
        expect(saved.pending).toEqual([]);
        expect(saved.xml).not.toContain('rPrChange');
        expect(saved.xml).toContain(
          status === 'accepted'
            ? '<w:r><w:rPr><w:noProof/></w:rPr></w:r>'
            : '<w:r><w:rPr><w:sz w:val="22"/></w:rPr></w:r>'
        );
      }
    });

  for (const action of ['accept', 'reject'] as const)
    test(`the editor's own ${action} all decides it in the same change`, async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review, wordFile: FILE });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        review,
        wordFile: FILE,
      });
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      for (const person of [ana, olive]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(person.editor.getReviewItems()).toEqual([]);
        const saved = await savedBody(person);
        expect(saved.pending).toEqual([]);
        expect(saved.xml).not.toContain('rPrChange');
      }
    });
});
