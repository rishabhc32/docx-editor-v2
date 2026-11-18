// Accept All and Reject All of many cards in a room, with real editors: one
// request from the page, one change in the room's order and log, one store change on each
// editor, and no card left. The organiser asks through the host's request (`requestDecisions`,
// the host's cards) or with the editor's own Accept All and Reject All. A card that changed after
// the organiser's copy showed it stays pending; the others are decided.

import { afterEach, describe, expect, test } from 'bun:test';
import { liveLayoutPassTestRecorder } from '../../../../editor/surface-layout-passes.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import type { TreeModelChange } from '../../../../store/store/tree-store.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
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

afterEach(closeAll);

const review = reviewTestModule();
const passes = liveLayoutPassTestRecorder();

/** Cards in the file: Word's own tracked changes by Ana, one in each clause. */
const CARDS = 320;
const DATE = '2026-10-01T00:00:00Z';
const ins = (k: number, text: string) =>
  `<w:ins w:id="${k + 1}" w:author="Ana" w:date="${DATE}">${t(text)}</w:ins>`;
const del = (k: number, text: string) =>
  `<w:del w:id="${k + 1}" w:author="Ana" w:date="${DATE}"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

// Each even clause has a new word, each odd one a deleted word
const MANY = docx(
  Array.from({ length: CARDS }, (_, k) =>
    p(t(`Clause ${k} `) + (k % 2 === 0 ? ins(k, 'new ') : del(k, 'old ')) + t('reads as agreed.'))
  ).join('')
);
const ACCEPTED = Array.from({ length: CARDS }, (_, k) =>
  k % 2 === 0 ? `Clause ${k} new reads as agreed.` : `Clause ${k} reads as agreed.`
);
const REJECTED = Array.from({ length: CARDS }, (_, k) =>
  k % 2 === 0 ? `Clause ${k} reads as agreed.` : `Clause ${k} old reads as agreed.`
);

/** The keys of the revision cards an editor shows. */
const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

/** Every room change's kind, in order. */
const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

/** The suggestions of every card a copy shows. */
const shown = (doc: OrderedDocument): string[] =>
  orderedCards(doc).flatMap((card) => card.suggestions);

/** The store changes each editor publishes while `run` runs, every queued task included. */
async function changesOf(people: readonly Person[], run: () => void): Promise<number[]> {
  const counts = people.map(() => 0);
  const stops = people.map((person, k) =>
    surface(person).session.subscribe((_change: TreeModelChange) => {
      counts[k] = (counts[k] ?? 0) + 1;
    })
  );
  passes.reset();
  run();
  await settle(16);
  for (const stop of stops) stop();
  return counts;
}

/** A party and an organiser on the file of many cards. */
async function seated() {
  const room = localRoom(MANY, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { review, wordFile: MANY });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: MANY,
  });
  expect(cards(olive)).toHaveLength(CARDS);
  return { room, ana, olive };
}

/** Every page shows the room's document, with no card, and the save reopens with none. */
async function decidedEverywhere(
  room: ReturnType<typeof localRoom>,
  people: readonly Person[],
  texts: readonly string[]
) {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
    expect(cards(person)).toEqual([]);
    expect(editorXml(person)).not.toContain('<w:ins ');
    expect(editorXml(person)).not.toContain('<w:del ');
    expect(editorTexts(person)).toEqual([...texts]);
    const saved = new Uint8Array(await person.editor.save());
    expect(shown(openOrderedWordFile(saved, { file: 2 }).document)).toEqual([]);
  }
}

describe('Accept All and Reject All of many cards in a room', () => {
  for (const [status, texts] of [
    ['accepted', ACCEPTED],
    ['rejected', REJECTED],
  ] as const)
    test(`the host's request decides ${CARDS} cards in one change (${status})`, async () => {
      const { room, ana, olive } = await seated();
      const session = olive.handle.session;
      const every = shown(session.orderedDocument());
      expect(every).toHaveLength(CARDS);
      let id: string | null = null;
      const published = await changesOf([olive, ana], () => {
        id = session.requestDecisions(every, status);
      });
      expect(id).not.toBeNull();
      // One change in the room's order and log, and one store change on each editor
      expect(kinds(room)).toEqual(['decide']);
      expect(published).toEqual([1, 1]);
      await decidedEverywhere(room, [ana, olive], texts);
    });

  for (const status of ['accepted', 'rejected'] as const)
    test(`a card that changed before the request arrives stays; the others are ${status}`, async () => {
      const { room, ana, olive } = await seated();
      const oscar = await join(room, 'Oscar', {
        mode: 'edit',
        login: ORGANISER,
        review,
        wordFile: MANY,
      });
      const every = shown(olive.handle.session.orderedDocument());
      // A file's word is nobody's in the room: another organiser types inside the new word of
      // clause 0, which changes that card, and the room orders it before the decision
      caretAt(oscar, 0, 11);
      surface(oscar).insertPlainText('x');
      oscar.handle.session.flushPendingJournals();
      expect(olive.handle.session.requestDecisions(every, status)).not.toBeNull();
      await settle(16);
      expect(kinds(room)).toEqual(['insert', 'decide']);
      for (const person of [ana, olive, oscar]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        // The organiser's copy tells how many cards the room left: the one that changed
        expect(orderedCards(person.handle.session.orderedDocument())).toHaveLength(1);
        expect(cards(person)).toHaveLength(1);
        expect(editorTexts(person)[0]).toBe('Clause 0 nexw reads as agreed.');
        expect(editorTexts(person).slice(1)).toEqual(
          (status === 'accepted' ? ACCEPTED : REJECTED).slice(1)
        );
      }
    });

  for (const [action, texts] of [
    ['accept', ACCEPTED],
    ['reject', REJECTED],
  ] as const)
    test(`the editor's own ${action} all of ${CARDS} cards is one change and one undo step`, async () => {
      const { room, ana, olive } = await seated();
      const published = await changesOf([olive, ana], () => {
        const done = olive.editor.exec({
          type: 'resolveAllReviewChanges',
          action,
          scope: 'document',
        });
        expect(done.ok).toBe(true);
      });
      expect(kinds(room)).toEqual(['decide']);
      // The party's editor takes the decision in one store change
      expect(published[1]).toBe(1);
      await decidedEverywhere(room, [ana, olive], texts);
      // A decision has no undo: its one step is refused by name, once, and goes
      expect(olive.handle.session.undo()).toBe(false);
      expect(olive.handle.session.undo()).toBe(false);
      expect(olive.refusals).toEqual([
        expect.objectContaining({ reason: 'no-inverse', detail: 'decided' }),
      ]);
    });

  // A file whose styles part holds a change record of Ana's that records no old properties. A
  // style's change record is a card of the room (`editor-style-revisions.test.ts`),
  // but the engine cannot reject this one, so the list holds no suggestion for it: the editor's
  // own Accept All decides every other card in one change, leaves the record as it is, and says
  // so once by name; its Reject All leaves the record as the editor alone does, with no word
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const STYLED = docx(
    Array.from({ length: 8 }, (_, k) =>
      p(t(`Clause ${k} `) + (k % 2 === 0 ? ins(k, 'new ') : del(k, 'old ')) + t('reads as agreed.'))
    ).join(''),
    {
      parts: {
        'word/styles.xml':
          `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Quote">` +
          '<w:name w:val="Quote"/><w:rPr><w:i/>' +
          `<w:rPrChange w:id="900" w:author="Ana" w:date="${DATE}"/>` +
          '</w:rPr></w:style></w:styles>',
      },
    }
  );
  const STYLE_CARD = /styles\.xml/;
  const stylesXml = (person: Person) =>
    serializeOoxmlPart(surface(person).session.currentPackage().parts.get('/word/styles.xml')!);

  for (const [action, texts] of [
    ['accept', ACCEPTED.slice(0, 8)],
    ['reject', REJECTED.slice(0, 8)],
  ] as const)
    test(`the editor's own ${action} all leaves a style record the room lacks and decides the rest in one change`, async () => {
      const room = localRoom(STYLED, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review, wordFile: STYLED });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        review,
        wordFile: STYLED,
      });
      expect(cards(olive)).toHaveLength(9);
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      expect(olive.refusals).toEqual(
        action === 'accept'
          ? [expect.objectContaining({ code: 'ordered-refused', reason: 'not-in-room' })]
          : []
      );
      for (const person of [ana, olive]) {
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        expect(editorTexts(person)).toEqual([...texts]);
        // The style's card and its revision stay, on the screen and in the save
        expect(cards(person)).toHaveLength(1);
        expect(cards(person)[0]).toMatch(STYLE_CARD);
        expect(stylesXml(person)).toContain('<w:rPrChange ');
      }
    });

  test("the editor's own accept all of a style record the room lacks, alone, is refused by name", async () => {
    const STYLE_ONLY = docx(p(t('Clause 0 reads as agreed.')), {
      parts: {
        'word/styles.xml':
          `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Quote">` +
          '<w:name w:val="Quote"/><w:rPr><w:i/>' +
          `<w:rPrChange w:id="900" w:author="Ana" w:date="${DATE}"/>` +
          '</w:rPr></w:style></w:styles>',
      },
    });
    const room = localRoom(STYLE_ONLY);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      review,
      wordFile: STYLE_ONLY,
    });
    const done = olive.editor.exec({
      type: 'resolveAllReviewChanges',
      action: 'accept',
      scope: 'document',
    });
    expect(done.ok).toBe(false);
    await settle(16);
    expect(kinds(room)).toEqual([]);
    expect(olive.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'not-in-room' }),
    ]);
    expect(cards(olive)).toHaveLength(1);
    expect(stylesXml(olive)).toContain('<w:rPrChange ');
  });

  // The live run on s2: a run with nothing in it but its properties, whose file
  // change of them shares its revision with runs of words. After Reject All its card stayed on
  // both screens, and the save kept its w:rPrChange
  for (const status of ['accepted', 'rejected'] as const)
    test(`a run with no text loses its format change with the decision (${status})`, async () => {
      const change = (inner: string) =>
        `<w:rPrChange w:id="0" w:author="Author" w:date="${DATE}"><w:rPr>${inner}</w:rPr></w:rPrChange>`;
      const file = docx(
        p(`<w:r><w:rPr><w:noProof/>${change('<w:sz w:val="22"/>')}</w:rPr></w:r>` + t('Words.')) +
          p(`<w:r><w:rPr><w:b/>${change('')}</w:rPr><w:t>Bold</w:t></w:r>`) +
          p(t('Clause ') + ins(5, 'new ') + t('one.'))
      );
      const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review, wordFile: file });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        review,
        wordFile: file,
      });
      const every = shown(olive.handle.session.orderedDocument());
      expect(olive.handle.session.requestDecisions(every, status)).not.toBeNull();
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      for (const person of [ana, olive]) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
        expect(cards(person)).toEqual([]);
        expect(editorXml(person)).not.toContain('rPrChange');
      }
    });
});
