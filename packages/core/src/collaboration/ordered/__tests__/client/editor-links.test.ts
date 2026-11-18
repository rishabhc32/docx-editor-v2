// Links in an ordered room (group 9.5), with real editors and one room: a new link, a new
// target, and an unlink, made by the organiser (in place, tracking off) and by a party (the
// editor's proposal: the words struck and a copy inserted). Each saves as the editor alone saves
// the same edit; the organiser's undo and redo and a party's take-back are exact; a rejected or
// taken-back link leaves no address in the file; the party rules hold; and a saved file
// reopens with its suggestions still to decide.

import { afterEach, describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../../client/open.ts';
import { decideOrdered, orderedCards } from '../../decide.ts';
import { withUndoRecord } from '../../model/undo-record.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
} from './editor-harness.ts';
import {
  act,
  alone,
  converged,
  linkAddresses,
  roomFile,
  sameAsAlone,
  select,
  storyAsRead,
  styleAsRead,
  type Act,
} from './links-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

/** Styles with a default character style and no Hyperlink style. */
const PLAIN_STYLES =
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">' +
  '<w:name w:val="Default Paragraph Font"/></w:style></w:styles>';
const NO_LINK_STYLE = docx(
  p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')) + p(t('Signed here.')),
  { parts: { 'word/styles.xml': PLAIN_STYLES } }
);

const link = (href: string, extra: { tooltip?: string; text?: string } = {}) =>
  ({ type: 'insertHyperlink', href, ...extra }) as const;

describe('links made by the organiser (tracking off)', () => {
  test('a new link, a new address and tooltip, and an unlink save as the editor alone does', async () => {
    await sameAsAlone(FILE, 'edit', [
      (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/'))),
    ]);
    await sameAsAlone(FILE, 'edit', [
      (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/'))),
      (e) => (select(e, 0, 5), e.exec(link('https://b.example/', { tooltip: 'Terms' }))),
    ]);
    await sameAsAlone(FILE, 'edit', [
      (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/'))),
      (e) => (select(e, 0, 5), e.exec({ type: 'removeHyperlink' })),
    ]);
  });

  test('a link at a caret types its text; a link to a bookmark has no address', async () => {
    const { room } = await sameAsAlone(FILE, 'edit', [
      (e) => (select(e, 2, 0), e.exec(link('https://c.example/', { text: 'See ' }))),
      (e) => (select(e, 1, 0, 3), e.exec(link('#_Toc1'))),
    ]);
    const saved = roomFile(room.room.document, FILE);
    expect(storyAsRead(saved)).toContain('w:anchor="_Toc1"');
    expect(linkAddresses(saved)).toEqual(['https://c.example/', 'https://example.com/old']);
  });

  test("a first link adds Word's Hyperlink style, as the editor alone does", async () => {
    const step: Act = (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/')));
    const { room } = await sameAsAlone(NO_LINK_STYLE, 'edit', [step]);
    const editorAlone = await alone(NO_LINK_STYLE, 'Ana', 'edit', step);
    const saved = roomFile(room.room.document, NO_LINK_STYLE);
    expect(styleAsRead(saved, 'Hyperlink')).not.toBeNull();
    expect(styleAsRead(saved, 'Hyperlink')).toBe(styleAsRead(editorAlone, 'Hyperlink'));
  });

  test("the organiser's undo and redo are exact: the address goes and comes back", async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana');
    const original = storyAsRead(roomFile(room.room.document, FILE));
    await act(olive, (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/'))));
    const linked = storyAsRead(roomFile(room.room.document, FILE));
    await act(olive, (e) => (select(e, 0, 5), e.exec({ type: 'removeHyperlink' })));
    await act(olive, (e) => e.surface?.undo());
    expect(storyAsRead(roomFile(room.room.document, FILE))).toBe(linked);
    await act(olive, (e) => e.surface?.undo());
    const undone = roomFile(room.room.document, FILE);
    expect(storyAsRead(undone)).toBe(original);
    expect(linkAddresses(undone)).toEqual(['https://example.com/old']);
    await act(olive, (e) => e.surface?.redo());
    expect(storyAsRead(roomFile(room.room.document, FILE))).toBe(linked);
    converged(room, [olive, ana]);
  });
});

describe("links a party proposes (the editor's proposal: struck words and a copy)", () => {
  const review = reviewTestModule();

  test('a new link, a new target and an unlink save as the editor alone proposes them', async () => {
    await sameAsAlone(FILE, 'suggesting', [
      (e) => (select(e, 1, 0, 3), e.exec(link('https://c.example/'))),
    ]);
    // A base link: the organiser's file link to change and to remove
    const linked = docx(
      p(
        t('The ') +
          '<w:hyperlink r:id="rIdLink">' +
          t('fee', '<w:rStyle w:val="Hyperlink"/>') +
          '</w:hyperlink>' +
          t(' is 100 USD.')
      ) + p(t('Due in 30 days.'))
    );
    await sameAsAlone(linked, 'suggesting', [
      (e) => (select(e, 0, 5), e.exec(link('https://d.example/'))),
    ]);
    await sameAsAlone(linked, 'suggesting', [
      (e) => (select(e, 0, 5), e.exec({ type: 'removeHyperlink' })),
    ]);
  });

  test('accepted, the link stays with its address; rejected, the words and no address', async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { review });
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
      await act(ana, (e) => (select(e, 1, 0, 3), e.exec(link('https://c.example/'))));
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      const decided =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(decided.ok).toBe(true);
      await settle();
      converged(room, [ana, olive]);
      const saved = roomFile(room.room.document, FILE);
      expect(storyAsRead(saved)).not.toContain('<w:ins');
      expect(storyAsRead(saved)).not.toContain('<w:del');
      expect(linkAddresses(saved)).toEqual(
        decision === 'accept'
          ? ['https://c.example/', 'https://example.com/old']
          : ['https://example.com/old']
      );
      expect(storyAsRead(saved).includes('<w:hyperlink')).toBe(decision === 'accept');
    }
  });

  test("a party's undo takes the proposal back whole, and redo makes it again", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana');
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const original = storyAsRead(roomFile(room.room.document, FILE));
    await act(ana, (e) => (select(e, 1, 0, 3), e.exec(link('https://c.example/'))));
    const proposed = storyAsRead(roomFile(room.room.document, FILE));
    await act(ana, (e) => e.surface?.undo());
    const back = roomFile(room.room.document, FILE);
    expect(storyAsRead(back)).toBe(original);
    expect(linkAddresses(back)).toEqual(['https://example.com/old']);
    await act(ana, (e) => e.surface?.redo());
    expect(storyAsRead(roomFile(room.room.document, FILE))).toBe(proposed);
    converged(room, [ana, olive]);
  });

  test('a saved proposal reopens, and its card is still to decide', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana');
    await act(ana, (e) => (select(e, 1, 0, 3), e.exec(link('https://c.example/'))));
    const { writeOoxmlPackage } = await import('../../../../store/package/ooxml-package.ts');
    const bytes = writeOoxmlPackage(roomFile(room.room.document, FILE));
    for (const decision of ['accept', 'reject'] as const) {
      const again = openOrderedWordFile(bytes, { file: 2 }).document;
      // Reopened, the struck words and the linked copy are two cards (the link's edge is between
      // them); the editor's review pairs them, and its decision decides both
      const cards = orderedCards(again);
      expect(cards.length).toBeGreaterThan(0);
      withUndoRecord(again, () => {
        for (const card of cards) decideOrdered(again, card.suggestions[0] ?? '', decision);
      });
      const base = openOrderedWordFile(bytes, { file: 2 }).base;
      const { composeOrderedPackage } = await import('../../compose/package.ts');
      const decided = composeOrderedPackage(again, base);
      expect(storyAsRead(decided).includes('<w:hyperlink')).toBe(decision === 'accept');
      // The reopened file's relationship is the file's now: rejected, nothing uses it (the
      // editor alone keeps an unused link relationship too)
      expect(storyAsRead(decided).includes('r:id="https://c.example/"')).toBe(
        decision === 'accept'
      );
    }
  });

  test("the party rules: a party's untracked link is refused as a package change", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { mode: 'edit' });
    await act(ana, (e) => (select(e, 0, 4, 7), e.exec(link('https://a.example/'))));
    // The room's refusal names the party rule's reason: not a suggestion, a package change
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toContain('ordered-refused:package');
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(storyAsRead(roomFile(room.room.document, FILE))).not.toContain('<w:hyperlink');
  });
});

describe('links while several people edit', () => {
  test('the organiser removes a link while a party types in its words', async () => {
    const linked = docx(
      p(
        t('The ') +
          '<w:hyperlink r:id="rIdLink">' +
          t('fee', '<w:rStyle w:val="Hyperlink"/>') +
          '</w:hyperlink>' +
          t(' is 100 USD.')
      )
    );
    const room = localRoom(linked, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: linked });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: linked });
    // Both edit before either hears of the other
    select(ana.editor, 0, 5);
    ana.editor.surface?.insertPlainText('X');
    select(olive.editor, 0, 5);
    olive.editor.exec({ type: 'removeHyperlink' });
    await settle();
    converged(room, [ana, olive]);
    const story = storyAsRead(roomFile(room.room.document, linked));
    expect(story).not.toContain('<w:hyperlink');
    expect(story).toContain('X');
  });

  test('two people link the same words: one link, the proposal beside it', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana');
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    select(ana.editor, 1, 0, 3);
    ana.editor.exec(link('https://c.example/'));
    select(olive.editor, 1, 0, 3);
    olive.editor.exec(link('https://o.example/'));
    await settle();
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
    }
  });

  test("a party's new address copies the link's words without text a decision took out", async () => {
    const linked = docx(
      p(
        t('The ') +
          '<w:hyperlink r:id="rIdLink">' +
          t('fee', '<w:rStyle w:val="Hyperlink"/>') +
          '</w:hyperlink>' +
          t(' is 100 USD.')
      )
    );
    const review = reviewTestModule();
    const room = localRoom(linked, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: linked, review });
    const bob = await join(room, 'Bob', { wordFile: linked, review });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: linked,
      review,
    });
    // Bob types in the file's link, and the organiser rejects it
    await act(bob, (e) => (select(e, 0, 5), e.surface?.insertPlainText('X')));
    const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    if (!card) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card.key).ok).toBe(true);
    await settle();
    // Ana proposes a new address: the words struck, and a copy of the words that show
    await act(ana, (e) => (select(e, 0, 5), e.exec(link('https://d.example/'))));
    converged(room, [ana, bob, olive]);
    const story = storyAsRead(roomFile(room.room.document, linked));
    expect(story).not.toContain('X');
    expect(story).toMatch(
      /<w:hyperlink r:id="https:\/\/d\.example\/"[^>]*>(?:(?!<\/w:hyperlink>).)*>fee</
    );
  });
});
