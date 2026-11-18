// Tracked changes to style definitions in a room, with real editors. Word and
// the editor alone show each change record of the styles part as a review card and decide it
// with Accept, Reject, Accept All and Reject All. In a room each is a card of its own: the
// organiser decides it as any other card, Accept All and Reject All take it in the same one
// change, every editor gets the styles part the decision leaves and lays out again, and the save
// holds it. A style card's record may have the `w:id`, author and date of a body card (s2): its
// decision never decides that body card.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPackage } from '../../compose/package.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { ACCEPTED_TEXTS, REJECTED_TEXTS, STYLED } from '../style-revision-fixture.ts';
import {
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

afterEach(closeAll);

const review = reviewTestModule();
const STYLES = '/word/styles.xml';

/** The review cards an editor shows, by key. */
const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

const styleCards = (person: Person): string[] =>
  cards(person).filter((key) => key.includes(STYLES));
const bodyCards = (person: Person): string[] =>
  cards(person).filter((key) => !key.includes(STYLES));

/** The key of the style card of one change record: its element and its `w:id`. */
const styleCard = (person: Person, element: string, id: number): string => {
  const key = styleCards(person).find((each) =>
    each.includes(`-${element}-${STYLES}\u0000${id}\u0000`)
  );
  if (!key) throw new Error(`no ${element} card`);
  return key;
};

const stylesXml = (person: Person): string =>
  serializeOoxmlPart(surface(person).session.currentPackage().parts.get(STYLES)!);

/** The styles part the room's document saves. */
const roomStyles = (room: ReturnType<typeof localRoom>): string =>
  serializeOoxmlPart(
    composeOrderedPackage(
      room.room.document,
      openOrderedWordFile(STYLED, { file: 1 }).base
    ).parts.get(STYLES)!
  );

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

const shown = (doc: OrderedDocument): string[] =>
  orderedCards(doc).flatMap((card) => card.suggestions);

/** The bold and italic the editor shows at the start of the first paragraph (Normal). */
function face(person: Person): { bold: boolean; italic: boolean } {
  caretAt(person, 0, 1);
  surface(person).layout();
  const formatting = surface(person).formatting();
  return { bold: formatting.bold === true, italic: formatting.italic === true };
}

/** The styles part and the body after the editor alone decides everything. */
function editorAlone(action: 'accept' | 'reject') {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: STYLED, modules: [review] });
  expect(editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' }).ok).toBe(true);
  const session = editor.surface!.session;
  const styles = serializeOoxmlPart(session.currentPackage().parts.get(STYLES)!);
  editor.destroy();
  container.remove();
  return styles;
}

async function people() {
  const room = localRoom(STYLED, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { review, wordFile: STYLED });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: STYLED,
  });
  return { room, ana, olive };
}

async function seated() {
  const { room, ana, olive } = await people();
  // Two body cards and five style cards, on every screen and in the room's list
  for (const person of [ana, olive]) {
    expect(bodyCards(person)).toHaveLength(2);
    expect(styleCards(person)).toHaveLength(5);
  }
  expect(orderedCards(room.room.document)).toHaveLength(7);
  return { room, ana, olive };
}

/** Every page shows the room's document and the room's styles part. */
function sameEverywhere(room: ReturnType<typeof localRoom>, people: readonly Person[]) {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
    expect(stylesXml(person)).toBe(roomStyles(room));
  }
}

describe('tracked changes to style definitions in a room', () => {
  // The default style's record has the `w:id`, author and date of Ana's body insertion. Before,
  // the client found the card by those three values alone, and the room decided the body card
  test("a style card's accept never decides the body card with the same w:id, author and date", async () => {
    const { room, ana, olive } = await people();
    expect(olive.editor.acceptReviewItem(styleCard(olive, 'rPrChange', 1)).ok).toBe(true);
    await settle(16);
    for (const person of [ana, olive]) {
      expect(bodyCards(person)).toHaveLength(2);
      expect(editorTexts(person)).toEqual(['Clause 0 new reads.', 'Quoted old text.']);
    }
    expect(shown(room.room.document)).toHaveLength(7 - 1);
  });

  for (const action of ['accept', 'reject'] as const)
    test(`the organiser's ${action} of one style card decides that card alone`, async () => {
      const { room, ana, olive } = await seated();
      expect(face(ana)).toEqual({ bold: true, italic: false });
      const key = styleCard(olive, 'rPrChange', 1);
      const done =
        action === 'accept'
          ? olive.editor.acceptReviewItem(key)
          : olive.editor.rejectReviewItem(key);
      expect(done.ok).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      sameEverywhere(room, [ana, olive]);
      for (const person of [ana, olive]) {
        // The body card with the same `w:id`, author and date stays pending
        expect(bodyCards(person)).toHaveLength(2);
        expect(editorTexts(person)).toEqual(['Clause 0 new reads.', 'Quoted old text.']);
        expect(styleCards(person)).toHaveLength(4);
        expect(stylesXml(person)).not.toContain('w:id="1"');
      }
      // The party's screen lays the default style out as the decision left it
      expect(face(ana)).toEqual(
        action === 'accept' ? { bold: true, italic: false } : { bold: false, italic: true }
      );
      // A decision has no undo (design D15)
      expect(olive.handle.session.undo()).toBe(false);
      expect(olive.refusals).toEqual([
        expect.objectContaining({ reason: 'no-inverse', detail: 'decided' }),
      ]);
    });

  for (const [action, texts] of [
    ['accept', ACCEPTED_TEXTS],
    ['reject', REJECTED_TEXTS],
  ] as const)
    test(`the editor's own ${action} all decides the style and body cards in one change`, async () => {
      const { room, ana, olive } = await seated();
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      sameEverywhere(room, [ana, olive]);
      const alone = editorAlone(action);
      for (const person of [ana, olive]) {
        expect(cards(person)).toEqual([]);
        expect(editorTexts(person)).toEqual(texts);
        expect(stylesXml(person)).not.toContain('PrChange');
        // The same styles part as the editor alone
        expect(stylesXml(person)).toBe(alone);
        const saved = openOrderedWordFile(new Uint8Array(await person.editor.save()), { file: 2 });
        expect(shown(saved.document)).toEqual([]);
        expect(serializeOoxmlPart(saved.base.parts.get(STYLES)!)).toBe(alone);
      }
      expect(face(ana)).toEqual(
        action === 'accept' ? { bold: true, italic: false } : { bold: false, italic: true }
      );
      expect(stylesXml(ana)).toContain(action === 'accept' ? 'w:after="300"' : 'w:after="100"');
    });

  for (const status of ['accepted', 'rejected'] as const)
    test(`the host's request for every card is one change (${status}); a late joiner sees it`, async () => {
      const { room, ana, olive } = await seated();
      const every = shown(olive.handle.session.orderedDocument());
      expect(every).toHaveLength(7);
      expect(olive.handle.session.requestDecisions(every, status)).not.toBeNull();
      await settle(16);
      expect(kinds(room)).toEqual(['decide']);
      const late = await join(room, 'Bo', { review, wordFile: STYLED });
      sameEverywhere(room, [ana, olive, late]);
      for (const person of [ana, olive, late]) {
        expect(cards(person)).toEqual([]);
        expect(stylesXml(person)).toBe(editorAlone(status === 'accepted' ? 'accept' : 'reject'));
      }
    });

  test('a party may not decide a style card', async () => {
    const { room, ana } = await seated();
    const done = ana.editor.acceptReviewItem(styleCard(ana, 'pPrChange', 7));
    await settle(16);
    expect(done.ok).toBe(true);
    expect(kinds(room)).toEqual([]);
    expect(ana.refusals).toEqual([expect.objectContaining({ reason: 'policy-decide' })]);
    expect(styleCards(ana)).toHaveLength(5);
    expect(stylesXml(ana)).toBe(roomStyles(room));
    expect(roomStyles(room)).toContain('<w:pPrChange ');
  });
});
