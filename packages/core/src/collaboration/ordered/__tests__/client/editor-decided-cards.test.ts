// Decisions, with real editors and one room (design D7, D15): a decision decides the whole card
// the page shows, and it is final.
//
// Lines typed with Enter between them, or struck a piece at a time, are one suggestion each in the
// room, and the page shows them as one card, as the editor alone and Word do. The page's Accept
// or Reject of that card decides every one of them in one `decide`; the card's first mark alone
// names one piece.
//
// A party types, then edits their own new words (types more into them, pastes a line into them,
// bolds or deletes a part of them). The organiser accepts or rejects the card. The party's undo
// of the last edit took the edit back in the decided words for everyone (the typed words went),
// while the decision stood. Now the undo is refused by name and nothing moves on any page, and
// the room refuses such a take-back when the decision reaches it first. So is the organiser's
// undo of words typed directly inside a party's insertion that was decided since.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedSuggestionsOfRevision } from '../../client/ops/review.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkFold, mediaStore } from '../fold-check.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

type Room = ReturnType<typeof localRoom>;
type Status = 'accepted' | 'rejected';

/** A party and an organiser in a room that holds the party rules. */
async function people() {
  const room = localRoom(undefined, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { review });
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
  return { room, ana, olive };
}

/** Every room change's kind, in order. */
const kinds = (room: Room): string[] => room.room.joinSnapshot(0).map((change) => change.kind);

const fingerprints = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].map((name) =>
    canonicalOoxmlFingerprint(composeOrderedPart(doc, name).part)
  );

/** The revision cards a page shows. */
const pageCards = (person: Person) =>
  person.editor.getReviewItems().filter((placement) => placement.kind === 'revision');

/** The organiser decides the one card the page shows, as the page's buttons do. */
async function decideCard(olive: Person, status: Status): Promise<void> {
  const cards = pageCards(olive);
  expect(cards).toHaveLength(1);
  const key = cards[0]!.key;
  const done =
    status === 'accepted' ? olive.editor.acceptReviewItem(key) : olive.editor.rejectReviewItem(key);
  expect(done.ok).toBe(true);
  await settle();
}

/**
 * The refusal each page reported, taken off; then every page shows the room's document, the
 * saved file reopens as it, and a fold keeps it.
 */
async function refusedAndSame(
  room: Room,
  label: string,
  refusals: Map<Person, ReturnType<typeof expect.objectContaining>[]>,
  ...everyone: Person[]
): Promise<void> {
  for (const person of everyone) {
    expect({ label, who: person.name, refusals: person.refusals.splice(0) }).toEqual({
      label,
      who: person.name,
      refusals: refusals.get(person) ?? [],
    });
  }

  converged(room, ...everyone);
  expect(orderedCards(room.room.document)).toEqual([]);
  const expected = fingerprints(room.room.document);
  const saved = new Uint8Array(await everyone[0]!.editor.save());
  expect(readOoxmlPackage(saved).ok).toBe(true);
  expect(fingerprints(openOrderedWordFile(saved, { file: 2 }).document)).toEqual(expected);
  const { base } = openOrderedWordFile(FILE, { file: 1 });
  await checkFold(room.room.document, base, mediaStore(base), label);
}

const DECIDED = { code: 'ordered-refused', reason: 'no-inverse', detail: 'decided' };

type Edit = (ana: Person) => void;

/** Ana's new words at the start of the first paragraph. */
const type: Edit = (ana) => {
  caretAt(ana, 0, 0);
  surface(ana).insertPlainText('abcdef ');
};

/** Ana's edits of her own new words. */
const OWN_EDITS: readonly (readonly [string, Edit])[] = [
  [
    'typing into them',
    (ana) => {
      caretAt(ana, 0, 3);
      surface(ana).insertPlainText('XY');
    },
  ],
  [
    'a line pasted into them',
    (ana) => {
      caretAt(ana, 0, 3);
      surface(ana).pasteRich('XY\nZW', null);
    },
  ],
  [
    'a bold on part of them',
    (ana) => {
      caretAt(ana, 0, 1, 4);
      surface(ana).setRunProperty('b');
    },
  ],
  [
    'a deletion of part of them',
    (ana) => {
      caretAt(ana, 0, 2, 4);
      surface(ana).deleteSelection();
    },
  ],
];

describe('undo of an edit of your own new words after a decision (a decision is final)', () => {
  for (const [name, edit] of OWN_EDITS) {
    for (const status of ['accepted', 'rejected'] as const) {
      test(`${name}, ${status}: the undo is refused by name and nothing moves`, async () => {
        const { room, ana, olive } = await people();
        type(ana);
        await settle();
        edit(ana);
        await settle();
        await decideCard(olive, status);
        const shown = editorXml(ana);
        const log = kinds(room);
        surface(ana).undo();
        await settle();
        expect(kinds(room)).toEqual(log);
        expect(editorXml(ana)).toBe(shown);

        // Nothing to redo, and the step stays refused
        surface(ana).redo();
        await settle();
        expect(kinds(room)).toEqual(log);
        await refusedAndSame(
          room,
          `${name}, ${status}`,
          new Map([[ana, [expect.objectContaining(DECIDED)]]]),
          ana,
          olive
        );
      });
    }
  }

  test('the decision reaches the room first: the room refuses the take-back, the page takes it back', async () => {
    const { room, ana, olive } = await people();
    type(ana);
    await settle();
    OWN_EDITS[0]![1](ana);
    await settle();

    // Olive's accept is ordered before Ana's undo, which Ana's copy sends before it hears of it
    const [card] = pageCards(olive);
    expect(olive.editor.acceptReviewItem(card!.key).ok).toBe(true);
    olive.handle.session.flushPendingJournals();
    surface(ana).undo();
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
    expect(editorTexts(ana)[0]).toBe('abcXYdef The fee is 100 USD.');
    expect(editorXml(ana)).toBe(editorXml(olive));

    // The refused take-back leaves nothing to redo
    surface(ana).redo();
    await settle();
    expect(kinds(room)).toEqual(['insert', 'insert', 'decide']);
    await refusedAndSame(
      room,
      'the decision first',
      new Map([
        [ana, [expect.objectContaining({ reason: 'pending-refused', detail: 'not-pending' })]],
      ]),
      ana,
      olive
    );
  });

  test('a redo after the decision types the words again as a new suggestion', async () => {
    const { room, ana, olive } = await people();
    type(ana);
    await settle();
    OWN_EDITS[0]![1](ana);
    await settle();
    surface(ana).undo();
    await settle();
    const [decided] = orderedCards(room.room.document);
    await decideCard(olive, 'accepted');
    surface(ana).redo();
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(olive)[0]).toBe('abcXYdef The fee is 100 USD.');
    const cards = orderedCards(room.room.document);
    expect(cards).toHaveLength(1);
    expect(cards[0]!.suggestions).not.toContain(decided!.id);
    await decideCard(olive, 'accepted');
    await refusedAndSame(room, 'a redo after the decision', new Map(), ana, olive);
  });
});

describe("the organiser's undo of words typed inside a party's insertion decided since", () => {
  for (const status of ['accepted', 'rejected'] as const) {
    test(`${status}: refused by name, and nothing moves`, async () => {
      const { room, ana, olive } = await people();
      type(ana);
      await settle();
      caretAt(olive, 0, 3);
      surface(olive).insertPlainText('QQ');
      await settle();
      await decideCard(olive, status);
      const shown = editorXml(olive);
      const log = kinds(room);

      // The first undo reaches the decision, which has no undo; the second the typing
      surface(olive).undo();
      surface(olive).undo();
      await settle();
      expect(kinds(room)).toEqual(log);
      expect(editorXml(olive)).toBe(shown);
      expect(editorTexts(olive)[0]).toBe(
        status === 'accepted' ? 'abcQQdef The fee is 100 USD.' : 'The fee is 100 USD.'
      );

      // A rejected insertion shows nowhere: the typing step is spent, with nothing to take back
      const refused = status === 'accepted' ? 2 : 1;
      const refusals = Array.from({ length: refused }, () => expect.objectContaining(DECIDED));
      await refusedAndSame(room, `direct, ${status}`, new Map([[olive, refusals]]), ana, olive);
    });
  }
});

const FILE_TEXTS = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

const TYPED_LINES = ['one', 'two', 'threeThe fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

/**
 * Ana's edits that the page shows as one card of several suggestions, how many, and the texts
 * once accepted.
 */
const PIECES: readonly (readonly [string, (ana: Person) => Promise<void>, number, string[]])[] = [
  [
    'lines typed with Enter between them',
    async (ana) => {
      caretAt(ana, 0, 0);
      for (const step of ['one', null, 'two', null, 'three']) {
        if (step === null) {
          surface(ana).splitParagraph();
        } else {
          surface(ana).insertPlainText(step);
        }

        await settle();
      }
    },
    5,
    TYPED_LINES,
  ],
  [
    'lines struck a piece at a time',
    async (ana) => {
      caretAt(ana, 0, 4, 19);
      surface(ana).deleteSelection();
      await settle();
      caretAt(ana, 0, 19);
      surface(ana).deleteForward();
      await settle();
      caretAt(ana, 1, 0, 15);
      surface(ana).deleteSelection();
      await settle();
      caretAt(ana, 1, 15);
      surface(ana).deleteForward();
      await settle();
    },
    4,
    ['The Signed here.'],
  ],
  [
    'a new paragraph and its words',
    async (ana) => {
      caretAt(ana, 0, 19);
      surface(ana).splitParagraph();
      await settle();
      surface(ana).insertPlainText('New words.');
      await settle();
    },
    2,
    ['The fee is 100 USD.', 'New words.', 'Due in 30 days.', 'Signed here.'],
  ],
];

describe('one card on the page is decided whole', () => {
  for (const [name, edit, pieces, accepted] of PIECES) {
    for (const status of ['accepted', 'rejected'] as const) {
      test(`${name}, ${status}: one decide for every suggestion of the card`, async () => {
        const { room, ana, olive } = await people();
        await edit(ana);
        const doc = room.room.document;
        const suggestions = orderedCards(doc).flatMap((card) => card.suggestions);
        expect(suggestions).toHaveLength(pieces);

        // The card's marks name every suggestion; its first mark alone names one piece
        const [card] = pageCards(olive);
        if (card?.item.kind !== 'revision') {
          throw new Error('no revision card');
        }

        const named = (addresses: typeof card.item.addresses) =>
          addresses.flatMap((address) =>
            orderedSuggestionsOfRevision(doc, { part: BODY, ...address })
          );
        expect(named(card.item.addresses.slice(0, 1))).toHaveLength(1);
        expect(new Set(named(card.item.addresses))).toEqual(new Set(suggestions));
        const log = kinds(room);
        await decideCard(olive, status);
        expect(kinds(room)).toEqual([...log, 'decide']);
        expect(editorTexts(olive)).toEqual(status === 'accepted' ? accepted : FILE_TEXTS);
        await refusedAndSame(room, `${name}, ${status}`, new Map(), ana, olive);
      });
    }
  }

  for (const action of ['accept', 'reject'] as const) {
    test(`${action} all: one decide for every suggestion`, async () => {
      const { room, ana, olive } = await people();
      await PIECES[0]![1](ana);
      const log = kinds(room);
      const done = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action,
        scope: 'document',
      });
      expect(done.ok).toBe(true);
      await settle();
      expect(kinds(room)).toEqual([...log, 'decide']);
      expect(editorTexts(olive)).toEqual(action === 'accept' ? TYPED_LINES : FILE_TEXTS);
      await refusedAndSame(room, `${action} all`, new Map(), ana, olive);
    });
  }
});
