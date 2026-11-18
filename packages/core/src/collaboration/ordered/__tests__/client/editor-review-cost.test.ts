// What a key press and a cursor move cost the editors' review queue in an ordered room (task
// 6.8): counts, not time. The queue of every suggestion and comment is derived from the whole
// document once; after that, a key reads again only the paragraph it touched, on the typist's
// editor (whose commit and the list's version of it reach the queue together) and on the other
// one. A key that adds to the person's own insertion is text too: the list's version of the
// paragraph keeps what the editor shows unchanged. A cursor move, the person's own or another's,
// changes no document and reads nothing.

import { afterEach, describe, expect, test } from 'bun:test';
import type { EditorModule, ReviewModuleContribution } from '../../../../contracts/modules.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const LONG = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.';
const inserted = (k: number) =>
  `<w:ins w:id="${900 + k}" w:author="Zed" w:date="2026-10-01T00:00:00Z">` +
  `<w:r><w:t xml:space="preserve"> added ${k}</w:t></w:r></w:ins>`;
// Paragraphs with properties, as a contract's are: the list's version of a key keeps the
// editor's `w:pPr`, or a key that adds to the person's insertion reads as a new structure
const FILE = docx(
  Array.from({ length: 60 }, (_, k) =>
    p(t(`${k} ${LONG}`) + (k % 5 === 0 ? inserted(k) : ''), '<w:spacing w:after="160"/>')
  ).join('')
);

interface Counts {
  /** Whole-document derivations of the queue. */
  full: number;
  /** One paragraph's suggestions read again. */
  paragraph: number;
}

/** The test review module, counting what it is asked. */
function counting(counts: Counts): EditorModule {
  const base = reviewTestModule();
  const review = base.review as ReviewModuleContribution;
  return {
    ...base,
    review: {
      ...review,
      collectReviewItems: (input) => {
        counts.full += 1;
        return review.collectReviewItems(input);
      },
      revisionItemsOfParagraph: (part, paragraphId) => {
        counts.paragraph += 1;
        return review.revisionItemsOfParagraph(part, paragraphId);
      },
    },
  };
}

interface Seat {
  readonly person: Person;
  readonly counts: Counts;
}

async function seat(
  room: ReturnType<typeof localRoom>,
  name: string,
  mode: 'edit' | 'suggesting'
): Promise<Seat> {
  const counts: Counts = { full: 0, paragraph: 0 };
  const person = await join(room, name, {
    mode,
    login: mode === 'edit' ? ORGANISER : PARTY,
    wordFile: FILE,
    review: counting(counts),
  });
  return { person, counts };
}

/** What `run` makes each editor's queue do, every queued task included. */
async function read(seats: readonly Seat[], run: () => void): Promise<Counts[]> {
  for (const { counts } of seats) Object.assign(counts, { full: 0, paragraph: 0 });
  run();
  await settle();
  return seats.map(({ counts }) => ({ ...counts }));
}

describe('the review queue in an ordered room', () => {
  for (const typistIsParty of [true, false]) {
    const who = typistIsParty ? 'a party suggesting' : 'the organiser editing';
    test(`${who}: a key reads its paragraph again on both editors, never the whole document`, async () => {
      const room = localRoom(FILE);
      const ana = await seat(room, 'Ana', 'suggesting');
      const org = await seat(room, 'Org', 'edit');
      const seats = [ana, org];
      const [typist] = typistIsParty ? [ana] : [org];
      const keys: readonly { readonly name: string; act(person: Person): void }[] = [
        { name: 'first key', act: (person) => surface(person).insertPlainText('x') },
        { name: 'next key', act: (person) => surface(person).insertPlainText('y') },
        { name: 'queued key', act: (person) => surface(person).enqueueType('q') },
        { name: 'Backspace', act: (person) => surface(person).deleteBackward() },
      ];
      await read(seats, () => caretAt(typist!.person, 12, 3));
      for (const key of keys) {
        const counted = await read(seats, () => key.act(typist!.person));
        expect({ key: key.name, full: counted.map((c) => c.full) }).toEqual({
          key: key.name,
          full: [0, 0],
        });
        for (const counts of counted) expect(counts.paragraph).toBeGreaterThan(0);
      }
      for (const { person } of seats) {
        expect(person.refusals).toEqual([]);
        expect(sameAsRoom(person, room)).toBe(true);
        expect(screenFollowsList(person)).toBe(true);
      }
    }, 60_000);
  }

  test('a cursor move, own or another person’s, reads nothing', async () => {
    const room = localRoom(FILE);
    const ana = await seat(room, 'Ana', 'suggesting');
    const org = await seat(room, 'Org', 'edit');
    const seats = [ana, org];
    await read(seats, () => {
      caretAt(ana.person, 12, 3);
      surface(ana.person).insertPlainText('x');
    });
    for (const move of [
      () => caretAt(org.person, 30, 5),
      () => caretAt(org.person, 31, 2, 6),
      () => caretAt(ana.person, 20, 1),
    ]) {
      const counted = await read(seats, move);
      expect(counted).toEqual([
        { full: 0, paragraph: 0 },
        { full: 0, paragraph: 0 },
      ]);
    }
  }, 60_000);
});
