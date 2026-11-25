// A fold shape session: one room on `SHAPES_FILE` with the real apply and the suggestions-only
// policy, edited at random by an organiser and two parties (`randomShapeChange`). At the end of
// each round a fold keeps what the room uses (`checkFold`), and each card decided on its own
// gives the same decision and Word file live and after one and two folds (`checkEveryCard`).
// Then the room starts again from its checkpoint, as a host restarts it, and goes on, checked
// the same way. The sessions reach the shapes a fold once lost or changed: controls showing their
// prompts beside stored values, typing, formats and values over them, values given in pending
// controls, moves of words that hold note references, section breaks beside another person's
// section records, row heights and columns in a table with notes, and links made by one person
// and typed in by another.

import { describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../../change/apply.ts';
import { orderedChangeId } from '../../change/ids.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { OrderedRoom } from '../../room/room.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkEveryCard, checkFold, folded, mediaStore } from '../fold-check.ts';
import { seededRandom } from '../random.ts';
import { actor } from '../room-harness.ts';
import { randomShapeChange, SHAPE_NAMES } from './random-shape.ts';
import { SHAPES_FILE } from './shapes-fixture.ts';

const PEOPLE = [
  actor('Olive', 'olive', 'organiser'),
  actor('Ana', 'ana', 'party'),
  actor('Ben', 'ben', 'party'),
];

/** What a session did: the changes the room took, by shape, and the cards it decided. */
export interface ShapeSessionCounts {
  readonly taken: Map<string, number>;
  cards: number;
}

/** Run a session of `rounds` rounds of `steps` changes, then as many after a restart. */
export async function runShapeSession(
  seed: number,
  steps: number,
  rounds: number
): Promise<ShapeSessionCounts> {
  const rand = seededRandom(seed);
  const opened = openOrderedWordFile(SHAPES_FILE, { file: 0 });
  const blobs = mediaStore(opened.base);
  const counts: ShapeSessionCounts = { taken: new Map(), cards: 0 };
  let state = { document: opened.document, base: opened.base };
  let n = 0;
  let tick = 0;
  for (const stage of ['live', 'restarted']) {
    const room = new OrderedRoom({
      document: state.document,
      applyChange: applyOrderedChange,
      policy: suggestionsOnlyPolicy(),
      now: () => Date.UTC(2026, 10, 1) + 1000 * tick++,
    });
    for (let round = 0; round < rounds; round += 1) {
      for (let step = 0; step < steps; step += 1) {
        const who = PEOPLE[Math.floor(rand() * PEOPLE.length)]!;
        const person = { author: who.person.name, organiser: who.decides };
        const made = randomShapeChange(room.document, person, rand);
        if (made === null) {
          continue;
        }

        n += 1;
        const change = { v: 1, id: orderedChangeId(who.namespace, n), ...made.body };
        if (room.submit(who, [change]).accepted.length > 0) {
          counts.taken.set(made.shape, (counts.taken.get(made.shape) ?? 0) + 1);
        }
      }

      const label = `seed ${seed} ${stage} round ${round}`;
      await checkFold(room.document, state.base, blobs, label);
      counts.cards += await checkEveryCard(room.document, state.base, blobs, label);
    }

    state = await folded(room.document, state.base, room.document.file + 1, blobs);
  }

  return counts;
}

/**
 * A test of the session of each seed, and one that the sessions take a change of every shape and
 * decide cards. The seeds are a fixed list, split over a few files to keep each one short.
 */
export function describeShapeSessions(seeds: readonly number[]): void {
  const sessions: ShapeSessionCounts[] = [];
  describe('a fold of a random session over the shapes', () => {
    for (const seed of seeds) {
      test(`seed ${seed}: each copy gives back what the room uses, card by card`, async () => {
        sessions.push(await runShapeSession(seed, 40, 3));
      }, 120_000);
    }

    test('the sessions take a change of every shape and decide cards', () => {
      const taken = (shape: string) => sessions.some((s) => (s.taken.get(shape) ?? 0) > 0);
      expect(SHAPE_NAMES.filter((shape) => !taken(shape))).toEqual([]);
      expect(sessions.reduce((sum, s) => sum + s.cards, 0)).toBeGreaterThan(0);
    });
  });
}
