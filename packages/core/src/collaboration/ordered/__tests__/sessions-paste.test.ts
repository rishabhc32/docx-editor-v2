// Random sessions with group 9.10's changes too (a paste with formatting, a new list, and for
// the organiser a paragraph removal and Protect Document), with the real apply: every copy
// equals the room after every step and round, decisions give one result in every order, and
// the saved file reopens to the same view and decisions.

import { describe, expect, test } from 'bun:test';
import { runSession } from './sessions/session.ts';

describe('random sessions with pastes and new lists', () => {
  // Seed 71 found a split link that kept its halves after the link went (the model case:
  // `change-apply.test.ts`, a suggested Enter inside another person's pending link). Seed 5
  // found a row made after a decided column that the saved file drew on the old grid (the model
  // case: `change-settle-columns.test.ts`). Each seed takes several seconds, so on a busy
  // machine the default 5 s limit fails a run that converges
  for (const [seed, people] of [
    [5, ['Olga', 'Ana']],
    [19, ['Olga', 'Ana']],
    [71, ['Olga', 'Ana']],
    [7, ['Olga', 'Pia', 'Quin', 'Ravi']],
  ] as const) {
    test(`seed ${seed}, ${people.length} people`, () => {
      const result = runSession(seed, people, people.length > 2 ? 960 : 720, 8, undefined, true);
      expect(result.order).toBeGreaterThan(200);
    }, 60_000);
  }
});
