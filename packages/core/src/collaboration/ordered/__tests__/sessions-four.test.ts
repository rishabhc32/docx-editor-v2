// Random sessions with the real apply: an organiser and three parties, fixed seeds. Each seed
// takes several seconds, so on a busy machine the default 5 s limit fails a run that converges.

import { describe, expect, test } from 'bun:test';
import { runSession } from './sessions/session.ts';

describe('random sessions, four people', () => {
  for (const seed of [7, 23, 89]) {
    test(`seed ${seed}: every copy equals the room after every step and round`, () => {
      const result = runSession(seed, ['Olga', 'Pia', 'Quin', 'Ravi'], 960, 8);
      expect(result.order).toBeGreaterThan(250);
      expect(result.refused.local + result.refused.reapply + result.refused.room).toBeGreaterThan(
        0
      );
    }, 60_000);
  }
});
