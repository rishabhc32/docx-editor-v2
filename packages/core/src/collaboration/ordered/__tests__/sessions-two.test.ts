// Random sessions with the real apply: an organiser and a party, fixed seeds.

import { describe, expect, test } from 'bun:test';
import { runSession } from './sessions/session.ts';

describe('random sessions, two people', () => {
  // Seed 71: a moved word inside a plain text control, saved and reopened, took the `w:id` of
  // its move's other half (fixed with the move range markers in every inline container;
  // `move-into-link.test.ts` holds the deterministic case).
  test('seed 71: a move inside a control keeps its own w:id after a save', () => {
    runSession(71, ['Olga', 'Ana'], 720, 8);
  });
  // Seed 103: the reopened file kept an empty pending link that held a comment's start, which
  // the live copy dropped after a decision (`link-left-with-markers.test.ts` holds the
  // deterministic case).
  test('seed 103: a decided link leaves no empty link around a comment start', () => {
    runSession(103, ['Olga', 'Ana'], 720, 8);
  });
  for (const seed of [5, 19]) {
    test(`seed ${seed}: every copy equals the room after every step and round`, () => {
      const result = runSession(seed, ['Olga', 'Ana'], 720, 8);
      expect(result.order).toBeGreaterThan(200);
      expect(result.refused.local + result.refused.reapply + result.refused.room).toBeGreaterThan(
        0
      );
    });
  }
});
