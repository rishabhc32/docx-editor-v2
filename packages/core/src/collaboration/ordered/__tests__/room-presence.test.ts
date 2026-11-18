// Presence follows the host's presence rules: own entries only, under a size cap, named from the
// login, and read-only people may send it.

import { describe, expect, test } from 'bun:test';
import { admitOrderedPresence } from '../room/presence.ts';
import { actor } from './room-harness.ts';
import { testPerson } from './test-persons.ts';

const ANA = actor('Ana', 'ana-1', 'party');
const RITA = actor('Rita', 'rita-1', 'reader');

describe('presence', () => {
  const state = {
    selection: { anchor: { paragraph: 'p1', offset: 0 }, head: { paragraph: 'p1', offset: 4 } },
  };

  test('is relayed for the sender own namespace, named from the login', () => {
    const verdict = admitOrderedPresence(RITA, {
      v: 3,
      type: 'presence',
      namespace: 'rita-1',
      state,
    });
    expect(verdict).toEqual({
      ok: true,
      message: { v: 3, type: 'presence', namespace: 'rita-1', person: testPerson('Rita'), state },
    });
  });

  test('is dropped for another namespace, a namespace another person owns, or over the cap', () => {
    expect(
      admitOrderedPresence(ANA, { v: 3, type: 'presence', namespace: 'ben-1', state })
    ).toEqual({
      ok: false,
      code: 'foreign-presence',
    });
    expect(
      admitOrderedPresence(
        ANA,
        { v: 3, type: 'presence', namespace: 'ana-1', state },
        { ownerOf: () => 'Ben' }
      )
    ).toEqual({ ok: false, code: 'foreign-presence' });
    const big = { selection: null, data: { note: 'x'.repeat(200) } };
    expect(
      admitOrderedPresence(
        ANA,
        { v: 3, type: 'presence', namespace: 'ana-1', state: big },
        { maxLength: 100 }
      )
    ).toEqual({ ok: false, code: 'presence-too-large' });
  });
});
