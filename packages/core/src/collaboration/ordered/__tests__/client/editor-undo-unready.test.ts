// Undo and redo while the room session cannot write, with two real editors
// and one room.
//
// A session that cannot write — offline with offline editing turned off — reports no step to
// undo, so `exec` refused the undo with "nothing to undo" while the person's own edit was
// still there. The refusal now gives the session's reason, through `exec` and the keyboard
// lane, and after reconnect the same undo runs.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const FILE = docx(p(t('First words.')) + p(t('Second words.')));
const NOT_READY = 'collaboration-session-not-ready';

const bodyText = (person: Person) => surface(person).session.storyText({ kind: 'body' });

async function people(offline: 'Olive' | 'Ana') {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    login: ORGANISER,
    mode: 'edit',
    offlineEditing: offline !== 'Olive',
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, offlineEditing: offline !== 'Ana' });
  return { room, olive, ana, person: offline === 'Olive' ? olive : ana };
}

/** The person types, then loses the connection. */
async function typeThenGoOffline(room: ReturnType<typeof localRoom>, person: Person) {
  caretAt(person, 0, 5);
  expect(person.editor.exec({ type: 'insertText', text: ' typed' }).ok).toBe(true);
  await settle();
  room.disconnect(person.handle.session.namespace);
  await settle();
  expect(person.handle.session.status()).toBe('disconnected');
}

describe('undo while the session cannot write gives the session reason', () => {
  for (const who of ['Olive', 'Ana'] as const) {
    test(`${who}: exec undo and redo`, async () => {
      const { room, olive, ana, person } = await people(who);
      await typeThenGoOffline(room, person);
      const typed = bodyText(person);

      expect(person.editor.exec({ type: 'undo' })).toEqual({
        ok: false,
        code: 'unsupported',
        reason: NOT_READY,
      });
      expect(person.editor.exec({ type: 'redo' })).toEqual({
        ok: false,
        code: 'unsupported',
        reason: NOT_READY,
      });
      expect(bodyText(person)).toBe(typed);

      // Back online, the same undo runs and both copies end with the room.
      person.handle.session.reconnect();
      await settle();
      expect(person.editor.exec({ type: 'undo' })).toEqual({ ok: true, changed: true });
      await settle();
      expect(bodyText(olive)).toBe(bodyText(ana));
      expect(bodyText(person)).not.toBe(typed);
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(sameAsRoom(ana, room)).toBe(true);
    });

    test(`${who}: the keyboard lane reports the same reason`, async () => {
      const { room, person } = await people(who);
      await typeThenGoOffline(room, person);
      const typed = bodyText(person);

      surface(person).undo();
      expect(surface(person).state().lastRejection).toBe(NOT_READY);
      expect(bodyText(person)).toBe(typed);
    });
  }

  test('a ready session with no step still says "nothing to undo"', async () => {
    const { olive } = await people('Ana');
    expect(olive.editor.exec({ type: 'undo' })).toEqual({
      ok: false,
      code: 'unsupported',
      reason: 'nothing to undo',
    });
  });
});
