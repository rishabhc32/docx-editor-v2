// A pasted picture's bytes go through the host's upload (design D14), with two real editors and
// one room: the paste and every later edit wait until the room's storage holds the bytes, a
// refusal of the storage takes the paste back by name, and a failed upload is tried again.

import { afterEach, describe, expect, test } from 'bun:test';
import { ORDERED_RECONNECT_DELAYS, type OrderedSchedule } from '../../client/link.ts';
import type { OrderedBlobRefusalCode } from '../../protocol/blobs.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { RED_PNG } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  FILE,
  HeldUpload,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedPicture } from './paste-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const RED_HASH = orderedBlobHash(RED_PNG);

/** Two people: the organiser, and a party whose host uploads through `held`. */
async function room(schedule: OrderedSchedule | null = null) {
  const r = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const held = new HeldUpload(r.uploader({ person: testPerson('Ben'), ...PARTY }));
  const olive = await join(r, 'Olive', { mode: 'edit', login: ORGANISER });
  const ben = await join(r, 'Ben', { login: PARTY, upload: held.upload, schedule });
  return { r, held, olive, ben };
}

/** Paste a copied picture into a person's editor, then type a word after it. */
async function pasteThenType(person: Person): Promise<void> {
  const { text, html } = copiedPicture();
  caretAt(person, 1, 3);
  expect(surface(person).pasteRich(text, html)).toBe(true);
  await settle(8);
  caretAt(person, 2, 0);
  person.editor.exec({ type: 'insertText', text: 'Later ' });
  await settle(8);
}

const pictures = (person: Person) => editorXml(person).match(/<w:drawing>/g)?.length ?? 0;

describe('a pasted picture and the upload', () => {
  test('the paste and every later edit wait until the storage holds the bytes', async () => {
    const { r, held, olive, ben } = await room();
    await pasteThenType(ben);
    expect(held.waiting.map((upload) => upload.sha256)).toEqual([RED_HASH]);
    expect(r.room.order).toBe(0);
    expect(ben.handle.session.unsentChanges()).toHaveLength(2);
    expect(pictures(olive)).toBe(0);

    held.goAll();
    await settle(24);
    expect(r.blob(RED_HASH)).toEqual(RED_PNG);
    expect(r.room.joinSnapshot(0).map((change) => change.kind)).toEqual([
      'insert-fragment',
      'insert',
    ]);
    expect(ben.refusals).toEqual([]);
    for (const person of [olive, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(pictures(person)).toBe(1);
      expect(editorXml(person)).toContain('Later ');
    }
  });

  const codes: readonly OrderedBlobRefusalCode[] = [
    'blob-too-large',
    'blob-hash',
    'blob-type',
    'blob-quota',
  ];
  for (const code of codes) {
    test(`a refused upload (${code}) takes the paste back by name, and later edits go`, async () => {
      const { r, held, olive, ben } = await room();
      await pasteThenType(ben);
      held.waiting[0]?.answer({ stored: false, code });
      await settle(24);
      expect(ben.refusals.map((x) => [x.reason, x.detail])).toEqual([['blob-refused', code]]);
      expect(r.blob(RED_HASH)).toBeUndefined();
      expect(r.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert']);
      for (const person of [olive, ben]) {
        expect(sameAsRoom(person, r)).toBe(true);
        expect(pictures(person)).toBe(0);
        expect(editorXml(person)).toContain('Later ');
      }
    });
  }

  test('a failed upload goes again after the reconnect delays, and lands after a rejoin', async () => {
    const timers: { run: () => void; delayMs: number; cancelled: boolean }[] = [];
    const schedule: OrderedSchedule = (run, delayMs) => {
      const timer = { run, delayMs, cancelled: false };
      timers.push(timer);
      return () => {
        timer.cancelled = true;
      };
    };
    const waits = () =>
      timers.filter((t) => !t.cancelled && ORDERED_RECONNECT_DELAYS.includes(t.delayMs));
    const { r, held, olive, ben } = await room(schedule);
    await pasteThenType(ben);
    held.waiting.shift()?.fail();
    await settle(4);
    expect(waits().map((t) => t.delayMs)).toEqual([ORDERED_RECONNECT_DELAYS[0]]);

    // The wait ends: the upload goes again, and fails again
    const first = waits()[0];
    if (!first) {
      throw new Error('no wait');
    }

    first.cancelled = true;
    first.run();
    expect(held.waiting).toHaveLength(1);
    held.waiting.shift()?.fail();
    await settle(4);
    expect(waits().map((t) => t.delayMs)).toEqual([ORDERED_RECONNECT_DELAYS[1]]);
    expect(r.room.order).toBe(0);

    // A rejoin tries again at once, before the wait ends
    ben.handle.session.reconnect();
    await settle(8);
    expect(waits()).toEqual([]);
    expect(held.waiting).toHaveLength(1);
    held.goAll();
    await settle(24);
    expect(ben.refusals).toEqual([]);
    for (const person of [olive, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(pictures(person)).toBe(1);
      expect(editorXml(person)).toContain('Later ');
    }
  });
});
