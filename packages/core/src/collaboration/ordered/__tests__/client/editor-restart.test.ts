// The connection and the room's storage (task 6.6, spec "Reconnect and restart"): a person who
// edits offline sends the edits when they come back; a room that restarts behind its clients
// (at 57, while the clients hold 61) gets each client's own lost changes again, under their own
// ids, and every copy ends equal to it; a room that moves to its next Word file sends `reload`,
// and the client keeps the changes it did not order.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import type { OrderedTransport } from '../../client/transport.ts';
import { decodeOrderedServerMessage, encodeOrderedServerMessage } from '../../protocol/codec.ts';
import {
  caretAt,
  closeAll,
  CODE,
  editorTexts,
  FILE,
  join,
  localRoom,
  PARTY,
  ROOM,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { createOrderedCollaboration } from '../../client/create.ts';
import { openOrderedRoomFile } from '../../client/open.ts';
import { checkpointBytes } from '../fold-check.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect({ who: person.name, refusals: person.refusals }).toEqual({
      who: person.name,
      refusals: [],
    });
    expect({ who: person.name, status: person.handle.session.status() }).toEqual({
      who: person.name,
      status: 'ready',
    });
    expect({ who: person.name, room: sameAsRoom(person, room) }).toEqual({
      who: person.name,
      room: true,
    });
    expect({ who: person.name, screen: screenFollowsList(person) }).toEqual({
      who: person.name,
      screen: true,
    });
  }
}

/** One keystroke at the end of a paragraph: one change. */
async function key(person: Person, paragraph: number, text: string): Promise<void> {
  const length = editorTexts(person)[paragraph]?.length ?? 0;
  caretAt(person, paragraph, length);
  surface(person).insertPlainText(text);
  await settle(2);
}

/** A connection whose second `welcome` loses its first change: the rejoin's catch-up has a gap. */
function gapOnRejoin(inner: OrderedTransport): OrderedTransport {
  let welcomes = 0;
  return {
    connect: (events) =>
      inner.connect({
        open: () => events.open(),
        close: (reason) => events.close(reason),
        message: (frame) => {
          const read = decodeOrderedServerMessage(frame);
          if (read.ok && read.value.type === 'welcome' && ++welcomes === 2) {
            events.message(
              encodeOrderedServerMessage({ ...read.value, changes: read.value.changes.slice(1) })
            );
            return;
          }
          events.message(frame);
        },
      }),
  };
}

describe('connection and storage', () => {
  test('edits made offline are sent on reconnect, and both copies end with the room', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    room.disconnect(ana.handle.session.namespace);
    await settle();
    expect(ana.handle.session.status()).toBe('disconnected');
    await key(ana, 0, ' Offline.');
    await key(ben, 1, ' Online.');
    expect(room.room.order).toBe(1);
    expect(ana.handle.session.unsentChanges()).toHaveLength(1);
    ana.handle.session.reconnect();
    await settle();
    converged(room, ana, ben);
    expect(room.room.order).toBe(2);
    expect(editorTexts(ben)).toEqual([
      'The fee is 100 USD. Offline.',
      'Due in 30 days. Online.',
      'Signed here.',
    ]);
  });

  test('a rejoin whose catch-up does not apply ends the session, and it stays ended', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana', {
      transport: gapOnRejoin(room.transport({ person: testPerson('Ana'), ...PARTY })),
    });
    const ben = await join(room, 'Ben');
    room.disconnect(ana.handle.session.namespace);
    await settle();
    await key(ben, 1, ' One.');
    await key(ben, 2, ' Two.');
    const statuses: string[] = [];
    ana.handle.session.subscribeStatus((status) => statuses.push(status));
    ana.handle.session.reconnect();
    await settle();
    // Before, the welcome set `ready` again after the failure: a session that sends nothing
    expect(statuses).toEqual(['error']);
    expect(ana.handle.session.statusSnapshot().reason?.code).toBe('remote-apply-failed');
  });

  test('a room that restarts at 57 while the clients hold 61 gets their own changes again', async () => {
    const room = localRoom(FILE, { autoSave: false });
    const ana = await join(room, 'Ana');
    const ben = await join(room, 'Ben');
    for (let k = 0; k < 57; k += 1)
      await key(k % 2 === 0 ? ana : ben, k % 3, String.fromCharCode(97 + (k % 26)));
    await settle();
    expect(room.room.order).toBe(57);
    room.save(57);
    await settle();
    await key(ana, 0, 'W');
    await key(ben, 1, 'X');
    await key(ana, 2, 'Y');
    await key(ben, 0, 'Z');
    await settle();
    expect(room.room.order).toBe(61);
    expect(ana.handle.session.order()).toBe(61);
    const lost = room.room.joinSnapshot(57).map((change) => change.id);
    const before = editorTexts(ana);
    // The room comes back from its storage at 57: the tail is gone
    room.restart(57);
    await settle();
    expect(room.room.order).toBe(57);
    expect(ana.handle.session.status()).toBe('disconnected');
    ana.handle.session.reconnect();
    ben.handle.session.reconnect();
    await settle(12);
    converged(room, ana, ben);
    expect(room.room.order).toBe(61);
    // Each client sent its own lost changes again, under their own ids
    expect(new Set(room.room.joinSnapshot(57).map((change) => change.id))).toEqual(new Set(lost));
    expect(editorTexts(ana)).toEqual(before);
    expect(editorTexts(ben)).toEqual(before);
  });

  test('a reconnect asks for the changes after the last save and repeats none', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    await key(ana, 0, ' One.');
    room.disconnect();
    await settle();
    ana.handle.session.reconnect();
    await settle();
    await key(ana, 0, ' Two.');
    await settle();
    converged(room, ana);
    expect(room.room.order).toBe(2);
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD. One. Two.');
  });

  test('a reload keeps the changes the room did not order, for the host', async () => {
    const room = localRoom();
    const ana = await join(room, 'Ana');
    room.disconnect(ana.handle.session.namespace);
    await settle();
    await key(ana, 0, ' Unsent.');
    const unsent = ana.handle.session.unsentChanges();
    expect(unsent).toHaveLength(1);
    ana.handle.session.reconnect();
    room.reloadAll('file');
    await settle();
    expect(ana.handle.session.status()).toBe('error');
    expect(ana.handle.session.statusSnapshot().reason?.code).toBe('room-generation-changed');
    expect(ana.handle.session.unsentChanges().map((change) => change.id)).toEqual(
      unsent.map((change) => change.id)
    );
    // The editor is read-only now: the gate refuses edits until the host loads again
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('No.');
    expect(surface(ana).state().lastRejection).toBe('collaboration-session-not-ready');
  });

  // Found by the editor acceptance cell (s2, seed 1000): after the room closed and opened its next
  // Word file, each editor's join waited 30 s, as the `reload` and the close came in one task (a
  // busy page reads its socket late) and the client dropped the frame with its connection
  test('a reload that comes in one task with its connection close still reaches the session', async () => {
    const queue: Array<() => void> = [];
    let held = false;
    const room = localRoom(FILE, {
      deliver: (run) => (held ? queue.push(run) : queueMicrotask(run)),
    });
    const ana = await join(room, 'Ana');
    held = true;
    room.reloadAll('file');
    // The frame, then the close, in one task
    while (queue.length > 0) queue.shift()?.();
    await settle();
    expect(ana.handle.session.status()).toBe('error');
    expect(ana.handle.session.statusSnapshot().reason?.code).toBe('room-generation-changed');
  });

  test('a client on another code version is told to reload, and does not join', async () => {
    const room = localRoom();
    await expect(
      createOrderedCollaboration({
        transport: room.transport({
          person: testPerson('Ana'),
          suggests: true,
          decides: false,
          readOnly: false,
          comments: true,
        }),
        room: ROOM,
        code: `${CODE}-old`,
        identity: { actorId: 'ana', name: 'Ana' },
        wordFile: async () => FILE,
        upload: room.uploader({
          person: testPerson('Ana'),
          suggests: true,
          decides: false,
          readOnly: false,
          comments: true,
        }),
        schedule: null,
      })
    ).rejects.toThrow('reload: version');
  });

  test('a Word file the host cannot read, or that does not open, ends the join by name', async () => {
    const room = localRoom();
    const login = {
      person: testPerson('Ana'),
      suggests: true,
      decides: false,
      readOnly: false,
      comments: true,
    };
    const attempt = (wordFile: () => Promise<Uint8Array>) =>
      createOrderedCollaboration({
        transport: room.transport(login),
        room: ROOM,
        code: CODE,
        identity: { actorId: 'ana', name: 'Ana' },
        wordFile,
        upload: room.uploader(login),
        schedule: null,
      });
    await expect(attempt(() => Promise.reject(new Error('offline')))).rejects.toThrow(
      'could not be read'
    );
    await expect(attempt(async () => new Uint8Array([1, 2, 3]))).rejects.toThrow(
      'cannot be opened'
    );
  });

  // A room's file keeps the spaces of its first Word file, so its own number is not the room's
  test('a room past two closes: a page types and reconnects, and the room asks for no reload', async () => {
    let file = FILE;
    for (const number of [1, 2]) {
      const room = localRoom(file, { file: number });
      const ana = await join(room, 'Ana', { wordFile: file });
      await key(ana, 0, ` Close ${number}.`);
      converged(room, ana);
      const { base } = openOrderedRoomFile(file, { file: number });
      file = await checkpointBytes(room.room.document, base, () => undefined);
      closeAll();
    }

    const room = localRoom(file, { file: 3 });
    const ana = await join(room, 'Ana', { wordFile: file });
    await key(ana, 0, ' Typed.');
    room.disconnect();
    await settle();
    ana.handle.session.reconnect();
    await settle();
    await key(ana, 0, ' Again.');
    await settle();
    converged(room, ana);
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD. Close 1. Close 2. Typed. Again.');
  });
});
