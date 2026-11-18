// Two real editors in one room, and a room that closes one of them right after `hello` (a full
// room, a login the API cannot check) for a while (backlog G). The page must come back later each
// time, not every 250 ms, and once the room lets it in, it catches up and both copies converge.
// Each wait is spread between half and all of its step, so the test reads ranges.

import { afterEach, describe, expect, test } from 'bun:test';
import { ORDERED_STABLE_CONNECTION_MS, type OrderedSchedule } from '../../client/link.ts';
import type { OrderedTransport } from '../../client/transport.ts';
import {
  closeAll,
  editorTexts,
  join,
  localRoom,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
} from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

interface Timer {
  readonly run: () => void;
  readonly delayMs: number;
  cancelled: boolean;
}

// A room that reads `hello` and closes the connection, while `refusing` is set
function refusingTransport(inner: OrderedTransport, state: { refusing: boolean; hellos: number }) {
  const transport: OrderedTransport = {
    connect(events) {
      if (!state.refusing) return inner.connect(events);
      let ended = false;
      const connection = inner.connect({
        open: () => events.open(),
        message: (frame) => {
          if (!ended) events.message(frame);
        },
        close: (reason) => {
          if (ended) return;
          ended = true;
          events.close(reason);
        },
      });
      return {
        send() {
          if (ended) return false;
          state.hellos += 1;
          ended = true;
          connection.close();
          void Promise.resolve().then(() => events.close('closed 4429 room_full'));
          return true;
        },
        close() {
          ended = true;
          connection.close();
        },
      };
    },
  };
  return transport;
}

describe('a room that closes a page right after hello', () => {
  test('the page comes back later each time, then catches up and converges', async () => {
    const room = localRoom();
    const timers: Timer[] = [];
    const schedule: OrderedSchedule = (run, delayMs) => {
      const timer: Timer = { run, delayMs, cancelled: false };
      timers.push(timer);
      return () => {
        timer.cancelled = true;
      };
    };
    // The reconnect waits only: the join timeout and the stable timer are longer
    const waits = () =>
      timers.filter((t) => !t.cancelled && t.delayMs < ORDERED_STABLE_CONNECTION_MS);
    const runWait = async (): Promise<number> => {
      const due = waits();
      expect(due.length).toBe(1);
      due[0]!.cancelled = true;
      due[0]!.run();
      await settle();
      return due[0]!.delayMs;
    };
    const state = { refusing: false, hellos: 0 };
    const ana = await join(room, 'Ana');
    const bob = await join(room, 'Bob', {
      schedule,
      transport: refusingTransport(room.transport({ person: testPerson('Bob'), ...PARTY }), state),
    });
    await settle();

    // The network drops Bob; the room then closes him after each hello
    state.refusing = true;
    room.disconnect(bob.handle.session.namespace);
    await settle();
    expect(bob.handle.session.status()).toBe('disconnected');
    const got: number[] = [];
    for (let k = 0; k < 5; k += 1) got.push(await runWait());
    expect(state.hellos).toBe(5);
    const steps = [250, 500, 1000, 2000, 4000];
    got.forEach((wait, k) => {
      expect({ k, wait, atLeast: wait >= steps[k]! / 2, atMost: wait <= steps[k]! }).toEqual({
        k,
        wait,
        atLeast: true,
        atMost: true,
      });
    });
    expect(bob.handle.session.status()).toBe('disconnected');
    // The status says the room keeps closing the page, which try comes next, and when
    expect(bob.handle.session.statusSnapshot().reason?.detail).toMatch(
      /^closed 4429 room_full; try 6 in \d+\.\d s$/
    );

    // Ana writes meanwhile; Bob's own edit waits offline
    surface(ana).insertPlainText('red ');
    surface(bob).insertPlainText('big ');
    await settle();

    // The room lets Bob in: he catches up, his edit goes out, and both copies converge
    state.refusing = false;
    const wait = await runWait();
    expect(wait).toBeGreaterThanOrEqual(4000);
    expect(wait).toBeLessThanOrEqual(8000);
    await settle(16);
    expect(bob.handle.session.status()).toBe('ready');
    for (const person of [ana, bob]) {
      expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
        person: person.name,
        room: true,
      });
      expect(screenFollowsList(person)).toBe(true);
    }
    expect(editorTexts(ana)).toEqual(editorTexts(bob));
    expect(editorTexts(bob).join('|')).toContain('red ');
    expect(editorTexts(bob).join('|')).toContain('big ');

    // The welcomed connection stays up: the next drop waits the first step again
    const stable = timers.filter((t) => !t.cancelled && t.delayMs === ORDERED_STABLE_CONNECTION_MS);
    expect(stable.length).toBe(1);
    stable[0]!.cancelled = true;
    stable[0]!.run();
    room.disconnect(bob.handle.session.namespace);
    await settle();
    const again = await runWait();
    expect(again).toBeGreaterThanOrEqual(125);
    expect(again).toBeLessThanOrEqual(250);
    await settle(16);
    expect(bob.handle.session.status()).toBe('ready');
    expect(sameAsRoom(bob, room)).toBe(true);
  });
});
