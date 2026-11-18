// The link's waits before it connects again. A room that closes the page right after
// `hello` (a full room, a login the API cannot check, a room that cannot open) must not be joined
// again every 250 ms for as long as the page stays open: each join costs the room server a call to
// the host's API, and a fresh copy's catch-up. The wait grows until the room welcomes the page and
// the connection stays up, and it is spread so that the pages of one room do not come back
// together.

import { describe, expect, test } from 'bun:test';
import {
  ORDERED_RECONNECT_DELAYS,
  ORDERED_STABLE_CONNECTION_MS,
  OrderedLink,
  type OrderedLinkHost,
  type OrderedReconnect,
  type OrderedSchedule,
} from '../../client/link.ts';
import type { OrderedConnectionEvents, OrderedTransport } from '../../client/transport.ts';

interface Timer {
  readonly run: () => void;
  readonly delayMs: number;
  cancelled: boolean;
}

function fixture(random: () => number = () => 1) {
  const timers: Timer[] = [];
  const schedule: OrderedSchedule = (run, delayMs) => {
    const timer: Timer = { run, delayMs, cancelled: false };
    timers.push(timer);
    return () => {
      timer.cancelled = true;
    };
  };
  const connections: OrderedConnectionEvents[] = [];
  const transport: OrderedTransport = {
    connect(events) {
      connections.push(events);
      return { send: () => true, close: () => {} };
    },
  };
  const closed: Array<{ reason: string; next: OrderedReconnect | null }> = [];
  const host: OrderedLinkHost = {
    hello: () => ({
      v: 3,
      type: 'hello',
      room: 'room',
      file: 1,
      last: 0,
      code: 'test',
      namespace: 'ns-a',
    }),
    message: () => {},
    opened: () => {},
    closed: (reason, next) => closed.push({ reason, next }),
    failed: (error) => {
      throw error;
    },
  };
  const link = new OrderedLink(transport, host, schedule, ORDERED_RECONNECT_DELAYS, { random });
  const live = () => timers.filter((t) => !t.cancelled);
  // Runs the one reconnect wait that is due, and returns its length
  const waitAndConnect = async (): Promise<number> => {
    const waits = live().filter((t) => t.delayMs !== ORDERED_STABLE_CONNECTION_MS);
    expect(waits.length).toBe(1);
    const wait = waits[0]!;
    wait.cancelled = true;
    wait.run();
    return wait.delayMs;
  };
  // The room opens the socket, reads `hello`, and closes it
  const refuse = async (reason = 'closed 4429 room_full'): Promise<void> => {
    const events = connections[connections.length - 1]!;
    events.open();
    events.close(reason);
    // The close reaches the session after the messages before it, a microtask later
    await Promise.resolve();
    await Promise.resolve();
  };
  return { link, timers, live, connections, closed, waitAndConnect, refuse };
}

describe('ordered link reconnect waits', () => {
  test('a room that closes the page after hello is joined again later each time, up to the cap', async () => {
    const f = fixture();
    f.link.start();
    const waits: number[] = [];
    for (let k = 0; k < 9; k += 1) {
      await f.refuse();
      waits.push(await f.waitAndConnect());
    }
    expect(waits).toEqual([250, 500, 1000, 2000, 4000, 8000, 15000, 15000, 15000]);
    expect(f.connections.length).toBe(10);
    // The session hears which try comes next and how long the link waits for it
    expect(f.closed.map((c) => c.next)).toEqual(
      waits.map((delayMs, k) => ({ attempt: k + 1, delayMs }))
    );
    expect(f.closed[0]?.reason).toBe('closed 4429 room_full');
  });

  test('each wait is spread between half and all of its step, never past the cap', async () => {
    const draws = [0, 0.5, 0.999, 0, 0.25, 1, 0.75, 0, 1];
    let at = 0;
    const f = fixture(() => draws[at++ % draws.length]!);
    f.link.start();
    const waits: number[] = [];
    for (let k = 0; k < draws.length; k += 1) {
      await f.refuse();
      waits.push(await f.waitAndConnect());
    }
    const steps = [250, 500, 1000, 2000, 4000, 8000, 15000, 15000, 15000];
    waits.forEach((wait, k) => {
      expect(wait).toBeGreaterThanOrEqual(steps[k]! / 2);
      expect(wait).toBeLessThanOrEqual(steps[k]!);
    });
    expect(waits[0]).toBe(125);
    expect(waits[5]).toBe(8000);
  });

  test('a welcome alone does not start the waits again: the connection must stay up', async () => {
    const f = fixture();
    f.link.start();
    const waits: number[] = [];
    for (let k = 0; k < 4; k += 1) {
      const events = f.connections[f.connections.length - 1]!;
      events.open();
      // The room welcomes the page, then drops it before the connection is stable
      f.link.welcomed();
      events.close('closed 1011 error');
      await Promise.resolve();
      await Promise.resolve();
      waits.push(await f.waitAndConnect());
    }
    expect(waits).toEqual([250, 500, 1000, 2000]);
  });

  test('a welcomed connection that stays up starts the waits again from the first', async () => {
    const f = fixture();
    f.link.start();
    for (let k = 0; k < 4; k += 1) {
      await f.refuse();
      await f.waitAndConnect();
    }
    const events = f.connections[f.connections.length - 1]!;
    events.open();
    f.link.welcomed();
    const stable = f.live().filter((t) => t.delayMs === ORDERED_STABLE_CONNECTION_MS);
    expect(stable.length).toBe(1);
    stable[0]!.run();
    events.close('closed 1006');
    await Promise.resolve();
    await Promise.resolve();
    expect(await f.waitAndConnect()).toBe(250);
  });

  test('a stopped link tells the session it does not connect again', async () => {
    const f = fixture();
    f.link.start();
    const events = f.connections[0]!;
    events.open();
    f.link.stop();
    events.close('closed 1000');
    await Promise.resolve();
    await Promise.resolve();
    // A close after stop is the link's own: no event, no wait
    expect(f.closed).toEqual([]);
    expect(f.live()).toEqual([]);
  });

  test('stop cancels the wait and the stable timer', async () => {
    const f = fixture();
    f.link.start();
    const events = f.connections[0]!;
    events.open();
    f.link.welcomed();
    f.link.stop();
    expect(f.live()).toEqual([]);
  });
});
