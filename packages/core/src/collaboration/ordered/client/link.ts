import { ORDERED_CHANGE_LIMITS, type OrderedChange } from '../change/types.ts';
import { decodeOrderedServerMessage, encodeOrderedClientMessage } from '../protocol/codec.ts';
import {
  ORDERED_PROTOCOL_LIMITS,
  ORDERED_PROTOCOL_VERSION,
  type OrderedClientMessage,
  type OrderedHelloMessage,
  type OrderedServerMessage,
} from '../protocol/messages.ts';
import type { OrderedConnection, OrderedTransport } from './transport.ts';

/** A frame from the room the codec does not read: the room speaks another protocol. */
export class OrderedFrameError extends Error {
  readonly code: string;

  constructor(code: string) {
    super(`a frame the client cannot read: ${code}`);
    this.name = 'OrderedFrameError';
    this.code = code;
  }
}

/** The host could not give the room's Word file. */
export class OrderedWordFileUnavailable extends Error {
  constructor(cause: unknown) {
    super(
      `the room's Word file could not be read: ${cause instanceof Error ? cause.name : typeof cause}`
    );
    this.name = 'OrderedWordFileUnavailable';
  }
}

/** Run `run` after `delayMs`; returns how to cancel it. @public */
export type OrderedSchedule = (run: () => void, delayMs: number) => () => void;

/** What the link asks of the session. */
export interface OrderedLinkHost {
  /** The `hello` for a new connection. */
  hello(): OrderedHelloMessage;
  /** One message from the room, in order. A rejection ends the session. */
  message(message: OrderedServerMessage): void | Promise<void>;
  /** The connection is open and `hello` sent. */
  opened(): void;
  /**
   * The connection dropped. `next` is the try the link makes and how long it waits first; null
   * when it does not connect again (stopped, or no timers).
   */
  closed(reason: string, next: OrderedReconnect | null): void;
  /** A frame the session cannot read, or a message it could not take: the link stops. */
  failed(error: unknown): void;
}

/**
 * The waits before each new connection after a drop, in milliseconds; the last repeats. Each wait
 * is spread between half and all of its step.
 */
export const ORDERED_RECONNECT_DELAYS: readonly number[] = Object.freeze([
  250, 500, 1000, 2000, 4000, 8000, 15000,
]);

/**
 * How long a connection the room welcomed stays up before the waits start again from the first.
 * A room that closes the page right after `hello`, or right after its welcome, is joined again
 * later each time: every join costs the room server a login check and a catch-up.
 */
export const ORDERED_STABLE_CONNECTION_MS = 10_000;

/** The next connection after a drop: its number since the last stable one, and the wait. */
export interface OrderedReconnect {
  readonly attempt: number;
  readonly delayMs: number;
}

/** How the link spreads its waits. */
export interface OrderedLinkOptions {
  /** A number from 0 to 1. Default: the host's `Math.random`. */
  readonly random?: () => number;
}

/** The host's timers, read where any window will do: a reconnect wait belongs to no page. */
export function defaultOrderedSchedule(): OrderedSchedule | null {
  const scope: { readonly setTimeout?: unknown; readonly clearTimeout?: unknown } = globalThis;
  const set = scope.setTimeout;
  const clear = scope.clearTimeout;
  if (typeof set !== 'function' || typeof clear !== 'function') return null;
  return (run, delayMs) => {
    const handle: unknown = Reflect.apply(set, globalThis, [run, delayMs]);
    return () => {
      Reflect.apply(clear, globalThis, [handle]);
    };
  };
}

/** Changes in frames under the batch cap and the client frame cap. */
export function changeFrames(changes: readonly OrderedChange[]): string[] {
  const out: string[] = [];
  const pack = (group: readonly OrderedChange[]): void => {
    if (group.length === 0) return;
    const frame = encodeOrderedClientMessage({
      v: ORDERED_PROTOCOL_VERSION,
      type: 'changes',
      changes: group,
    });
    if (frame.length <= ORDERED_PROTOCOL_LIMITS.clientFrame || group.length === 1) {
      out.push(frame);
      return;
    }

    const half = Math.ceil(group.length / 2);
    pack(group.slice(0, half));
    pack(group.slice(half));
  };
  for (let at = 0; at < changes.length; at += ORDERED_CHANGE_LIMITS.batch)
    pack(changes.slice(at, at + ORDERED_CHANGE_LIMITS.batch));
  return out;
}

/**
 * The protocol link (design D8): one connection at a time over the host's transport. It sends
 * `hello` when a connection opens, hands the room's messages to the session one after another,
 * and connects again after a drop, later each time until a welcomed connection stays up.
 */
export class OrderedLink {
  readonly #transport: OrderedTransport;
  readonly #host: OrderedLinkHost;
  readonly #schedule: OrderedSchedule | null;
  readonly #delays: readonly number[];
  readonly #random: () => number;
  #connection: OrderedConnection | null = null;
  #open = false;
  // Drops since the last connection that stayed up; it picks the next wait
  #attempt = 0;
  #cancelWait: (() => void) | null = null;
  #cancelStable: (() => void) | null = null;
  #stopped = false;
  #chain: Promise<void> = Promise.resolve();
  #generation = 0;

  constructor(
    transport: OrderedTransport,
    host: OrderedLinkHost,
    schedule: OrderedSchedule | null,
    delays: readonly number[] = ORDERED_RECONNECT_DELAYS,
    options: OrderedLinkOptions = {}
  ) {
    this.#transport = transport;
    this.#host = host;
    this.#schedule = schedule;
    this.#delays = delays.length > 0 ? delays : ORDERED_RECONNECT_DELAYS;
    this.#random = options.random ?? Math.random;
  }

  /** Whether a connection is open and `hello` went out. */
  get open(): boolean {
    return this.#open;
  }

  start(): void {
    if (this.#stopped || this.#connection) return;
    this.#cancelWait?.();
    this.#cancelWait = null;
    const generation = (this.#generation += 1);
    const current = (): boolean => generation === this.#generation && !this.#stopped;
    this.#connection = this.#transport.connect({
      open: () => {
        if (!current()) return;

        // The waits do not start again here: a room may close the page right after `hello`
        this.#open = true;
        this.#write(encodeOrderedClientMessage(this.#host.hello()));
        this.#host.opened();
      },
      message: (frame) => {
        if (!current()) return;
        const decoded = decodeOrderedServerMessage(frame);
        if (!decoded.ok) {
          this.#fail(new OrderedFrameError(decoded.error.code));
          return;
        }

        const message = decoded.value;
        this.#chain = this.#chain
          .then(() => (current() ? this.#host.message(message) : undefined))
          .catch((error: unknown) => this.#fail(error));
      },
      close: (reason) => {
        if (!current()) return;
        this.#connection = null;
        this.#open = false;

        // The messages that came before the close reach the session first, in order: the room
        // sends `reload` right before it closes the socket, and a page that reads its socket late
        // gets both at once
        this.#chain = this.#chain
          .then(() => (current() ? this.#dropped(reason) : undefined))
          .catch((error: unknown) => this.#fail(error));
      },
    });
  }

  /** Connect again now, closing the connection there is. */
  reconnect(): void {
    if (this.#stopped) return;
    this.#close();
    this.start();
  }

  /**
   * The room welcomed this connection. Once it stays up for `ORDERED_STABLE_CONNECTION_MS`, the
   * next drop waits the first step again.
   */
  welcomed(): void {
    if (this.#stopped || !this.#open) return;
    this.#cancelStable?.();
    if (this.#schedule === null) {
      this.#attempt = 0;
      return;
    }

    const generation = this.#generation;
    this.#cancelStable = this.#schedule(() => {
      this.#cancelStable = null;
      if (generation === this.#generation && this.#open) this.#attempt = 0;
    }, ORDERED_STABLE_CONNECTION_MS);
  }

  /** Stop for good. */
  stop(): void {
    this.#stopped = true;
    this.#cancelWait?.();
    this.#cancelWait = null;
    this.#close();
  }

  /** Send a message; false while no connection is open. */
  send(message: OrderedClientMessage): boolean {
    return this.#open && this.#write(encodeOrderedClientMessage(message));
  }

  /** Send changes in frames under the caps; false while no connection is open. */
  sendChanges(changes: readonly OrderedChange[]): boolean {
    if (!this.#open) return false;
    for (const frame of changeFrames(changes)) if (!this.#write(frame)) return false;
    return true;
  }

  #write(frame: string): boolean {
    return this.#connection?.send(frame) ?? false;
  }

  /** Forget the connection: its late callbacks no longer count. Returns the connection. */
  #detach(): OrderedConnection | null {
    this.#generation += 1;
    this.#cancelStable?.();
    this.#cancelStable = null;
    const connection = this.#connection;
    this.#connection = null;
    this.#open = false;
    return connection;
  }

  #close(): void {
    this.#detach()?.close();
  }

  #dropped(reason: string): void {
    this.#detach();
    const schedule = this.#schedule;
    if (this.#stopped || schedule === null) {
      this.#host.closed(reason, null);
      return;
    }

    const step = this.#delays[Math.min(this.#attempt, this.#delays.length - 1)] ?? 15000;

    // Spread between half and all of the step, so the pages a room dropped together come back apart
    const spread = Math.min(1, Math.max(0, this.#random()));
    const delay = Math.round(step / 2 + (step / 2) * spread);
    this.#attempt += 1;
    this.#host.closed(reason, { attempt: this.#attempt, delayMs: delay });
    if (this.#stopped) return;
    this.#cancelWait = schedule(() => {
      this.#cancelWait = null;
      this.start();
    }, delay);
  }

  #fail(error: unknown): void {
    if (this.#stopped) return;
    this.stop();
    this.#host.failed(error);
  }
}
