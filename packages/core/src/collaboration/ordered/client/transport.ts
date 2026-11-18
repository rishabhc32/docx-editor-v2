// Transports (design D8, task 6.5): they move text frames and say when a connection opens and
// ends; the session reads and writes the protocol.

import { ORDERED_PROTOCOL_LIMITS } from '../protocol/messages.ts';

/** What one connection tells the session. @public */
export interface OrderedConnectionEvents {
  /** The connection is open: the session sends `hello`. */
  open(): void;
  /** One text frame from the room. */
  message(frame: string): void;
  /** The connection ended, once, with a short reason for logs. Not called after `close()`. */
  close(reason: string): void;
}

/** One connection to the room. @public */
export interface OrderedConnection {
  /** Send one text frame; false when the connection is not open. */
  send(frame: string): boolean;
  /** End the connection. No `close` event follows. */
  close(): void;
}

/** Opens connections to one room. Each `connect` is a new connection. @public */
export interface OrderedTransport {
  connect(events: OrderedConnectionEvents): OrderedConnection;
}

/**
 * The part of a WebSocket the transport uses. A browser's `WebSocket`, Node's global one and
 * the `ws` package's all have it.
 *
 * @public
 */
export interface OrderedSocket {
  /** 1 while open, as the WebSocket standard numbers it. */
  readonly readyState: number;
  send(data: string): void;
  close(code?: number, reason?: string): void;
  addEventListener(
    type: 'open' | 'message' | 'close' | 'error',
    listener: (event: unknown) => void
  ): void;
}

/** How the WebSocket transport connects. @public */
export interface OrderedWebSocketOptions {
  readonly url: string;
  /**
   * Open a socket to `url`: the host's WebSocket, with its subprotocols and credentials. Called
   * once for every connection, so a reconnect gets a new socket. A transport never reaches for a
   * script global: it takes the host's own socket, so it runs in a browser, in another window,
   * and on Node.
   */
  readonly socket: (url: string) => OrderedSocket;
  /** Longest frame read, in UTF-16 units. Default and ceiling: the protocol's server frame cap. */
  readonly maxFrame?: number;
}

const OPEN = 1;

/** A message event's data, read without trusting its shape. */
function dataOf(event: unknown): unknown {
  return typeof event === 'object' && event !== null && 'data' in event ? event.data : undefined;
}

/** A close event's code and reason, as a short note. */
function closeNote(event: unknown): string {
  if (typeof event !== 'object' || event === null) return 'closed';
  const code = 'code' in event && typeof event.code === 'number' ? event.code : null;
  const reason = 'reason' in event && typeof event.reason === 'string' ? event.reason : '';
  return `closed${code === null ? '' : ` ${code}`}${reason ? ` ${reason.slice(0, 120)}` : ''}`;
}

/**
 * A transport over the host's WebSocket. Text frames only: a binary frame, or one past the
 * frame cap, ends the connection.
 *
 * @public
 */
export function orderedWebSocketTransport(options: OrderedWebSocketOptions): OrderedTransport {
  const ceiling = ORDERED_PROTOCOL_LIMITS.serverFrame;
  const max =
    options.maxFrame !== undefined && Number.isSafeInteger(options.maxFrame) && options.maxFrame > 0
      ? Math.min(options.maxFrame, ceiling)
      : ceiling;
  return {
    connect(events: OrderedConnectionEvents): OrderedConnection {
      let ended = false;
      const socket = options.socket(options.url);
      const end = (reason: string, code?: number): void => {
        if (ended) return;
        ended = true;
        if (code !== undefined) {
          try {
            socket.close(code, reason);
          } catch {
            // The socket is going already
          }
        }

        events.close(reason);
      };
      socket.addEventListener('open', () => {
        if (!ended) events.open();
      });
      socket.addEventListener('message', (event: unknown) => {
        if (ended) return;
        const data = dataOf(event);
        if (typeof data !== 'string') end('binary-frame', 1003);
        else if (data.length > max) end('frame-too-large', 1009);
        else events.message(data);
      });
      socket.addEventListener('close', (event: unknown) => end(closeNote(event)));
      socket.addEventListener('error', () => end('error'));
      return {
        send(frame: string): boolean {
          if (ended || socket.readyState !== OPEN) return false;
          socket.send(frame);
          return true;
        },
        close(): void {
          if (ended) return;
          ended = true;
          try {
            socket.close(1000, 'client');
          } catch {
            // The socket is going already
          }
        },
      };
    },
  };
}
