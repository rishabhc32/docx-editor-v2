// The WebSocket transport over real sockets: a server that speaks the protocol (the in-process
// room behind Bun's WebSocket server) and the standard `WebSocket` client the host passes in.
// Two editors join through it, type at once and converge; the server's drop is a reconnect; a
// binary frame ends the connection by name.

import { afterEach, describe, expect, test } from 'bun:test';
import type { Server } from 'bun';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import type { OrderedConnection } from '../../client/transport.ts';
import { orderedWebSocketTransport } from '../../client/transport.ts';
import {
  caretAt,
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

interface Socket {
  connection: OrderedConnection | null;
}

const servers: Server<Socket>[] = [];
afterEach(() => {
  closeAll();
  for (const server of servers.splice(0)) void server.stop(true);
});

/** A WebSocket server in front of a local room: one room connection for each socket. */
function serve(room: OrderedLocalRoom): Server<Socket> {
  const server = Bun.serve<Socket>({
    port: 0,
    fetch(request, s) {
      const author = new URL(request.url).searchParams.get('author') ?? 'nobody';
      const upgraded = s.upgrade(request, { data: { connection: null } });
      if (!upgraded) return new Response('upgrade', { status: 400 });
      (pendingAuthors as string[]).push(author);
      return undefined;
    },
    websocket: {
      open(ws) {
        const author = pendingAuthors.shift() ?? 'nobody';
        ws.data.connection = room.transport({ person: testPerson(author), ...PARTY }).connect({
          open: () => {},
          message: (frame) => ws.send(frame),
          close: () => ws.close(),
        });
      },
      message(ws, message) {
        ws.data.connection?.send(typeof message === 'string' ? message : '');
      },
      close(ws) {
        ws.data.connection?.close();
      },
    },
  });
  servers.push(server);
  return server;
}
const pendingAuthors: string[] = [];

const overSockets = (server: Server<Socket>, author: string) =>
  orderedWebSocketTransport({
    url: `ws://localhost:${server.port}/room?author=${author}`,
    socket: (url) => new WebSocket(url),
  });

describe('the WebSocket transport', () => {
  test('two editors join through real sockets, type at once, and converge', async () => {
    const room = localRoom();
    const server = serve(room);
    const ana = await join(room, 'Ana', { transport: overSockets(server, 'Ana') });
    const ben = await join(room, 'Ben', { transport: overSockets(server, 'Ben') });
    caretAt(ana, 0, 4);
    caretAt(ben, 1, 0);
    surface(ana).insertPlainText('new ');
    surface(ben).insertPlainText('Rent ');
    for (let k = 0; k < 20 && room.room.order < 2; k += 1) await settle();
    await settle();
    for (const person of [ana, ben]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
    }
    expect(editorTexts(ben).slice(0, 2)).toEqual([
      'The new fee is 100 USD.',
      'Rent Due in 30 days.',
    ]);
  });

  test('a dropped socket ends the connection; the session reconnects and sends what it holds', async () => {
    const room = localRoom();
    const server = serve(room);
    const ana = await join(room, 'Ana', { transport: overSockets(server, 'Ana') });
    room.disconnect();
    for (let k = 0; k < 20 && ana.handle.session.status() !== 'disconnected'; k += 1)
      await settle();
    expect(ana.handle.session.status()).toBe('disconnected');
    caretAt(ana, 2, 0);
    surface(ana).insertPlainText('Now ');
    await settle();
    ana.handle.session.reconnect();
    for (let k = 0; k < 20 && room.room.order < 1; k += 1) await settle();
    await settle();
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(ana.handle.session.status()).toBe('ready');
  });

  test('a binary frame ends the connection by name', async () => {
    const server = Bun.serve({
      port: 0,
      fetch: (request, s) =>
        s.upgrade(request) ? undefined : new Response('upgrade', { status: 400 }),
      websocket: {
        open(ws) {
          ws.send(new Uint8Array([1, 2, 3]));
        },
        message() {},
      },
    });
    try {
      const reasons: string[] = [];
      const transport = orderedWebSocketTransport({
        url: `ws://localhost:${server.port}`,
        socket: (url) => new WebSocket(url),
      });
      transport.connect({
        open: () => {},
        message: () => {},
        close: (reason) => reasons.push(reason),
      });
      for (let k = 0; k < 40 && reasons.length === 0; k += 1) await settle(1);
      expect(reasons).toEqual(['binary-frame']);
    } finally {
      void server.stop(true);
    }
  });
});
