// The room's host events reach the client's host as they were sent, and change nothing else.

import { afterEach, describe, expect, test } from 'bun:test';
import { createOrderedCollaboration } from '../../client/create.ts';
import type { OrderedTransport } from '../../client/transport.ts';
import type { OrderedEventMessage } from '../../protocol/messages.ts';
import { closeAll, CODE, FILE, localRoom, PARTY, ROOM, settle } from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

describe('host events', () => {
  test('an event after the welcome reaches onEvent; a failing host callback stops nothing', async () => {
    const room = localRoom();
    const inner = room.transport({ person: testPerson('Ana'), ...PARTY });
    const sent = { event: 'discussion:new', comment: { anchor_id: 7, author: 'Ana' } };
    // The room's frames as they come, with a host event after the welcome
    const transport: OrderedTransport = {
      connect(events) {
        return inner.connect({
          ...events,
          message(frame) {
            events.message(frame);
            if ((JSON.parse(frame) as { type?: string }).type === 'welcome')
              events.message(JSON.stringify({ v: 3, type: 'event', event: sent }));
          },
        });
      },
    };
    const heard: OrderedEventMessage['event'][] = [];
    const handle = await createOrderedCollaboration({
      transport,
      room: ROOM,
      code: CODE,
      identity: { actorId: 'ana', name: 'Ana' },
      wordFile: async () => FILE,
      upload: room.uploader({ person: testPerson('Ana'), ...PARTY }),
      schedule: null,
      onEvent: (event) => {
        heard.push(event);
        throw new Error('the host failed');
      },
    });
    await settle();
    expect(heard).toEqual([sent]);
    expect(handle.session.status()).toBe('ready');
    handle.destroy();
  });
});
