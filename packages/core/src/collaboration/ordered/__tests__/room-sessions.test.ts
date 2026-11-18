// Random sessions against one room: clients edit their own copies at once, frames reach the
// room and come back in random interleavings (each connection keeps its own order), and the
// room refuses what the policy refuses. At the end every copy equals the room's document, and
// every client received the same changes in the same order. Seeds are fixed.

import { describe, expect, test } from 'bun:test';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { isHidden } from '../model/document.ts';
import type { OrderedDocument } from '../model/types.ts';
import { decodeOrderedClientMessage, encodeOrderedServerMessage } from '../protocol/codec.ts';
import { OrderedRoom } from '../room/room.ts';
import { suggestionsOnlyPolicy } from '../room/suggestions-only.ts';
import type { OrderedChangePlace } from '../change/types.ts';
import { seededRandom } from './random.ts';
import { TestClient, type ChangeBody } from './room-client.ts';
import { actor, BODY, bodyFingerprint, loadRoom, snapshot, realApply } from './room-harness.ts';

type Role = 'party' | 'organiser';

function shown(doc: OrderedDocument): { units: OrderedChangePlace[]; marks: OrderedChangePlace[] } {
  const units: OrderedChangePlace[] = [];
  const marks: OrderedChangePlace[] = [];
  for (const item of doc.stories.get(BODY)?.tree.items() ?? []) {
    if (isHidden(doc, item)) continue;
    if (item.kind === 'text')
      for (let k = 0; k < item.length; k += 1)
        units.push({ space: item.space, offset: item.offset + k });
    if (item.kind === 'mark') {
      units.push({ space: item.space, offset: item.offset });
      marks.push({ mark: { space: item.space, offset: item.offset } });
    }
  }
  return { units, marks };
}

function randomEdit(client: TestClient, rand: () => number): ChangeBody | null {
  const doc = client.doc;
  const pick = <T>(list: readonly T[]): T | undefined => list[Math.floor(rand() * list.length)];
  const { units, marks } = shown(doc);
  const own = [...doc.suggestions.values()].filter(
    (s) => s.status === 'pending' && s.attribution.person?.id === client.actor.person.id
  );
  const pending = [...doc.suggestions.values()].filter((s) => s.status === 'pending');
  const organiser = !client.actor.suggests;
  const mode = organiser && rand() < 0.3 ? 'edit' : 'suggest';
  const into = mode === 'suggest' && rand() < 0.2 ? pick(own)?.id : undefined;
  const edit = { part: BODY, mode, ...(into === undefined ? {} : { into }) } as const;
  const roll = rand();
  if (roll < 0.35) {
    const place = rand() < 0.1 ? pick(marks) : pick(units);
    if (!place) return null;
    const text = 'abxyz '.charAt(Math.floor(rand() * 6)).repeat(1 + Math.floor(rand() * 3));
    return { kind: 'insert', ...edit, place, text };
  }
  if (roll < 0.7) {
    const i = Math.floor(rand() * units.length);
    const j = i + 1 + Math.floor(rand() * 6);
    const from = units[i];
    const to = units[Math.min(j, units.length - 1)];
    if (!from || !to || j >= units.length) return null;
    if (roll < 0.55) return { kind: 'delete', ...edit, from, to };
    return { kind: 'format', ...edit, mode: 'suggest', from, to, properties: [{ localName: 'b' }] };
  }
  if (roll < 0.8) {
    const place = pick(units);
    return place ? { kind: 'split', ...edit, place } : null;
  }
  if (roll < 0.9) {
    const suggestion = pick(own);
    return suggestion ? { kind: 'retract', suggestions: [suggestion.id] } : null;
  }
  const suggestion = pick(pending);
  if (!suggestion || !organiser) return null;
  const seen = orderedCardSeen(doc, suggestion.id);
  if (!seen) return null;
  return {
    kind: 'decide',
    cards: [{ suggestions: seen }],
    status: rand() < 0.5 ? 'accepted' : 'rejected',
  };
}

function session(seed: number, roles: readonly (readonly [string, Role])[], steps: number) {
  const rand = seededRandom(seed);
  let tick = 0;
  const room = new OrderedRoom({
    document: loadRoom(),
    applyChange: realApply,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
  });
  const clients = roles.map(
    ([author, role], k) =>
      new TestClient(loadRoom(), actor(author, `${author.toLowerCase()}-${k}`, role))
  );
  const toRoom: string[][] = clients.map(() => []);
  const toClient: string[][] = clients.map(() => []);
  let refused = 0;

  const reachRoom = (k: number) => {
    const frame = toRoom[k]?.shift();
    const client = clients[k];
    if (frame === undefined || !client) return;
    const decoded = decodeOrderedClientMessage(frame, { namespace: client.actor.namespace });
    if (!decoded.ok || decoded.value.type !== 'changes')
      throw new Error('the room cannot read a frame');
    const result = room.submit(client.actor, decoded.value.changes);
    refused += result.refused.length;
    if (result.accepted.length > 0) {
      const frameOut = encodeOrderedServerMessage({
        v: 3,
        type: 'changes',
        changes: result.accepted,
      });
      for (const queue of toClient) queue.push(frameOut);
    }
    if (result.acks.length > 0)
      toClient[k]?.push(encodeOrderedServerMessage({ v: 3, type: 'ack', acks: result.acks }));
    if (result.refused.length > 0)
      toClient[k]?.push(
        encodeOrderedServerMessage({ v: 3, type: 'refuse', refusals: result.refused })
      );
  };
  const reachClient = (k: number) => {
    const frame = toClient[k]?.shift();
    if (frame !== undefined) clients[k]?.receive(frame);
  };

  for (let step = 0; step < steps; step += 1) {
    const k = Math.floor(rand() * clients.length);
    const client = clients[k];
    if (!client) continue;
    const roll = rand();
    if (roll < 0.45) {
      const body = randomEdit(client, rand);
      if (body) client.local(body);
    } else if (roll < 0.6) {
      const frame = client.flush();
      if (frame !== null) toRoom[k]?.push(frame);
    } else if (roll < 0.8) reachRoom(k);
    else reachClient(k);
  }
  for (let round = 0; round < 100; round += 1) {
    clients.forEach((client, k) => {
      const frame = client.flush();
      if (frame !== null) toRoom[k]?.push(frame);
    });
    if (toRoom.every((q) => q.length === 0) && toClient.every((q) => q.length === 0)) break;
    clients.forEach((_, k) => {
      while ((toRoom[k]?.length ?? 0) > 0) reachRoom(k);
    });
    clients.forEach((_, k) => {
      while ((toClient[k]?.length ?? 0) > 0) reachClient(k);
    });
  }
  return { room, clients, refused };
}

describe('random sessions against one room', () => {
  const casts = {
    'two people': [
      ['Ana', 'organiser'],
      ['Ben', 'party'],
    ],
    'four people': [
      ['Olga', 'organiser'],
      ['Pia', 'party'],
      ['Quin', 'party'],
      ['Ravi', 'party'],
    ],
  } as const;
  for (const [name, roles] of Object.entries(casts)) {
    for (const seed of [11, 23, 37]) {
      test(`${name}, seed ${seed}: every copy ends equal to the room, in one order`, () => {
        const { room, clients, refused } = session(seed, roles, roles.length * 120);
        expect(room.order).toBeGreaterThan(20);
        expect(refused).toBeGreaterThan(0);
        const expected = snapshot(room.document);
        const fingerprint = bodyFingerprint(room.document);
        const order = Array.from({ length: room.order }, (_, k) => k + 1);
        for (const client of clients) {
          expect(client.pending).toEqual([]);
          expect(client.last).toBe(room.order);
          expect(client.received).toEqual(order);
          expect(snapshot(client.doc)).toEqual(expected);
          expect(bodyFingerprint(client.doc)).toBe(fingerprint);
        }
      });
    }
  }
});
