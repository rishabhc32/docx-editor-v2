// A random session (spec "ordered-collaboration-client", task 4.3): one room with the real
// apply and the suggestions-only policy, and clients that each edit their own copy at random
// all over the document. Frames reach the room and come back after random delays (each
// connection keeps its order). Checks, as the session runs:
//
//  - after every changes frame a client takes, its copy without its pending changes equals the
//    room's document at that order number: items, ids, suggestions, and every story composed
//    with its Word revision ids;
//  - after every round, when every frame has arrived: every copy equals the room, with no
//    pending change left; decisions of random cards give one document in every order; and the
//    saved file reopens with the same view and the same results of all decisions.

import { expect } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import {
  canonicalOoxmlFingerprint,
  serializeOoxmlPart,
} from '../../../../store/package/ooxml-tree.ts';
import { orderedCardOf } from '../../card-of.ts';
import { applyOrderedChange } from '../../change/apply.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import type { OrderedResourceOptions } from '../../compose/resources.ts';
import { orderedCards } from '../../decide.ts';
import { loadOrderedPackage } from '../../load/load.ts';
import { setStatus } from '../../model/mutations.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { takeBack, withUndoRecord } from '../../model/undo-record.ts';
import { decodeOrderedClientMessage, encodeOrderedServerMessage } from '../../protocol/codec.ts';
import type { OrderedStampedChange } from '../../change/types.ts';
import { OrderedRoom } from '../../room/room.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { normalizedBody, semanticBody } from '../matrix/compare.ts';
import { seededRandom } from '../random.ts';
import { actor } from '../room-harness.ts';
import { snapshot } from '../snapshot.ts';
import { SessionClient } from './client.ts';
import { liveOf, wordFormLimited } from './reopen.ts';
import { sessionDocument, sessionPackage } from './fixture.ts';
import { randomChange } from './random-change.ts';

/** Everything a copy holds, as one string: its model and its composed stories. */
export function printOf(doc: OrderedDocument): string {
  return JSON.stringify({
    model: snapshot(doc, false),
    paraIds: [...doc.paraIds].sort(),
    parts: [...doc.stories.keys()].map((name) =>
      serializeOoxmlPart(composeOrderedPart(doc, name).part)
    ),
  });
}

type Status = 'accepted' | 'rejected';

/** Every story after decisions, in an order; the document is restored. */
function decided(doc: OrderedDocument, order: readonly (readonly [string, Status])[]): string[] {
  const { record } = withUndoRecord(doc, () => {
    for (const [id, status] of order)
      if (doc.suggestions.get(id)?.status === 'pending') setStatus(doc, id, status);
  });
  const parts = [...doc.stories.keys()].map((name) =>
    semanticBody(normalizedBody(composeOrderedPart(doc, name).part, (id) => id))
  );
  takeBack(doc, record);
  return parts;
}

function shuffled<T>(list: readonly T[], rand: () => number): T[] {
  const out = [...list];
  for (let k = out.length - 1; k > 0; k -= 1) {
    const j = Math.floor(rand() * (k + 1));
    [out[k], out[j]] = [out[j] as T, out[k] as T];
  }
  return out;
}

/** Decisions of random cards, in three random orders, give one document. */
function checkDecisionOrders(doc: OrderedDocument, rand: () => number, label: string): void {
  const cards = shuffled(orderedCards(doc), rand).slice(0, 5);
  if (cards.length < 2) return;
  const choice = cards.flatMap((card) => {
    const status: Status = rand() < 0.5 ? 'accepted' : 'rejected';
    return card.suggestions.map((id) => [id, status] as const);
  });
  const first = decided(doc, choice);
  for (let k = 0; k < 2; k += 1)
    expect({ label, order: k, parts: decided(doc, shuffled(choice, rand)) }).toEqual({
      label,
      order: k,
      parts: first,
    });
}

/**
 * The saved file reopens with the same view; and every suggestion decided at once, both ways,
 * gives the same, but for the suggestions a documented limit of Word's form covers, which stay
 * pending on both sides (`reopen.ts`). Returns how many suggestions the limits kept out. The
 * save reads media through `options` (the room's blobs).
 */
export function checkSaveAndLoad(
  doc: OrderedDocument,
  label: string,
  base: OoxmlPackage = sessionPackage(),
  options: OrderedResourceOptions = {}
): number {
  const read = readOoxmlPackage(orderedPackageBytes(doc, base, options), {});
  if (!read.ok) throw new Error(`${label}: the saved file does not open: ${read.reason}`);
  const again = loadOrderedPackage(read.package);
  const fingerprints = (d: OrderedDocument) =>
    [...d.stories.keys()].map((name) =>
      canonicalOoxmlFingerprint(composeOrderedPart(d, name).part)
    );
  expect({ label, view: fingerprints(again) }).toEqual({ label, view: fingerprints(doc) });
  const limited = wordFormLimited(doc);
  const live = liveOf(doc, again);
  const pending = (d: OrderedDocument) =>
    [...d.suggestions.values()].filter((s) => s.status === 'pending').map((s) => s.id);
  // Every reopened suggestion was written by one live suggestion
  expect({ label, unmatched: pending(again).filter((id) => !live.has(id)) }).toEqual({
    label,
    unmatched: [],
  });
  const liveIds = pending(doc).filter((id) => !limited.has(id));
  const againIds = pending(again).filter((id) => !limited.has(live.get(id) ?? ''));
  for (const status of ['accepted', 'rejected'] as const)
    expect({
      label,
      status,
      parts: decided(
        again,
        againIds.map((id) => [id, status] as const)
      ),
    }).toEqual({
      label,
      status,
      parts: decided(
        doc,
        liveIds.map((id) => [id, status] as const)
      ),
    });
  return limited.size;
}

/** What a session did. */
export interface SessionResult {
  readonly order: number;
  readonly changes: number;
  readonly refused: {
    readonly invalid: number;
    readonly local: number;
    readonly reapply: number;
    readonly room: number;
  };
  readonly decided: number;
  readonly frames: number;
  /** Suggestions a documented Word-form limit kept out of the save check, over the rounds. */
  readonly limited: number;
  /** The room's log, and its length at the end of each round. */
  readonly log: readonly OrderedStampedChange[];
  readonly rounds: readonly number[];
}

/**
 * Run a session: `people[0]` is the organiser, the others are parties. `steps` random steps in
 * `rounds` rounds; every round ends with every frame delivered and every check.
 */
export function runSession(
  seed: number,
  people: readonly string[],
  steps: number,
  rounds: number,
  inspect?: (doc: OrderedDocument, round: number) => void,
  /** Group 9.10's changes too: a paste, a new list, a paragraph removal, protection. */
  paste = false
): SessionResult {
  const rand = seededRandom(seed);
  let tick = 0;
  const room = new OrderedRoom({
    document: sessionDocument(),
    applyChange: applyOrderedChange,
    policy: suggestionsOnlyPolicy(),
    now: () => Date.UTC(2026, 9, 28) + 1000 * tick++,
  });
  const prints = new Map<number, string>([[0, printOf(room.document)]]);
  const clients = people.map((name, k) => {
    const who = actor(name, `${name.toLowerCase()}-${k}`, k === 0 ? 'organiser' : 'party');
    return new SessionClient(sessionDocument(), who, (client) => {
      expect({
        client: name,
        n: client.last,
        same: printOf(client.doc) === prints.get(client.last),
      }).toEqual({
        client: name,
        n: client.last,
        same: true,
      });
    });
  });
  const toRoom: string[][] = clients.map(() => []);
  const toClient: string[][] = clients.map(() => []);
  let frames = 0;
  let changes = 0;
  let decidedCount = 0;
  let limitedTotal = 0;
  const ends: number[] = [];

  const reachRoom = (k: number) => {
    const frame = toRoom[k]?.shift();
    const client = clients[k];
    if (frame === undefined || !client) return;
    const decoded = decodeOrderedClientMessage(frame, { namespace: client.actor.namespace });
    if (!decoded.ok || decoded.value.type !== 'changes')
      throw new Error('the room cannot read a frame');
    const result = room.submit(client.actor, decoded.value.changes);
    frames += 1;
    if (result.accepted.length > 0) {
      prints.set(room.order, printOf(room.document));
      decidedCount += result.accepted.filter((c) => c.kind === 'decide').length;
      const out = encodeOrderedServerMessage({ v: 3, type: 'changes', changes: result.accepted });
      for (const queue of toClient) queue.push(out);
    }
    if (result.refused.length > 0)
      toClient[k]?.push(
        encodeOrderedServerMessage({ v: 3, type: 'refuse', refusals: result.refused })
      );
  };
  const reachClient = (k: number) => {
    const frame = toClient[k]?.shift();
    if (frame !== undefined) clients[k]?.receive(frame);
  };
  const drain = () => {
    for (let pass = 0; pass < 50; pass += 1) {
      clients.forEach((client, k) => {
        const frame = client.flush();
        if (frame !== null) toRoom[k]?.push(frame);
      });
      if (toRoom.every((q) => q.length === 0) && toClient.every((q) => q.length === 0)) return;
      clients.forEach((_, k) => {
        while ((toRoom[k]?.length ?? 0) > 0) reachRoom(k);
      });
      clients.forEach((_, k) => {
        while ((toClient[k]?.length ?? 0) > 0) reachClient(k);
      });
    }
    throw new Error('the session does not settle');
  };

  const perRound = Math.ceil(steps / rounds);
  for (let round = 0; round < rounds; round += 1) {
    for (let step = 0; step < perRound; step += 1) {
      const k = Math.floor(rand() * clients.length);
      const client = clients[k];
      if (!client) continue;
      const roll = rand();
      if (roll < 0.5) {
        const who = { author: client.actor.person.name, organiser: !client.actor.suggests };
        const body = randomChange(client.doc, who, rand, paste);
        if (body && client.local(body)) changes += 1;
      } else if (roll < 0.65) {
        const frame = client.flush();
        if (frame !== null) toRoom[k]?.push(frame);
      } else if (roll < 0.82) reachRoom(k);
      else reachClient(k);
    }
    drain();
    inspect?.(room.document, round);
    const label = `seed ${seed} round ${round}`;
    const expected = printOf(room.document);
    for (const client of clients) {
      expect({ label, client: client.actor.person.name, pending: client.pending.length }).toEqual({
        label,
        client: client.actor.person.name,
        pending: 0,
      });
      expect({
        label,
        client: client.actor.person.name,
        same: printOf(client.doc) === expected,
      }).toEqual({
        label,
        client: client.actor.person.name,
        same: true,
      });
    }
    // The card each pending suggestion is in, found from it alone, is the whole list's
    for (const card of orderedCards(room.document))
      for (const id of card.suggestions)
        expect(orderedCardOf(room.document, id)?.suggestions).toEqual(card.suggestions);
    checkDecisionOrders(room.document, rand, label);
    limitedTotal += checkSaveAndLoad(room.document, label);
    ends.push(room.order);
  }
  const refused = { invalid: 0, local: 0, reapply: 0, room: 0 };
  for (const client of clients) {
    refused.invalid += client.counts.invalid;
    refused.local += client.counts.local;
    refused.reapply += client.counts.reapply;
    refused.room += client.counts.room;
  }
  return {
    order: room.order,
    changes,
    refused,
    decided: decidedCount,
    frames,
    limited: limitedTotal,
    log: room.joinSnapshot(0),
    rounds: ends,
  };
}
