// The protocol codec: every message survives a round trip, frames over the cap and other
// versions are refused before anything else, malformed messages are refused by code, and random
// frames never throw.

import { describe, expect, test } from 'bun:test';
import { ORDERED_CHANGE_LIMITS, type OrderedStampedChange } from '../change/types.ts';
import {
  decodeOrderedClientMessage,
  decodeOrderedServerMessage,
  encodeOrderedClientMessage,
  encodeOrderedServerMessage,
} from '../protocol/codec.ts';
import {
  ORDERED_PROTOCOL_LIMITS,
  type OrderedClientMessage,
  type OrderedServerMessage,
} from '../protocol/messages.ts';
import { seededRandom } from './random.ts';
import { testPerson } from './test-persons.ts';

const SPACE = 'b:/word/document.xml#0.0.0';
const PART = '/word/document.xml';
const INSERT = {
  v: 1,
  id: 'ana-1:1',
  kind: 'insert',
  part: PART,
  mode: 'suggest',
  place: { space: SPACE, offset: 3 },
  text: 'fee',
} as const;
const STAMPED = (n: number): OrderedStampedChange => ({
  ...INSERT,
  id: `ana-1:${n}`,
  n,
  by: testPerson('Ana'),
  at: '2026-10-28T09:30:00Z',
});
const BEN = testPerson('Ben');
const STATE = {
  selection: { anchor: { paragraph: 'p1', offset: 0 }, head: { paragraph: 'p1', offset: 4 } },
  data: { seat: 'buyer', color: null, active: true, rank: 2 },
};

const CLIENT: readonly OrderedClientMessage[] = [
  {
    v: 3,
    type: 'hello',
    room: 'contract-42',
    file: 3,
    last: 57,
    code: '2.19.0',
    namespace: 'ana-1',
  },
  {
    v: 3,
    type: 'changes',
    changes: [
      INSERT,
      { v: 1, id: 'ana-1:2', kind: 'split', part: PART, mode: 'edit', place: INSERT.place },
    ],
  },
  { v: 3, type: 'presence', namespace: 'ana-1', state: STATE },
  { v: 3, type: 'presence', namespace: 'ana-1', state: null },
  {
    v: 3,
    type: 'decide',
    change: {
      v: 1,
      id: 'ana-1:3',
      kind: 'decide',
      cards: [{ suggestions: [{ suggestion: 'S2', version: 3 }] }],
      status: 'accepted',
    },
  },
  {
    v: 3,
    type: 'changes',
    changes: [
      {
        v: 1,
        id: 'ana-1:4',
        kind: 'comment',
        part: PART,
        from: { space: SPACE, offset: 0 },
        to: { mark: { space: 'm:/word/document.xml#0.0.0', offset: 0 } },
        text: 'Why?\nSee clause 2.',
      },
      { v: 1, id: 'ana-1:5', kind: 'comment-reply', comment: '1000000001', text: 'Agreed.' },
      { v: 1, id: 'ana-1:6', kind: 'comment-resolve', comment: '1000000001', note: 'Done.' },
      { v: 1, id: 'ana-1:7', kind: 'comment-resolve', comment: '7' },
    ],
  },
];

const SERVER: readonly OrderedServerMessage[] = [
  {
    v: 3,
    type: 'welcome',
    file: 3,
    location: 'https://files.example.com/rooms/42/file-3.docx',
    n: 59,
    saved: 57,
    changes: [STAMPED(58), STAMPED(59)],
  },
  {
    v: 3,
    type: 'welcome',
    file: 1,
    location: '/rooms/42/file-1.docx',
    n: 0,
    saved: 0,
    changes: [],
    contentLockPolicy: { allowFormattingInLockedContent: true },
  },
  { v: 3, type: 'changes', changes: [STAMPED(60)] },
  { v: 3, type: 'ack', acks: [{ id: 'ana-1:1', n: 60 }] },
  {
    v: 3,
    type: 'refuse',
    refusals: [
      { id: 'ana-1:2', code: 'policy-nested', reason: 'nested' },
      { id: 'ana-1:3', code: 'apply-refused', reason: 'not-text', detail: 'a split of a picture' },
      { id: null, code: 'too-many' },
    ],
  },
  { v: 3, type: 'saved', n: 60 },
  { v: 3, type: 'presence', namespace: 'ben-2', person: testPerson('Ben'), state: STATE },
  { v: 3, type: 'reload', reason: 'version' },
  { v: 3, type: 'reload', reason: 'slow' },
  {
    v: 3,
    type: 'event',
    event: {
      event: 'discussion:new',
      comment: { anchor_id: 7, author: 'Ana', tags: ['a', null, 1.5] },
    },
  },
];

/** A list nested `n` deep. */
const deep = (n: number): unknown => (n === 0 ? 1 : [deep(n - 1)]);

const clientCode = (frame: string, namespace?: string): string => {
  const result = decodeOrderedClientMessage(frame, namespace === undefined ? {} : { namespace });
  return result.ok ? 'ok' : result.error.code;
};
const serverCode = (frame: string): string => {
  const result = decodeOrderedServerMessage(frame);
  return result.ok ? 'ok' : result.error.code;
};

describe('the protocol codec', () => {
  test('every message survives a round trip', () => {
    for (const message of CLIENT) {
      const decoded = decodeOrderedClientMessage(encodeOrderedClientMessage(message), {
        namespace: 'ana-1',
      });
      expect(decoded).toEqual({ ok: true, value: message });
    }
    for (const message of SERVER) {
      const decoded = decodeOrderedServerMessage(encodeOrderedServerMessage(message));
      expect(decoded).toEqual({ ok: true, value: message });
    }
  });

  test('a frame over the cap is refused before it is read; encoding one throws', () => {
    expect(clientCode('x'.repeat(ORDERED_PROTOCOL_LIMITS.clientFrame + 1))).toBe('frame-too-large');
    expect(serverCode(' '.repeat(ORDERED_PROTOCOL_LIMITS.serverFrame + 1))).toBe('frame-too-large');
    const big = Array.from({ length: 20 }, (_, k) => ({
      ...INSERT,
      id: `ana-1:${k + 1}`,
      text: 'a'.repeat(60_000),
    }));
    expect(() => encodeOrderedClientMessage({ v: 3, type: 'changes', changes: big })).toThrow(
      RangeError
    );
  });

  test('an Accept All of the most suggestions one decision names fits one client frame', () => {
    // Minted suggestion ids are at most 80 units: `S:`, a 64-unit namespace, `:`, the change
    // counter, `.` and the suggestion's number. Each card holds one suggestion, the longest form
    const namespace = 'n'.repeat(64);
    const cards = Array.from({ length: ORDERED_CHANGE_LIMITS.decisions }, (_, k) => {
      const suggestion = `S:${namespace}:${String(k + 1).padStart(11, '0')}.1`;
      return { suggestions: [{ suggestion, version: 99_999 }] };
    });
    expect(cards[0]?.suggestions[0]?.suggestion).toHaveLength(80);
    const message: OrderedClientMessage = {
      v: 3,
      type: 'decide',
      change: { v: 1, id: 'olga-1:1', kind: 'decide', cards, status: 'accepted' },
    };
    const frame = encodeOrderedClientMessage(message);
    expect(frame.length).toBeLessThan(ORDERED_PROTOCOL_LIMITS.clientFrame);
    const decoded = decodeOrderedClientMessage(frame, { namespace: 'olga-1' });
    expect(decoded.ok && decoded.value).toEqual(message);
  });

  test('another version is refused first, with the version it named', () => {
    const hello = {
      type: 'hello',
      room: 'r',
      file: 1,
      last: 0,
      code: '3.0.0',
      namespace: 'ana-1',
      more: true,
    };
    for (const v of [1, 2, 4, 0, '3', null]) {
      const result = decodeOrderedClientMessage(JSON.stringify({ v, ...hello }));
      expect(result.ok ? 'ok' : result.error.code).toBe('invalid-version');
    }

    // A page of the protocol before media uploads is refused by name
    const one = decodeOrderedClientMessage(JSON.stringify({ v: 1, ...hello }));
    expect(one.ok ? '' : one.error.detail).toBe('version 1');
    expect(serverCode(JSON.stringify({ type: 'saved', n: 1 }))).toBe('invalid-version');
  });

  test('malformed frames are refused by code', () => {
    const client: [unknown, string, string?][] = [
      ['{', 'invalid-json'],
      ['[]', 'invalid-type'],
      ['"hello"', 'invalid-type'],
      [{ v: 3, type: 'bye' }, 'unknown-type'],
      [
        { v: 3, type: 'hello', room: 'r', file: 0, last: 0, code: '1', namespace: 'ana-1' },
        'invalid-number',
      ],
      [
        { v: 3, type: 'hello', room: 'r r', file: 1, last: 0, code: '1', namespace: 'ana-1' },
        'invalid-value',
      ],
      [
        { v: 3, type: 'hello', room: 'r', file: 1, last: 0, code: '1', namespace: 'a b' },
        'invalid-namespace',
      ],
      [{ v: 3, type: 'changes', changes: [INSERT], n: 3 }, 'unknown-field'],
      [{ v: 3, type: 'changes', changes: [] }, 'empty'],
      [{ v: 3, type: 'changes', changes: [{ ...INSERT, by: 'Ben' }] }, 'server-field'],
      [{ v: 3, type: 'changes', changes: [{ ...INSERT, id: 'ben-1:1' }] }, 'foreign-id', 'ana-1'],
      [{ v: 3, type: 'presence', namespace: 'ben-1', state: null }, 'foreign-id', 'ana-1'],
      [
        {
          v: 3,
          type: 'presence',
          namespace: 'ana-1',
          state: { selection: null, data: { a: { b: 1 } } },
        },
        'invalid-type',
      ],
      [
        {
          v: 3,
          type: 'presence',
          namespace: 'ana-1',
          state: {
            selection: {
              anchor: { paragraph: 'p', offset: -1 },
              head: { paragraph: 'p', offset: 0 },
            },
          },
        },
        'invalid-number',
      ],
      [{ v: 3, type: 'decide', change: INSERT }, 'invalid-value'],
      [
        {
          v: 3,
          type: 'marker',
          change: {
            v: 1,
            id: 'ana-1:1',
            kind: 'decide',
            cards: [{ suggestions: [{ suggestion: 'S1', version: 0 }] }],
            status: 'accepted',
          },
        },
        'unknown-type',
      ],
    ];
    for (const [value, code, namespace] of client) {
      const frame = typeof value === 'string' ? value : JSON.stringify(value);
      expect({ value, code: clientCode(frame, namespace) }).toEqual({ value, code });
    }
    // A JSON key, not an object literal's prototype
    const protoFrame =
      '{"v":3,"type":"presence","namespace":"ana-1","state":{"selection":null,"data":{"__proto__":1}}}';
    expect(clientCode(protoFrame)).toBe('forbidden-key');
    expect(clientCode('{"v":3,"type":"hello","__proto__":{},"room":"r"}')).toBe('forbidden-key');
    const WELCOME = {
      v: 3,
      type: 'welcome',
      file: 1,
      location: '/f.docx',
      n: 0,
      saved: 0,
      changes: [],
    };
    const server: [unknown, string][] = [
      [
        {
          v: 3,
          type: 'welcome',
          file: 1,
          location: 'javascript:alert(1)',
          n: 0,
          saved: 0,
          changes: [],
        },
        'invalid-href',
      ],
      [
        { v: 3, type: 'welcome', file: 1, location: '/f.docx', n: 1, saved: 2, changes: [] },
        'invalid-number',
      ],
      [
        {
          v: 3,
          type: 'welcome',
          file: 1,
          location: '/f.docx',
          n: 5,
          saved: 0,
          changes: [STAMPED(2), STAMPED(4)],
        },
        'invalid-number',
      ],
      [
        {
          v: 3,
          type: 'welcome',
          file: 1,
          location: '/f.docx',
          n: 1,
          saved: 0,
          changes: [STAMPED(2)],
        },
        'invalid-number',
      ],
      [
        { ...WELCOME, contentLockPolicy: { allowFormattingInLockedContent: 'yes' } },
        'invalid-type',
      ],
      [{ ...WELCOME, contentLockPolicy: { allowEditingInLockedContent: true } }, 'unknown-field'],
      [{ v: 3, type: 'changes', changes: [INSERT] }, 'missing-field'],
      [{ v: 3, type: 'changes', changes: [] }, 'empty'],
      [{ v: 3, type: 'ack', acks: [{ id: 'nope', n: 1 }] }, 'invalid-id'],
      [{ v: 3, type: 'refuse', refusals: [{ id: null, code: 'made-up' }] }, 'invalid-value'],
      [
        { v: 3, type: 'refuse', refusals: [{ id: null, code: 'stamp', reason: 'Not A Reason' }] },
        'invalid-value',
      ],
      [{ v: 3, type: 'presence', namespace: 'ben-2', author: 'Ben', state: null }, 'unknown-field'],
      [
        { v: 3, type: 'presence', namespace: 'ben-2', person: { ...BEN, id: '' }, state: null },
        'empty',
      ],
      [
        {
          v: 3,
          type: 'presence',
          namespace: 'ben-2',
          person: { ...BEN, org: 'Acme ' },
          state: null,
        },
        'invalid-value',
      ],
      [{ v: 3, type: 'reload', reason: 'tired' }, 'invalid-value'],
      [{ v: 3, type: 'event', event: 'room:frozen' }, 'invalid-type'],
      [{ v: 3, type: 'event', event: [1] }, 'invalid-type'],
      [{ v: 3, type: 'event', event: { a: 'x'.repeat(70_000) } }, 'too-long'],
      [{ v: 3, type: 'event', event: { a: deep(20) } }, 'too-many'],
    ];
    for (const [value, code] of server)
      expect({ value, code: serverCode(JSON.stringify(value)) }).toEqual({ value, code });
    // A key that would reach an object's prototype, at any depth
    expect(serverCode('{"v":3,"type":"event","event":{"a":{"__proto__":{"x":1}}}}')).toBe(
      'forbidden-key'
    );
  });

  test('random frames are refused by code, never thrown', () => {
    const rand = seededRandom(8184);
    const frames = [
      ...CLIENT.map(encodeOrderedClientMessage),
      ...SERVER.map(encodeOrderedServerMessage),
    ];
    const alphabet = '{}[]":,0123456789abcdefvtype\\u_ -';
    let refused = 0;
    for (let k = 0; k < 3000; k += 1) {
      const base = frames[k % frames.length] ?? '{}';
      let frame: string;
      if (k % 3 === 0) {
        frame = Array.from({ length: Math.floor(rand() * 80) }, () =>
          alphabet.charAt(Math.floor(rand() * alphabet.length))
        ).join('');
      } else {
        const at = Math.floor(rand() * base.length);
        const cut = Math.floor(rand() * 6);
        const insert = alphabet.charAt(Math.floor(rand() * alphabet.length));
        frame = base.slice(0, at) + (k % 3 === 1 ? insert : '') + base.slice(at + cut);
      }
      for (const result of [
        decodeOrderedClientMessage(frame, { namespace: 'ana-1' }),
        decodeOrderedServerMessage(frame),
      ])
        if (!result.ok) refused += 1;
    }
    expect(refused).toBeGreaterThan(5000);
  });
});
