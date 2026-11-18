// The protocol codec: one message per JSON text frame. Decoding checks the frame's length
// before it parses, then the version before anything else (so the room can answer an old or a
// new client with `reload`), then the message's shape with the change validator's readers.
// Encoding trusts the typed message and checks only the frame cap.

import { ORDERED_NAMESPACE_PATTERN, isOrderedChangeId } from '../change/ids.ts';
import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import {
  at,
  attempt,
  bool,
  fieldsOf,
  int,
  Invalid,
  isForbiddenKey,
  list,
  oneOf,
  plainText,
  required,
  str,
  type Fields,
  type OrderedChangeResult,
} from '../change/read.ts';
import type { OrderedStampedChange } from '../change/types.ts';
import { readPerson } from '../model/person.ts';
import { readClientChange, readClientChanges, readStampedChanges } from '../change/validate.ts';
import { href } from '../change/values.ts';
import type { OrderedAck, OrderedRoomRefusal, OrderedRoomRefusalCode } from '../room/room.ts';
import {
  ORDERED_PROTOCOL_LIMITS as LIMITS,
  ORDERED_PROTOCOL_VERSION,
  type OrderedClientMessage,
  type OrderedEventValue,
  type OrderedPresencePoint,
  type OrderedPresenceState,
  type OrderedServerMessage,
} from './messages.ts';

/** What the room's decoder checks a client frame against. @public */
export interface OrderedDecodeOptions {
  /** The connection's namespace, from its `hello`: change ids and presence must be in it. */
  readonly namespace?: string;
}

const ROOM_NAME = /^[\x21-\x7E]{1,256}$/;
const CODE_VERSION = /^[A-Za-z0-9.+_-]{1,64}$/;
const REASON = /^[a-z_-]{1,64}$/;
const DATA_KEY = /^[A-Za-z][A-Za-z0-9_-]{0,31}$/;

const REFUSAL_CODES: Readonly<Record<OrderedRoomRefusalCode, true>> = {
  'invalid-json': true,
  'frame-too-large': true,
  'invalid-type': true,
  'missing-field': true,
  'unknown-field': true,
  'forbidden-key': true,
  'server-field': true,
  'invalid-version': true,
  'unknown-kind': true,
  'unknown-type': true,
  'invalid-id': true,
  'foreign-id': true,
  'repeated-id': true,
  'invalid-part': true,
  'unknown-part': true,
  'invalid-space': true,
  'invalid-number': true,
  'invalid-text': true,
  'invalid-value': true,
  'invalid-date': true,
  'invalid-suggestion': true,
  'invalid-namespace': true,
  'too-long': true,
  'too-many': true,
  empty: true,
  'unknown-property': true,
  'repeated-property': true,
  'invalid-attribute': true,
  'invalid-href': true,
  'unknown-story': true,
  'unknown-address': true,
  'other-file': true,
  'unknown-suggestion': true,
  'unknown-comment': true,
  'wrong-target': true,
  'invalid-range': true,
  'policy-read-only': true,
  'policy-decide': true,
  'policy-furniture': true,
  'policy-nested': true,
  'policy-direct': true,
  'policy-not-own': true,
  'policy-comment': true,
  'policy-failed': true,
  stamp: true,
  'missing-blob': true,
  'apply-refused': true,
  'apply-failed': true,
};
const CODES = Object.keys(REFUSAL_CODES) as OrderedRoomRefusalCode[];

// ---- frames ----

/** Parse a frame and read its version and type; the rest is the caller's. */
function frameOf(
  frame: unknown,
  cap: number,
  types: readonly string[]
): { value: object; type: string } {
  if (typeof frame !== 'string') throw new Invalid('invalid-type', '', 'not a text frame');
  if (frame.length > cap) throw new Invalid('frame-too-large', '', `over ${cap}`);
  let value: unknown;
  try {
    value = JSON.parse(frame);
  } catch {
    throw new Invalid('invalid-json', '');
  }

  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Invalid('invalid-type', '', 'not an object');
  const v: unknown = Object.getOwnPropertyDescriptor(value, 'v')?.value;
  if (v !== ORDERED_PROTOCOL_VERSION) {
    throw new Invalid(
      'invalid-version',
      'v',
      Number.isSafeInteger(v) ? `version ${String(v)}` : undefined
    );
  }

  const type: unknown = Object.getOwnPropertyDescriptor(value, 'type')?.value;
  if (typeof type !== 'string') throw new Invalid('invalid-type', 'type', 'not a string');
  if (!types.includes(type)) throw new Invalid('unknown-type', 'type');
  return { value, type };
}

const fields = (value: object, keys: readonly string[]): Fields =>
  fieldsOf(value, '', new Set(['v', 'type', ...keys]));

function namespace(value: unknown, path: string, expected?: string): string {
  const name = str(value, path, 64);
  if (!ORDERED_NAMESPACE_PATTERN.test(name)) throw new Invalid('invalid-namespace', path);
  if (expected !== undefined && name !== expected) throw new Invalid('foreign-id', path);
  return name;
}

/** The room's relaxation of content-control locks: its flags, each a boolean. */
function contentLockPolicy(value: unknown, path: string): ContentLockPolicy {
  const f = fieldsOf(value, path, new Set(['allowFormattingInLockedContent']));
  return f.has('allowFormattingInLockedContent')
    ? {
        allowFormattingInLockedContent: bool(
          f.get('allowFormattingInLockedContent'),
          at(path, 'allowFormattingInLockedContent')
        ),
      }
    : {};
}

function point(value: unknown, path: string): OrderedPresencePoint {
  const f = fieldsOf(value, path, new Set(['paragraph', 'offset']));
  return {
    paragraph: plainText(required(f, 'paragraph'), at(path, 'paragraph'), 1, LIMITS.name),
    offset: int(required(f, 'offset'), at(path, 'offset')),
  };
}

function presenceData(
  value: unknown,
  path: string
): Record<string, string | number | boolean | null> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new Invalid('invalid-type', path, 'not an object');
  const keys = Object.keys(value);
  if (keys.length > LIMITS.presenceData) throw new Invalid('too-many', path);
  const f = fieldsOf(value, path, new Set(keys));
  const out: Record<string, string | number | boolean | null> = {};
  for (const key of keys.sort()) {
    // `fieldsOf` refused a forbidden key already
    if (!DATA_KEY.test(key)) {
      throw new Invalid('invalid-value', at(path, key), 'key');
    }

    const item = f.get(key);
    if (item === null || typeof item === 'boolean') out[key] = item;
    else if (typeof item === 'number') out[key] = int(item, at(path, key));
    else out[key] = plainText(item, at(path, key), 0, LIMITS.presenceValue);
  }

  return out;
}

function presenceState(value: unknown, path: string): OrderedPresenceState | null {
  if (value === null) return null;
  const f = fieldsOf(value, path, new Set(['selection', 'data']));
  const raw = required(f, 'selection');
  let selection: OrderedPresenceState['selection'] = null;
  if (raw !== null) {
    const s = fieldsOf(raw, at(path, 'selection'), new Set(['anchor', 'head']));
    selection = {
      anchor: point(required(s, 'anchor'), at(at(path, 'selection'), 'anchor')),
      head: point(required(s, 'head'), at(at(path, 'selection'), 'head')),
    };
  }

  return f.has('data')
    ? { selection, data: presenceData(f.get('data'), at(path, 'data')) }
    : { selection };
}

/** Changes in order with no gap, as the room sends them. */
function inOrder(changes: OrderedStampedChange[], path: string): OrderedStampedChange[] {
  const first = changes[0];
  if (first === undefined) {
    return changes;
  }

  for (const [k, change] of changes.entries()) {
    if (change.n !== first.n + k) {
      throw new Invalid('invalid-number', at(at(path, k), 'n'), 'not in order');
    }
  }

  return changes;
}

/** A change id field of an answer. */
function changeIdAt(value: unknown, path: string): string {
  const id = str(value, path, 96);
  if (!isOrderedChangeId(id)) {
    throw new Invalid('invalid-id', path);
  }

  return id;
}

function ack(value: unknown, path: string): OrderedAck {
  const f = fieldsOf(value, path, new Set(['id', 'n']));
  const id = changeIdAt(required(f, 'id'), at(path, 'id'));
  return { id, n: int(required(f, 'n'), at(path, 'n'), 1) };
}

function refusal(value: unknown, path: string): OrderedRoomRefusal {
  const f = fieldsOf(value, path, new Set(['id', 'code', 'reason', 'detail']));
  const rawId = required(f, 'id');
  const id = rawId === null ? null : changeIdAt(rawId, at(path, 'id'));
  const code = oneOf(required(f, 'code'), at(path, 'code'), CODES);
  const out: { id: string | null; code: OrderedRoomRefusalCode; reason?: string; detail?: string } =
    {
      id,
      code,
    };
  if (f.has('reason')) {
    const reason = str(f.get('reason'), at(path, 'reason'), 64);
    if (!REASON.test(reason)) throw new Invalid('invalid-value', at(path, 'reason'));
    out.reason = reason;
  }

  if (f.has('detail'))
    out.detail = plainText(f.get('detail'), at(path, 'detail'), 0, LIMITS.detail);
  return out;
}

// ---- decoding ----

const CLIENT_TYPES = ['hello', 'changes', 'presence', 'decide'] as const;
const SERVER_TYPES = [
  'welcome',
  'changes',
  'ack',
  'refuse',
  'saved',
  'presence',
  'reload',
  'event',
] as const;

/** A host event's JSON value, as a plain copy: no forbidden key, no deeper than the cap. */
function eventValue(value: unknown, path: string, depth: number): OrderedEventValue {
  if (value === null || typeof value === 'boolean' || typeof value === 'string') return value;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Invalid('invalid-number', path);
    return value;
  }

  if (depth >= LIMITS.eventDepth) throw new Invalid('too-many', path, 'nested too deep');
  if (Array.isArray(value))
    return value.map((item: unknown, k) => eventValue(item, at(path, k), depth + 1));
  if (typeof value !== 'object') throw new Invalid('invalid-type', path);
  return eventObject(value, path, depth + 1);
}

function eventObject(
  value: object,
  path: string,
  depth: number
): { readonly [key: string]: OrderedEventValue } {
  const out: Record<string, OrderedEventValue> = {};
  for (const [key, item] of Object.entries(value)) {
    if (isForbiddenKey(key)) throw new Invalid('forbidden-key', path);
    out[key] = eventValue(item, at(path, key), depth);
  }

  return out;
}

function readClient(frame: unknown, options: OrderedDecodeOptions): OrderedClientMessage {
  const { value, type } = frameOf(frame, LIMITS.clientFrame, CLIENT_TYPES);
  switch (type) {
    case 'hello': {
      const f = fields(value, ['room', 'file', 'last', 'code', 'namespace']);
      const room = str(required(f, 'room'), 'room', LIMITS.name);
      if (!ROOM_NAME.test(room)) throw new Invalid('invalid-value', 'room');
      const code = str(required(f, 'code'), 'code', 64);
      if (!CODE_VERSION.test(code)) throw new Invalid('invalid-value', 'code');
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        room,
        file: int(required(f, 'file'), 'file', 1),
        last: int(required(f, 'last'), 'last'),
        code,
        namespace: namespace(required(f, 'namespace'), 'namespace'),
      };
    }
    case 'changes': {
      const f = fields(value, ['changes']);
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        changes: readClientChanges(required(f, 'changes'), 'changes', options),
      };
    }
    case 'presence': {
      const f = fields(value, ['namespace', 'state']);
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        namespace: namespace(required(f, 'namespace'), 'namespace', options.namespace),
        state: presenceState(required(f, 'state'), 'state'),
      };
    }
    case 'decide': {
      const f = fields(value, ['change']);
      const change = readClientChange(required(f, 'change'), 'change', options);
      if (change.kind === 'decide') {
        return { v: ORDERED_PROTOCOL_VERSION, type, change };
      }

      throw new Invalid('invalid-value', 'change.kind', `not a ${type} change`);
    }
  }

  throw new Invalid('unknown-type', 'type');
}

function readServer(frame: unknown): OrderedServerMessage {
  const { value, type } = frameOf(frame, LIMITS.serverFrame, SERVER_TYPES);
  switch (type) {
    case 'welcome': {
      const f = fields(value, [
        'file',
        'location',
        'n',
        'saved',
        'changes',
        'blobs',
        'contentLockPolicy',
      ]);
      const n = int(required(f, 'n'), 'n');
      const saved = int(required(f, 'saved'), 'saved', 0, n);
      const changes = inOrder(
        readStampedChanges(required(f, 'changes'), 'changes', LIMITS.serverChanges),
        'changes'
      );
      const last = changes[changes.length - 1];
      if (last !== undefined && last.n > n)
        throw new Invalid('invalid-number', 'changes', 'past n');
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        file: int(required(f, 'file'), 'file', 1),
        location: href(required(f, 'location'), 'location', LIMITS.location),
        n,
        saved,
        changes,
        ...(f.has('blobs') ? { blobs: href(f.get('blobs'), 'blobs', LIMITS.location) } : {}),
        ...(f.has('contentLockPolicy')
          ? {
              contentLockPolicy: contentLockPolicy(f.get('contentLockPolicy'), 'contentLockPolicy'),
            }
          : {}),
      };
    }
    case 'changes': {
      const f = fields(value, ['changes']);
      const changes = readStampedChanges(required(f, 'changes'), 'changes', LIMITS.serverChanges);
      if (changes.length === 0) throw new Invalid('empty', 'changes');
      return { v: ORDERED_PROTOCOL_VERSION, type, changes: inOrder(changes, 'changes') };
    }
    case 'ack': {
      const f = fields(value, ['acks']);
      const acks = list(required(f, 'acks'), 'acks', 1, LIMITS.answers).map((item, k) =>
        ack(item, at('acks', k))
      );
      return { v: ORDERED_PROTOCOL_VERSION, type, acks };
    }
    case 'refuse': {
      const f = fields(value, ['refusals']);
      const refusals = list(required(f, 'refusals'), 'refusals', 1, LIMITS.answers).map((item, k) =>
        refusal(item, at('refusals', k))
      );
      return { v: ORDERED_PROTOCOL_VERSION, type, refusals };
    }
    case 'saved': {
      const f = fields(value, ['n']);
      return { v: ORDERED_PROTOCOL_VERSION, type, n: int(required(f, 'n'), 'n') };
    }
    case 'presence': {
      const f = fields(value, ['namespace', 'person', 'state']);
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        namespace: namespace(required(f, 'namespace'), 'namespace'),
        person: readPerson(required(f, 'person'), 'person'),
        state: presenceState(required(f, 'state'), 'state'),
      };
    }
    case 'event': {
      const f = fields(value, ['event']);
      const event = required(f, 'event');
      if (typeof event !== 'object' || event === null || Array.isArray(event))
        throw new Invalid('invalid-type', 'event', 'not an object');
      if (JSON.stringify(event).length > LIMITS.event) throw new Invalid('too-long', 'event');
      return { v: ORDERED_PROTOCOL_VERSION, type, event: eventObject(event, 'event', 1) };
    }
    case 'reload': {
      const f = fields(value, ['reason']);
      const reasons = ['version', 'file', 'protocol', 'closed', 'writer', 'slow'] as const;
      return {
        v: ORDERED_PROTOCOL_VERSION,
        type,
        reason: oneOf(required(f, 'reason'), 'reason', reasons),
      };
    }
  }

  throw new Invalid('unknown-type', 'type');
}

/**
 * Decode a frame a client sent, for the room. Refuses frames over
 * {@link ORDERED_PROTOCOL_LIMITS}.clientFrame, other versions, and every malformed field.
 *
 * @public
 */
export function decodeOrderedClientMessage(
  frame: string,
  options: OrderedDecodeOptions = {}
): OrderedChangeResult<OrderedClientMessage> {
  return attempt(() => readClient(frame, options));
}

/**
 * Decode a frame the room sent, for a client.
 *
 * @public
 */
export function decodeOrderedServerMessage(
  frame: string
): OrderedChangeResult<OrderedServerMessage> {
  return attempt(() => readServer(frame));
}

function encode(message: object, cap: number): string {
  const frame = JSON.stringify(message);
  if (frame.length > cap) throw new RangeError(`The message is over the frame cap of ${cap}.`);
  return frame;
}

/**
 * Encode a client's message as one frame.
 *
 * @throws RangeError when the frame is over the client frame cap: send fewer changes.
 * @public
 */
export function encodeOrderedClientMessage(message: OrderedClientMessage): string {
  return encode(message, LIMITS.clientFrame);
}

/**
 * Encode the room's message as one frame.
 *
 * @throws RangeError when the frame is over the server frame cap: split the changes.
 * @public
 */
export function encodeOrderedServerMessage(message: OrderedServerMessage): string {
  return encode(message, LIMITS.serverFrame);
}
