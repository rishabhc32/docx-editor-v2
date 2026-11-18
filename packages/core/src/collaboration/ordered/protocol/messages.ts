// The protocol's messages (design D8): JSON text frames, one message per frame, each with the
// protocol version `v` and a `type`. The codec (`codec.ts`) checks every frame against these
// shapes and caps; the transport (WebSocket, or in-process for tests) only moves strings.

import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import type { OrderedDecideChange, OrderedChange, OrderedStampedChange } from '../change/types.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedAck, OrderedRoomRefusal } from '../room/room.ts';

/** The protocol version every message carries in `v`. @public */
export const ORDERED_PROTOCOL_VERSION = 3;

/**
 * The protocol's caps. Frame lengths are in UTF-16 units of the frame's text; the transport
 * also caps the frame's bytes.
 *
 * @public
 */
export const ORDERED_PROTOCOL_LIMITS = Object.freeze({
  /** Longest frame a client sends. */
  clientFrame: 1_048_576,
  /** Longest frame the room sends. The host splits a long catch-up into `changes` frames. */
  serverFrame: 16_777_216,
  /** Changes in one `welcome` or `changes` frame from the room. */
  serverChanges: 4096,
  /** Acks or refusals in one frame. */
  answers: 4096,
  /** Longest presence state, as JSON. */
  presence: 2048,
  /** Fields in a presence state's `data`. */
  presenceData: 16,
  /** UTF-16 units of a presence `data` value. */
  presenceValue: 256,
  /** UTF-16 units of a room name, a code version, a paragraph id. */
  name: 256,
  /** UTF-16 units of the Word file's location. */
  location: 2048,
  /** UTF-16 units of a refusal's detail. */
  detail: 1024,
  /** Longest host event, as JSON. */
  event: 65_536,
  /** Nesting of a host event's objects and lists. */
  eventDepth: 16,
});

/** A caret or a selection end: a paragraph's node id and a UTF-16 offset in its text. @public */
export interface OrderedPresencePoint {
  readonly paragraph: string;
  readonly offset: number;
}

/** A person's selection. @public */
export interface OrderedPresenceSelection {
  readonly anchor: OrderedPresencePoint;
  readonly head: OrderedPresencePoint;
}

/** What a client says about its person. Needs no order. @public */
export interface OrderedPresenceState {
  readonly selection: OrderedPresenceSelection | null;
  /** The host's own fields (a seat, a colour): flat, small, plain values. */
  readonly data?: Readonly<Record<string, string | number | boolean | null>>;
}

// ---- client to room ----

/** The first message: which room, which Word file, the last order number held, the code. @public */
export interface OrderedHelloMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'hello';
  readonly room: string;
  /** The Word file number the client's copy was built from. */
  readonly file: number;
  /**
   * The last order number the client holds that the room saved; 0 for none. The room sends
   * every change after it, so a room that lost changes it never saved can still answer.
   */
  readonly last: number;
  /** The engine's code version: copies apply changes the same way only on the same code. */
  readonly code: string;
  /** The client's change id namespace. */
  readonly namespace: string;
}

/** A batch of the client's changes. @public */
export interface OrderedClientChangesMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'changes';
  readonly changes: readonly OrderedChange[];
}

/** The client's presence; `state: null` when it leaves. @public */
export interface OrderedClientPresenceMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'presence';
  /** The client's own namespace: presence for any other is refused. */
  readonly namespace: string;
  readonly state: OrderedPresenceState | null;
}

/** An organiser's decision, as a request: the client applies it when the room orders it. @public */
export interface OrderedClientDecideMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'decide';
  readonly change: OrderedDecideChange;
}

/** A message from a client. @public */
export type OrderedClientMessage =
  | OrderedHelloMessage
  | OrderedClientChangesMessage
  | OrderedClientPresenceMessage
  | OrderedClientDecideMessage;

// ---- room to client ----

/**
 * The answer to `hello`: the Word file, the changes after the client's last order number (the
 * first part; more follow in `changes` frames), the current order number and the last saved one.
 *
 * @public
 */
export interface OrderedWelcomeMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'welcome';
  readonly file: number;
  /** Where the client reads the Word file: an `http(s)` URL or a path on the host. */
  readonly location: string;
  /** The room's current order number. */
  readonly n: number;
  /** The last order number the room's storage holds. */
  readonly saved: number;
  readonly changes: readonly OrderedStampedChange[];
  /**
   * Where the client reads media by SHA-256: a location the hash is appended to (an `http(s)`
   * URL or a path on the host). Absent when the room holds no media.
   */
  readonly blobs?: string;
  /**
   * The room's relaxation of content-control locks (`OrderedRoomOptions.contentLockPolicy`):
   * the client applies every change with it, as the room does. Absent: Word's lock.
   */
  readonly contentLockPolicy?: ContentLockPolicy;
}

/** Ordered changes, for everyone in the room. @public */
export interface OrderedServerChangesMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'changes';
  readonly changes: readonly OrderedStampedChange[];
}

/** Order numbers of the client's own changes (new ones and repeated ids). @public */
export interface OrderedAckMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'ack';
  readonly acks: readonly OrderedAck[];
}

/** The client's refused changes. @public */
export interface OrderedRefuseMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'refuse';
  readonly refusals: readonly OrderedRoomRefusal[];
}

/** The room's storage holds every change up to `n`. @public */
export interface OrderedSavedMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'saved';
  readonly n: number;
}

/** Another client's presence, named by the room from that client's login. @public */
export interface OrderedServerPresenceMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'presence';
  readonly namespace: string;
  /** The person the login gives that client. */
  readonly person: OrderedPerson;
  readonly state: OrderedPresenceState | null;
}

/**
 * Why a client must load again.
 *
 * - `version`: the client's protocol or code version is not the room's.
 * - `file`: the room moved to another Word file.
 * - `protocol`: the client sent a frame the room cannot read.
 * - `closed`: the room closed.
 * - `writer`: another server process took the room.
 * - `slow`: the client read the room's messages too slowly, and the room stopped queueing them
 *   for it. It loads the current file again.
 *
 * @public
 */
export type OrderedReloadReason = 'version' | 'file' | 'protocol' | 'closed' | 'writer' | 'slow';

/**
 * Load again from the Word file. The client keeps a copy of its unsent changes for the host.
 *
 * @public
 */
export interface OrderedReloadMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'reload';
  readonly reason: OrderedReloadReason;
}

/** A JSON value of a host event. @public */
export type OrderedEventValue =
  | null
  | boolean
  | number
  | string
  | readonly OrderedEventValue[]
  | { readonly [key: string]: OrderedEventValue };

/**
 * An event of the room's host (the contract service's broadcasts: a freeze, a new version),
 * passed on unread. The client hands `event` to its host; the room's document does not change.
 * Plain JSON, at most {@link ORDERED_PROTOCOL_LIMITS}.event long.
 *
 * @public
 */
export interface OrderedEventMessage {
  readonly v: typeof ORDERED_PROTOCOL_VERSION;
  readonly type: 'event';
  readonly event: { readonly [key: string]: OrderedEventValue };
}

/** A message from the room. @public */
export type OrderedServerMessage =
  | OrderedWelcomeMessage
  | OrderedServerChangesMessage
  | OrderedAckMessage
  | OrderedRefuseMessage
  | OrderedSavedMessage
  | OrderedServerPresenceMessage
  | OrderedReloadMessage
  | OrderedEventMessage;
