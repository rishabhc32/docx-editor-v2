import type { ContentLockPolicy } from '../../../store/package/ooxml-edit.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange, OrderedStampedChange } from '../change/types.ts';
import { decodeOrderedClientMessage, encodeOrderedServerMessage } from '../protocol/codec.ts';
import {
  ORDERED_PROTOCOL_LIMITS,
  ORDERED_PROTOCOL_VERSION,
  type OrderedHelloMessage,
  type OrderedReloadReason,
  type OrderedServerMessage,
  type OrderedServerPresenceMessage,
} from '../protocol/messages.ts';
import { allowWritersPolicy, type OrderedActor, type OrderedRoomPolicy } from '../room/policy.ts';
import { admitOrderedPresence } from '../room/presence.ts';
import { ORDERED_RESOURCE_LIMITS } from '../change/resources.ts';
import { checkedPerson } from '../model/person.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { OrderedRoom } from '../room/room.ts';
import { isOrderedMediaBytes } from './media-check.ts';
import type { OrderedMediaUploader } from './media-uploads.ts';
import { openOrderedRoomFile } from './open.ts';
import type { OrderedConnection, OrderedConnectionEvents, OrderedTransport } from './transport.ts';

/** How a local room is set up. @public */
export interface OrderedLocalRoomOptions {
  /** The room's name, as `hello` names it. */
  readonly room: string;
  /** The room's Word file, or a room checkpoint (`writeOrderedCheckpoint`), and its number. */
  readonly wordFile: Uint8Array;
  readonly file: number;
  /**
   * The media the room's storage holds from the start, by SHA-256: the pictures a checkpoint
   * names (`orderedPackageMedia` of the Word file it came from).
   */
  readonly blobs?: ReadonlyMap<string, Uint8Array>;
  /** The engine code version the room speaks. */
  readonly code: string;
  readonly policy?: OrderedRoomPolicy;
  /** The room's clock. Default: the host's `Date.now`. */
  readonly now?: () => number;
  /** Run a delivery later. Default: a microtask. A test can queue them and run them itself. */
  readonly deliver?: (run: () => void) => void;
  /** Send `saved` after every batch, as storage that saves at once. Default true. */
  readonly autoSave?: boolean;
  /** The room's relaxation of content-control locks (`OrderedRoomOptions.contentLockPolicy`). */
  readonly contentLockPolicy?: ContentLockPolicy;
}

/**
 * Who a connection logs in as: everything of an actor but the namespace its `hello` names.
 *
 * @public
 */
export type OrderedLocalLogin = Omit<OrderedActor, 'namespace'>;

interface Peer {
  readonly login: OrderedLocalLogin;
  readonly events: OrderedConnectionEvents;
  namespace: string | null;
  open: boolean;
}

const later = (run: () => void): void => {
  void Promise.resolve().then(run);
};

/**
 * A room in this process, with one transport for each login, for tests and local hosts:
 * `OrderedRoom` behind the protocol, as the host's server answers it, with frames delivered later
 * (a microtask by default, or whenever a test says). It keeps the log in memory.
 *
 * @public
 */
export class OrderedLocalRoom {
  readonly #options: OrderedLocalRoomOptions;
  readonly #deliver: (run: () => void) => void;
  readonly #peers = new Set<Peer>();
  readonly #presence = new Map<Peer, OrderedServerPresenceMessage>();
  /** The media the room's storage holds, by SHA-256: kept across restarts, as storage is. */
  readonly #blobs: Map<string, Uint8Array>;
  #room: OrderedRoom;
  #saved = 0;

  constructor(options: OrderedLocalRoomOptions) {
    this.#options = options;
    this.#deliver = options.deliver ?? later;
    this.#blobs = new Map(options.blobs);
    this.#room = this.#open([]);
  }

  #open(log: readonly OrderedStampedChange[]): OrderedRoom {
    const { document } = openOrderedRoomFile(this.#options.wordFile, { file: this.#options.file });
    return new OrderedRoom({
      document,
      applyChange: applyOrderedChange,
      policy: this.#options.policy ?? allowWritersPolicy,
      now: this.#options.now ?? (() => Date.now()),
      log,
      blobs: this.#blobs,
      ...(this.#options.contentLockPolicy === undefined
        ? {}
        : { contentLockPolicy: this.#options.contentLockPolicy }),
    });
  }

  /** The bytes of media the room's storage holds, by SHA-256. */
  blob(sha256: string): Uint8Array | undefined {
    return this.#blobs.get(sha256);
  }

  /** The room core: its document and its order. */
  get room(): OrderedRoom {
    return this.#room;
  }

  /** The last order number the room's storage holds. */
  get saved(): number {
    return this.#saved;
  }

  /** Open connections. */
  get connections(): number {
    let open = 0;
    for (const peer of this.#peers) if (peer.open) open += 1;
    return open;
  }

  /**
   * A transport whose connections log in as `login`, with its person trimmed and checked, as the
   * room keeps a namespace's owner.
   *
   * @throws RangeError when the login's person is not valid.
   */
  transport(login: OrderedLocalLogin): OrderedTransport {
    const checked = { ...login, person: checkedPerson(login.person) };
    return {
      connect: (events: OrderedConnectionEvents): OrderedConnection => {
        const peer: Peer = { login: checked, events, namespace: null, open: true };
        this.#peers.add(peer);
        this.#deliver(() => {
          if (peer.open) events.open();
        });
        return {
          send: (frame: string): boolean => {
            if (!peer.open) return false;
            this.#deliver(() => this.#receive(peer, frame));
            return true;
          },
          close: () => this.#drop(peer, null),
        };
      },
    };
  }

  /**
   * How a host of `login` stores media, as the host's server answers an upload: checked by size,
   * content type and hash, then kept under the hash. A read-only login, or another location than
   * the room's, fails as a transient fault does.
   */
  uploader(login: OrderedLocalLogin): OrderedMediaUploader {
    return async ({ location, sha256, contentType, bytes }) => {
      if (login.readOnly || location !== this.#blobsLocation()) {
        throw new Error('The room does not take this upload.');
      }

      if (bytes.byteLength > ORDERED_RESOURCE_LIMITS.mediaSize) {
        return { stored: false, code: 'blob-too-large' };
      }

      if (!isOrderedMediaBytes(bytes, contentType)) {
        return { stored: false, code: 'blob-type' };
      }

      if (orderedBlobHash(bytes) !== sha256) {
        return { stored: false, code: 'blob-hash' };
      }

      // A copy of its own, as storage keeps
      this.#blobs.set(sha256, bytes.slice());
      return { stored: true };
    };
  }

  /** Where the room's files are: its Word files and its media. */
  #roomPath(): string {
    return `/ordered/${encodeURIComponent(this.#options.room)}/`;
  }

  /** Where copies read and store media, by hash. */
  #blobsLocation(): string {
    return `${this.#roomPath()}blobs/`;
  }

  /** The storage holds every change up to `n` (default: all): send `saved`. */
  save(n = this.#room.order): void {
    this.#saved = Math.min(Math.max(this.#saved, n), this.#room.order);
    this.#broadcast({ v: ORDERED_PROTOCOL_VERSION, type: 'saved', n: this.#saved });
  }

  /** Drop every connection, or one person's (by namespace), as a network fault would. */
  disconnect(namespace?: string): void {
    for (const peer of [...this.#peers])
      if (namespace === undefined || peer.namespace === namespace) this.#drop(peer, 'network');
  }

  /** Send `reload` to every connection and drop it, as a room that moved to its next file does. */
  reloadAll(reason: OrderedReloadReason): void {
    for (const peer of [...this.#peers]) this.#reload(peer, reason);
  }

  /**
   * A restart that kept the first `keep` changes of the log: the room is rebuilt from the Word
   * file and them, and every connection drops.
   */
  restart(keep: number): void {
    const log = this.#room.joinSnapshot(0).slice(0, Math.max(0, keep));
    this.disconnect();
    this.#room = this.#open(log);
    this.#saved = Math.min(this.#saved, this.#room.order);
  }

  #send(peer: Peer, message: OrderedServerMessage): void {
    if (!peer.open) return;
    const frame = encodeOrderedServerMessage(message);
    this.#deliver(() => {
      if (peer.open) peer.events.message(frame);
    });
  }

  #broadcast(message: OrderedServerMessage, except: Peer | null = null): void {
    for (const peer of this.#peers)
      if (peer !== except && peer.namespace !== null) this.#send(peer, message);
  }

  #drop(peer: Peer, reason: string | null): void {
    if (!peer.open) return;
    peer.open = false;
    this.#peers.delete(peer);
    const presence = this.#presence.get(peer);
    this.#presence.delete(peer);
    if (presence) this.#broadcast({ ...presence, state: null });
    if (reason !== null) this.#deliver(() => peer.events.close(reason));
  }

  #reload(peer: Peer, reason: OrderedReloadReason): void {
    this.#send(peer, { v: ORDERED_PROTOCOL_VERSION, type: 'reload', reason });
    this.#deliver(() => this.#drop(peer, `reload ${reason}`));
  }

  #actor(peer: Peer): OrderedActor | null {
    return peer.namespace === null ? null : { ...peer.login, namespace: peer.namespace };
  }

  #receive(peer: Peer, frame: string): void {
    if (!peer.open) return;
    const decoded = decodeOrderedClientMessage(
      frame,
      peer.namespace === null ? {} : { namespace: peer.namespace }
    );
    if (!decoded.ok) {
      this.#reload(peer, decoded.error.code === 'invalid-version' ? 'version' : 'protocol');
      return;
    }

    const message = decoded.value;
    if (message.type === 'hello') {
      this.#hello(peer, message);
      return;
    }

    const actor = this.#actor(peer);
    if (actor === null) {
      this.#reload(peer, 'protocol');
      return;
    }

    if (message.type === 'presence') {
      const verdict = admitOrderedPresence(actor, message, {
        ownerOf: (namespace) => this.#room.namespaceOwner(namespace),
      });
      if (!verdict.ok) return;
      if (message.state === null) this.#presence.delete(peer);
      else this.#presence.set(peer, verdict.message);
      this.#broadcast(verdict.message, peer);
      return;
    }

    this.#submit(peer, actor, message.type === 'changes' ? message.changes : [message.change]);
  }

  #hello(peer: Peer, hello: OrderedHelloMessage): void {
    const { room, file, last, code, namespace } = hello;
    if (room !== this.#options.room || code !== this.#options.code) {
      this.#reload(peer, room !== this.#options.room ? 'protocol' : 'version');
      return;
    }

    const owner = this.#room.namespaceOwner(namespace);
    if ((owner !== undefined && owner !== peer.login.person.id) || peer.namespace !== null) {
      this.#drop(peer, 'namespace');
      return;
    }

    const sameFile = file === this.#options.file;
    if ((!sameFile && last > 0) || (sameFile && last > this.#room.order)) {
      this.#reload(peer, 'file');
      return;
    }

    peer.namespace = namespace;
    const changes = this.#room.joinSnapshot(sameFile ? last : 0);
    const cap = ORDERED_PROTOCOL_LIMITS.serverChanges;
    this.#send(peer, {
      v: ORDERED_PROTOCOL_VERSION,
      type: 'welcome',
      file: this.#options.file,
      location: `${this.#roomPath()}${this.#options.file}.docx`,
      blobs: this.#blobsLocation(),
      n: this.#room.order,
      saved: Math.min(this.#saved, this.#room.order),
      changes: changes.slice(0, cap),
      ...(this.#options.contentLockPolicy === undefined
        ? {}
        : { contentLockPolicy: this.#options.contentLockPolicy }),
    });
    for (let at = cap; at < changes.length; at += cap) {
      this.#send(peer, {
        v: ORDERED_PROTOCOL_VERSION,
        type: 'changes',
        changes: changes.slice(at, at + cap),
      });
    }

    for (const [other, presence] of this.#presence) {
      if (other !== peer) {
        this.#send(peer, presence);
      }
    }
  }

  #submit(peer: Peer, actor: OrderedActor, changes: readonly OrderedChange[]): void {
    const result = this.#room.submit(actor, changes);
    if (result.accepted.length > 0) {
      this.#broadcast({ v: ORDERED_PROTOCOL_VERSION, type: 'changes', changes: result.accepted });
    }

    if (result.acks.length > 0) {
      this.#send(peer, { v: ORDERED_PROTOCOL_VERSION, type: 'ack', acks: result.acks });
    }

    if (result.refused.length > 0) {
      this.#send(peer, { v: ORDERED_PROTOCOL_VERSION, type: 'refuse', refusals: result.refused });
    }

    if (result.accepted.length > 0 && this.#options.autoSave !== false) {
      this.save();
    }
  }
}
