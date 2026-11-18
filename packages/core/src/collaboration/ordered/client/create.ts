import type {
  CollaborationEditorPosition,
  CollaborationIdentity,
  EditorCollaborationSession,
} from '../../index.ts';
import type { CollaborationFailureCode } from '../../failure.ts';
import type { OoxmlPackageLimits } from '../../../store/package/ooxml-package.ts';
import { orderedSeenCards } from '../change/apply-decide.ts';
import { orderedKindSpecOf } from '../change/kinds/registry.ts';
import {
  isOrderedEditChange,
  type OrderedChange,
  type OrderedStampedChange,
} from '../change/types.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { checkedPerson, type OrderedPerson } from '../model/person.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  ORDERED_PROTOCOL_VERSION,
  type OrderedPresenceState,
  type OrderedEventMessage,
  type OrderedReloadReason,
  type OrderedRefuseMessage,
  type OrderedServerMessage,
  type OrderedWelcomeMessage,
} from '../protocol/messages.ts';
import { placeOfPoint, type OrderedPlacedRange } from './caret.ts';
import { checkedIdentity, defaultParagraphStyle, sessionNamespace } from './identity.ts';
import {
  defaultOrderedSchedule,
  OrderedFrameError,
  OrderedLink,
  OrderedWordFileUnavailable,
  type OrderedReconnect,
  type OrderedSchedule,
} from './link.ts';
import { openOrderedRoomFile, OrderedWordFileError } from './open.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedBlobRefusalCode } from '../protocol/blobs.ts';
import {
  OrderedMediaReads,
  type OrderedMediaReader,
  type OrderedMediaReadResult,
} from './media-reads.ts';
import {
  OrderedMediaUploads,
  type OrderedMediaUploader,
  type OrderedMediaUploadResult,
} from './media-uploads.ts';
import { OrderedDirty, OrderedPublisher } from './publish.ts';
import type { OrderedClientRefusalReport } from './refusal.ts';
import { OrderedReplica, OrderedReplicaDiverged } from './replica.ts';
import { OrderedSessionBase, type OrderedSessionCore } from './session.ts';
import type { OrderedTransport } from './transport.ts';

/** How a host creates an ordered client. @public */
export interface OrderedCollaborationOptions {
  readonly transport: OrderedTransport;
  /** The room's name, as the room server knows it. */
  readonly room: string;
  /**
   * The engine's code version, the same string the room server runs (for example the
   * `@docx-editor.dev/core` package version): copies apply changes alike only on one code.
   */
  readonly code: string;
  readonly identity: CollaborationIdentity;
  /**
   * The person the room stamps for this login, as the host's login gives it: the local copy
   * stamps it until the room's stamp replaces it. Default: the identity's actor id and name, with
   * no company.
   */
  readonly person?: OrderedPerson;
  /** The editor port's document id. Default: the room's name. */
  readonly documentId?: string;
  readonly sessionId?: string;
  /** The namespace of this session's change ids. Default: a new one. Never reuse one. */
  readonly namespace?: string;
  /**
   * Read the room's Word file `file` from `location`, the place `welcome` names: a Word file, or
   * a room checkpoint (`writeOrderedCheckpoint`), whose media the client reads through `media`.
   */
  readonly wordFile: (location: string, file: number) => Promise<Uint8Array>;
  /**
   * Read a picture the room stores, by its SHA-256, from `location` (the place `welcome` names
   * for media). The bytes are checked against the hash and the content type before they show.
   * Without it, other people's new pictures show no image.
   */
  readonly media?: OrderedMediaReader;
  /**
   * Store a picture the room does not hold yet, at the place `welcome` names for media. A change
   * that names the picture waits until it is stored; one the storage refuses is taken back by
   * name (`blob-refused`).
   */
  readonly upload: OrderedMediaUploader;
  /** Caps on reading the Word file. */
  readonly limits?: OoxmlPackageLimits;
  /** Keep edits while the connection is down and send them on reconnect. Default true. */
  readonly offlineEditing?: boolean;
  /** A local edit, a pending change or an undo was refused. */
  readonly onRefused?: (report: OrderedClientRefusalReport) => void;
  /**
   * Called when a pending change throws something other than `OrderedRefusal` (a defect) as the
   * client applies it again. The change was taken back and refused with `apply-failed`; the host
   * logs the error here.
   */
  readonly onApplyError?: (error: unknown, change: OrderedChange) => void;
  /**
   * This person only suggests (a party in a room with the suggestions-only policy). The client
   * then refuses, before it sends, a section break that gives the next section another type, as
   * the room refuses it (`policy-direct`, reason `package`). The room still checks every change.
   * Default false.
   */
  readonly suggests?: boolean;
  /** The room said to load again; `unsent` are this person's changes it has not ordered. */
  readonly onReload?: (reason: OrderedReloadReason, unsent: readonly OrderedChange[]) => void;
  /**
   * An event of the room's host, passed on unread (the contract service's broadcasts: a new
   * discussion, a freeze, a new version). The document does not change.
   */
  readonly onEvent?: (event: OrderedEventMessage['event']) => void;
  /** The clock for a change's date until the room stamps it. Default: the host's `Date.now`. */
  readonly now?: () => number;
  /** Timers for reconnects and the join timeout. Default: the host's; null: none. */
  readonly schedule?: OrderedSchedule | null;
  /** Waits before each reconnect, in milliseconds. */
  readonly reconnectDelays?: readonly number[];
  /** Give up the first join after this long. Default 30 000 ms. */
  readonly joinTimeoutMs?: number;
}

/**
 * The editor's collaboration session over an ordered room, and what a host adds.
 *
 * @public
 */
export interface OrderedCollaborationSession extends EditorCollaborationSession {
  /** The namespace of this session's change ids. */
  readonly namespace: string;
  /** The last order number this copy applied. */
  order(): number;
  /** The list this copy holds. Read it; write it only through the session. */
  orderedDocument(): OrderedDocument;
  /** This person's changes the room has not ordered, also after a `reload`. */
  unsentChanges(): readonly OrderedChange[];
  /** The last refusal the session reported. */
  lastRefusal(): OrderedClientRefusalReport | null;
  /**
   * Ask the room to decide the card that holds `suggestion`, as this copy shows it (an
   * organiser's decision). Returns the request's change id, or null without a card or a
   * connection. A card that changed since is refused with `stale-decision` through `onRefused`.
   */
  requestDecision(suggestion: string, status: 'accepted' | 'rejected'): string | null;
  /**
   * Ask the room to decide, in one change, the cards that hold `suggestions`, each as this copy
   * shows it (an organiser's Accept All, Reject All, or any choice of several cards). The room
   * decides each card on its own: a card that changed since stays as it is and the others are
   * decided. The change is refused with `stale-decision` through `onRefused` only when no card
   * is as this copy showed it. Returns the request's change id, or null without a
   * card or a connection. After the room's answer, a card whose suggestions are still pending on
   * this copy is one the room left because it changed.
   *
   * @throws Error (an `ordered-refused` refusal, reason `invalid-change`) for a request the room
   * would not read, for example more than `ORDERED_CHANGE_LIMITS.decisions` suggestions in all.
   */
  requestDecisions(suggestions: readonly string[], status: 'accepted' | 'rejected'): string | null;
  /**
   * Add a comment thread of `text` on a range of this editor (from `from` up to `to`; a caret
   * when they are one place). `text` is plain text, one paragraph for each line (`\n`), of 1 to
   * `ORDERED_CHANGE_LIMITS.comment` UTF-16 units. The copy shows it at once and sends it in the
   * next batch; the room stamps the person, whom `orderedDocument().comments.get(id).person`
   * reads. A refusal takes it back and reports through `onRefused`. Comments never enter the
   * undo history. Returns the comment's id, or null when the session takes no edit now, the
   * places are not in one story of the copy, or the copy refuses it.
   */
  addComment(
    text: string,
    from: CollaborationEditorPosition,
    to: CollaborationEditorPosition
  ): string | null;
  /**
   * Add a comment thread of `text` on the selection this session's editor holds (the range it
   * last published, which a host keeps with the editor's `retainSelection` while the person
   * writes), as `addComment` does. The session reads the selection as places of its copy, as it
   * reads the caret for typing, so the host maps no paragraphs itself. Null for a caret, a
   * selection over two stories or in a paragraph the copy does not hold yet.
   */
  addSelectionComment(text: string): string | null;
  /**
   * Reply with `text` to the thread of comment `comment` (its `w:id`), as `addComment` does: the
   * reply answers the thread's first comment, and its range sits beside that comment's. Returns
   * the reply's id, or null.
   */
  replyToComment(comment: string, text: string): string | null;
  /**
   * Resolve the thread of comment `comment` (its `w:id`), as `addComment` does: `note`, when
   * given, is added as this person's last reply, then every comment of the thread is resolved.
   * The room lets only an organiser resolve. Returns whether the copy took it.
   */
  resolveComment(comment: string, note?: string): boolean;
  /** Connect again now. */
  reconnect(): void;
}

/** A created client: the document for the editor, and the session. @public */
export interface OrderedCollaborationHandle {
  /** The room's document as `.docx` bytes, for the editor to open. */
  readonly document: Uint8Array;
  readonly session: OrderedCollaborationSession;
  destroy(): void;
}

const RELOAD_CODES: Readonly<Record<OrderedReloadReason, CollaborationFailureCode>> = {
  version: 'protocol-version-mismatch',
  file: 'room-generation-changed',
  protocol: 'collaboration-format-mismatch',
  closed: 'room-generation-changed',
  writer: 'room-generation-changed',
  // The room stopped queueing for a client that read too slowly: the copy is sound, the link
  // was not
  slow: 'transport',
};

/**
 * The session follows the room's messages: changes, refusals, saves, presence, and `reload`,
 * after which it keeps the unsent changes for the host.
 */
class OrderedClientSession extends OrderedSessionBase implements OrderedCollaborationSession {
  readonly namespace: string;
  readonly #options: OrderedCollaborationOptions;
  readonly #person: OrderedPerson;
  readonly #link: OrderedLink;
  readonly #requests = new Set<string>();
  #unsent: readonly OrderedChange[] = [];
  /** A reload or a fatal error stopped the link: `#unsent` names what was not ordered. */
  #ended = false;
  #presence: OrderedPresenceState | null = null;
  #presenceQueued = false;
  #joinTarget = 0;
  readonly #reads: OrderedMediaReads;
  readonly #uploads: OrderedMediaUploads;
  #joined: { resolve: () => void; reject: (error: unknown) => void } | null = null;

  constructor(
    options: OrderedCollaborationOptions,
    identity: CollaborationIdentity,
    namespace: string,
    schedule: OrderedSchedule | null
  ) {
    const documentId = (options.documentId ?? options.room).trim();
    super({
      documentId,
      sessionId: options.sessionId ?? `${namespace}-session`,
      identity,
      offlineEditing: options.offlineEditing !== false,
      suggests: options.suggests === true,
      ...(options.onRefused === undefined ? {} : { onRefused: options.onRefused }),
      send: (changes) => this.#link.sendChanges(changes),
      sendPresence: (state) => this.#sendPresence(state),
    });
    if (documentId.length === 0 || documentId.length > 256)
      throw new RangeError('A document id has 1 to 256 characters.');
    this.namespace = namespace;
    this.#options = options;
    this.#person = checkedPerson(
      options.person ?? { id: identity.actorId, name: identity.name, org: '' }
    );
    this.#reads = new OrderedMediaReads(this.media, (result) => this.#mediaRead(result));
    this.#uploads = new OrderedMediaUploads(
      options.upload,
      schedule,
      options.reconnectDelays,
      (sha256, result) => this.#uploaded(sha256, result)
    );
    this.#link = new OrderedLink(
      options.transport,
      {
        hello: () => ({
          v: ORDERED_PROTOCOL_VERSION,
          type: 'hello',
          room: options.room,
          // A copy with nothing applied yet names file 1: the room sends its current file
          file: Math.max(1, this.core?.replica.doc.file ?? 1),
          last: this.core?.replica.saved ?? 0,
          code: options.code,
          namespace,
        }),
        message: (message) => this.#message(message),
        opened: () => {
          if (this.#presence !== null) this.#queuePresence();
        },
        closed: (reason, next) => this.#closed(reason, next),
        failed: (error) => this.#failed(error),
      },
      schedule,
      options.reconnectDelays
    );
  }

  /** Connect and wait for the first `welcome` and its catch-up. */
  join(): Promise<void> {
    return new Promise((resolve, reject) => {
      this.#joined = { resolve, reject };
      this.#link.start();
    });
  }

  documentBytes(): Uint8Array {
    const core = this.core;
    if (!core) throw new Error('The room is not open yet.');

    // Media whose bytes the copy has not read yet is left out of the editor's first document
    return orderedPackageBytes(core.replica.doc, core.base, {
      missing: 'skip',
      blobs: (sha256) => this.media.bytes(sha256),
    });
  }

  order(): number {
    return this.core?.replica.confirmed ?? 0;
  }

  orderedDocument(): OrderedDocument {
    const core = this.core;
    if (!core) throw new Error('The room is not open yet.');
    return core.replica.doc;
  }

  unsentChanges(): readonly OrderedChange[] {
    return this.core && !this.#ended ? this.core.replica.pending() : this.#unsent;
  }

  reconnect(): void {
    if (!this.destroyed) this.#link.reconnect();
  }

  // ---- the room's messages ----

  async #message(message: OrderedServerMessage): Promise<void> {
    if (this.destroyed) return;
    switch (message.type) {
      case 'welcome':
        await this.#welcome(message);
        return;
      case 'changes':
        this.#changes(message.changes);
        return;
      case 'refuse':
        this.#refuse(message);
        return;
      case 'saved':
        this.core?.replica.markSaved(message.n);
        return;
      case 'presence':
        if (this.presence.take(message, this.namespace)) {
          this.notifyParticipants();
          this.notifySelections();
        }

        return;
      case 'reload':
        this.#reload(message.reason);
        return;
      case 'event':
        try {
          this.#options.onEvent?.(message.event);
        } catch {
          // A failing host callback must not stop the session
        }

        return;
      case 'ack':
        // The room's `changes` order the person's own changes too; an ack only confirms
        return;
    }
  }

  async #welcome(message: OrderedWelcomeMessage): Promise<void> {
    const core = this.core;
    if (core) {
      this.#rejoin(message, core);
    } else {
      await this.#firstWelcome(message);
    }
  }

  /** The first join: read the room's Word file, open it as the room does, and catch up. */
  async #firstWelcome(message: OrderedWelcomeMessage): Promise<void> {
    let bytes: Uint8Array;
    try {
      bytes = await this.#options.wordFile(message.location, message.file);
    } catch (error) {
      throw new OrderedWordFileUnavailable(error);
    }

    if (this.destroyed) {
      return;
    }

    const opened = openOrderedRoomFile(bytes, {
      file: message.file,
      ...(this.#options.limits === undefined ? {} : { limits: this.#options.limits }),
    });
    const replica = new OrderedReplica({
      doc: opened.document,
      person: this.#person,
      namespace: this.namespace,
      now: this.#options.now ?? (() => Date.now()),
      ...(this.#options.onApplyError === undefined
        ? {}
        : { onApplyError: this.#options.onApplyError }),
      ...(message.contentLockPolicy === undefined
        ? {}
        : { contentLockPolicy: message.contentLockPolicy }),
    });
    this.core = {
      replica,
      base: opened.base,
      defaultParagraphStyle: defaultParagraphStyle(opened.base),
    };

    // The room's storage holds the media a checkpoint names: a copy of it needs no upload
    for (const item of opened.document.resources.fileMedia.values()) {
      this.media.markStored(item.sha256);
    }

    this.publisher = new OrderedPublisher(opened.document, opened.base, {
      blobs: (sha256) => this.media.bytes(sha256),
    });
    this.#reads.connect(message.blobs ?? null, this.#options.media ?? null);
    this.#uploads.connect(message.blobs ?? null);
    this.#joinTarget = message.n;
    this.#take(message.changes);
    replica.markSaved(message.saved);
    this.#settleJoin();
    this.#link.welcomed();
  }

  /** A join after a dropped link: check the file, go back to the save, take the room's order. */
  #rejoin(message: OrderedWelcomeMessage, core: OrderedSessionCore): void {
    const replica = core.replica;
    if (message.file !== replica.doc.file) {
      this.#reload('file');
      return;
    }

    if (message.n < replica.saved) {
      this.#fatal(
        'room-generation-changed',
        `the room is at ${message.n}, before its save at ${replica.saved}`
      );
      return;
    }

    // A restart may have lost what came after the save: back to it, then the room's order
    this.#reads.connect(message.blobs ?? null, this.#options.media ?? null);
    this.#uploads.connect(message.blobs ?? null);
    const held = this.holdSelection();
    this.dirty(replica.rollbackToSaved());
    this.#take(message.changes, held);

    // The catch-up did not apply: the session ended
    if (this.#ended) {
      return;
    }

    replica.markSaved(message.saved);
    this.setStatus('ready');
    this.sendQueued();
    this.#link.welcomed();
  }

  /** The room's changes: take back, apply, apply again, publish. */
  #take(changes: readonly OrderedStampedChange[], held = this.holdSelection()): void {
    const replica = this.core?.replica;
    if (!replica) return;
    try {
      const cycle = replica.receive(changes);
      this.dirty(cycle.dirty);
      for (const refused of cycle.refused) this.reportRefusal(refused);
    } catch (error) {
      if (error instanceof OrderedReplicaDiverged) {
        this.#fatal('remote-apply-failed', error.message);
        return;
      }

      throw error;
    }

    this.publishNow(null, held);
    this.#reads.readMissing(replica.doc);
  }

  #changes(changes: readonly OrderedStampedChange[]): void {
    if (!this.core) return;

    // A decision the room ordered is answered
    for (const change of changes) this.#requests.delete(change.id);
    this.#take(changes);
    this.#settleJoin();
  }

  #refuse(message: OrderedRefuseMessage): void {
    const replica = this.core?.replica;
    if (!replica) return;
    const ids = new Set<string>();
    for (const refusal of message.refusals) {
      // A batch refused before its ids could be read: this client sent what it cannot send
      if (refusal.id === null) {
        this.#fatal(
          'collaboration-format-mismatch',
          `the room cannot read a batch: ${refusal.code}`
        );
        return;
      }

      this.reportRefusal({
        code: 'ordered-refused',
        reason: refusal.reason ?? refusal.code,
        ...(refusal.detail === undefined ? {} : { detail: refusal.detail }),
        change: refusal.id,
      });
      if (this.#requests.delete(refusal.id)) continue;
      ids.add(refusal.id);
    }

    if (ids.size === 0) {
      return;
    }

    this.#takeBack(replica, ids);
  }

  /** Take back the person's refused changes by name, and publish what is left. */
  #takeBack(replica: OrderedReplica, ids: ReadonlySet<string>): void {
    const held = this.holdSelection();
    const cycle = replica.refuse(ids);
    this.dirty(cycle.dirty);
    for (const refused of cycle.refused) {
      this.reportRefusal(refused);
    }

    this.publishNow(null, held);
  }

  #settleJoin(): void {
    const joined = this.#joined;
    const core = this.core;
    if (!joined || !core || core.replica.confirmed < this.#joinTarget) return;
    this.#joined = null;
    this.setStatus('ready');
    joined.resolve();
  }

  #reload(reason: OrderedReloadReason): void {
    this.#unsent = this.core?.replica.pending() ?? [];
    this.#ended = true;
    this.#link.stop();
    this.#uploads.stop();
    this.setStatus('error', RELOAD_CODES[reason], `reload: ${reason}`);
    this.#joined?.reject(new Error(`The room asked for a reload: ${reason}.`));
    this.#joined = null;
    try {
      this.#options.onReload?.(reason, this.#unsent);
    } catch {
      // A failing host callback must not stop the session
    }
  }

  #fatal(code: CollaborationFailureCode, detail: string): void {
    this.#unsent = this.core?.replica.pending() ?? [];
    this.#ended = true;
    this.#link.stop();
    this.#uploads.stop();
    this.setStatus('error', code, detail.slice(0, 300));
    this.#joined?.reject(new Error(detail));
    this.#joined = null;
  }

  // ---- media (design D14) ----

  /** A change waits until the room stores the media it names; its upload starts here. */
  protected override mediaReady(change: OrderedChange): boolean {
    if (!isOrderedEditChange(change)) return true;
    const hashes = orderedKindSpecOf(change.kind)?.blobs?.(change) ?? [];
    return hashes.length === 0 || this.media.ready(hashes, (blob) => this.#uploads.start(blob));
  }

  /** The storage's answer to an upload: the changes that wait for it go, or are taken back. */
  #uploaded(sha256: string, result: OrderedMediaUploadResult): void {
    if (this.destroyed) {
      return;
    }

    if (result.stored) {
      this.media.markStored(sha256);
      this.sendQueued();
      return;
    }

    this.#blobRefused(sha256, result.code);
  }

  /**
   * The storage refused a blob: the person's changes that name it cannot go. They are taken back
   * by name, as a room's refusal is, and the changes after them are sent.
   */
  #blobRefused(sha256: string, code: OrderedBlobRefusalCode): void {
    this.media.markRefused(sha256, code);
    const replica = this.core?.replica;
    if (!replica) return;
    const ids = new Set<string>();
    for (const change of replica.pending()) {
      if (!isOrderedEditChange(change)) continue;
      if (orderedKindSpecOf(change.kind)?.blobs?.(change).includes(sha256)) ids.add(change.id);
    }

    for (const id of ids)
      this.reportRefusal({
        code: 'ordered-refused',
        reason: 'blob-refused',
        detail: code,
        change: id,
      });
    if (ids.size === 0) {
      return;
    }

    this.#takeBack(replica, ids);
    this.sendQueued();
  }

  /** A picture read back from the room: shown at once, or reported by name. */
  #mediaRead(result: OrderedMediaReadResult): void {
    if (this.destroyed) return;
    if (!result.ok) {
      this.report({ code: 'ordered-refused', reason: result.reason, detail: result.sha256 });
      return;
    }

    // The media part and its relationship are written at the next publish
    const dirty = new OrderedDirty();
    dirty.stale = true;
    this.dirty(dirty);
    this.publishNow();
  }

  /** Every picture read the session started so far, settled. */
  mediaSettled(): Promise<void> {
    return this.#reads.settled();
  }

  #closed(reason: string, next: OrderedReconnect | null): void {
    if (this.presence.clear()) {
      this.notifyParticipants();
      this.notifySelections();
    }

    this.core?.replica.markAllUnsent();

    // The detail says which try comes next and when, so a host can show a room that keeps
    // closing the page
    const detail = next
      ? `${reason.slice(0, 160)}; try ${next.attempt} in ${(next.delayMs / 1000).toFixed(1)} s`
      : reason.slice(0, 200);
    if (this.core && this.statusState.status() !== 'error')
      this.setStatus('disconnected', 'transport-disconnected', detail);
  }

  #failed(error: unknown): void {
    const message = error instanceof Error ? error.message : 'failed';
    if (error instanceof OrderedFrameError) this.#fatal('collaboration-format-mismatch', message);
    else if (error instanceof OrderedWordFileUnavailable) this.#fatal('blob-read', message);
    // The Word file does not open, or is past the load limits
    else if (error instanceof OrderedWordFileError || error instanceof OrderedRefusal)
      this.#fatal('invalid-baseline', message);
    else this.#fatal('remote-apply-failed', error instanceof Error ? error.name : 'failed');
  }

  // ---- presence and requests ----

  #sendPresence(state: OrderedPresenceState | null): void {
    this.#presence = state;
    this.#queuePresence();
  }

  #queuePresence(): void {
    if (this.#presenceQueued) return;
    this.#presenceQueued = true;
    void Promise.resolve().then(() => {
      this.#presenceQueued = false;
      this.#link.send({
        v: ORDERED_PROTOCOL_VERSION,
        type: 'presence',
        namespace: this.namespace,
        state: this.#presence,
      });
    });
  }

  requestDecision(suggestion: string, status: 'accepted' | 'rejected'): string | null {
    return this.requestDecisions([suggestion], status);
  }

  requestDecisions(suggestions: readonly string[], status: 'accepted' | 'rejected'): string | null {
    const core = this.core;
    if (!core || !this.#link.open) return null;
    const cards = orderedSeenCards(core.replica.doc, suggestions);
    if (cards.length === 0) return null;
    const change = this.validate({
      v: 1,
      id: core.replica.nextId(),
      kind: 'decide',
      cards,
      status,
    });
    if (
      change.kind !== 'decide' ||
      !this.#link.send({ v: ORDERED_PROTOCOL_VERSION, type: 'decide', change })
    ) {
      return null;
    }

    this.#requests.add(change.id);
    return change.id;
  }

  addComment(
    text: string,
    from: CollaborationEditorPosition,
    to: CollaborationEditorPosition
  ): string | null {
    const doc = this.core?.replica.doc;
    if (!doc) {
      return null;
    }

    const start = placeOfPoint(doc, from, null);
    const end = placeOfPoint(doc, to, null);
    if (!start || !end || start.story !== end.story) {
      return null;
    }

    return this.#newThread(text, { story: start.story, from: start.place, to: end.place });
  }

  addSelectionComment(text: string): string | null {
    const range = this.selectionRange();
    return range ? this.#newThread(text, range) : null;
  }

  #newThread(text: string, range: OrderedPlacedRange): string | null {
    const added = this.writeComment({
      kind: 'comment',
      part: range.story,
      from: range.from,
      to: range.to,
      text,
    });
    return added?.[0] ?? null;
  }

  replyToComment(comment: string, text: string): string | null {
    return this.writeComment({ kind: 'comment-reply', comment, text })?.[0] ?? null;
  }

  resolveComment(comment: string, note?: string): boolean {
    const body = note === undefined ? { comment } : { comment, note };
    return this.writeComment({ kind: 'comment-resolve', ...body }) !== null;
  }

  override destroy(): void {
    if (this.destroyed) return;
    super.destroy();
    this.#link.stop();
    this.#uploads.stop();
    this.#joined?.reject(new Error('The session was destroyed.'));
    this.#joined = null;
  }
}

/**
 * Join an ordered room: connect, read the room's Word file, catch up, and return the document
 * for the editor with the session to give the editor's collaboration module.
 *
 * @throws RangeError for an invalid identity, person, namespace or document id.
 * @throws Error when the room asks for a reload, the join times out, or the copy cannot follow.
 * @public
 */
export async function createOrderedCollaboration(
  options: OrderedCollaborationOptions
): Promise<OrderedCollaborationHandle> {
  const identity = checkedIdentity(options.identity);
  const namespace = sessionNamespace(identity.actorId, options.namespace);
  const schedule = options.schedule === undefined ? defaultOrderedSchedule() : options.schedule;
  const joinTimeoutMs = options.joinTimeoutMs ?? 30_000;
  const session = new OrderedClientSession(options, identity, namespace, schedule);
  const cancel = schedule?.(() => {
    if (session.status() === 'initializing') {
      session.destroy();
    }
  }, joinTimeoutMs);
  try {
    await session.join();

    // The pictures of the catch-up, read before the editor's first document, for at most the
    // join timeout: a picture read later shows then, and one that does not read is reported
    let stopWaiting: (() => void) | undefined;
    await Promise.race([
      session.mediaSettled(),
      new Promise<void>((resolve) => {
        stopWaiting = schedule?.(resolve, joinTimeoutMs);
      }),
    ]);
    stopWaiting?.();
  } catch (error) {
    session.destroy();
    throw error;
  } finally {
    cancel?.();
  }

  return Object.freeze({
    document: session.documentBytes(),
    session,
    destroy: () => session.destroy(),
  });
}
