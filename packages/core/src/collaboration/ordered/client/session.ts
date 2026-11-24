// The ordered client session (design D7, tasks 6.2 to 6.4): the editor's collaboration session
// over the ordered list.

import type {
  CollaborationIdentity,
  CollaborationGateContext,
  CollaborationLocalSelection,
  CollaborationReviewWrite,
  CollaborationParticipant,
  CollaborationRemoteSelection,
  CollaborationSelectionMove,
  EditorCollaborationSession,
} from '../../index.ts';
import {
  createCollaborationStatusTracker,
  type CollaborationFailureCode,
  type CollaborationStatus,
  type CollaborationStatusSnapshot,
} from '../../failure.ts';
import type { CollaborationDocumentPort } from '../../replication.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import type { HistoryGroup } from '../../../store/store/history-group.ts';
import type { TreeModelChange } from '../../../store/store/tree-store.ts';
import type { StoryScope } from '../../../store/store/tree-package-store.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import { ORIGIN_IDS } from '../../../store/registry/frozen-ids.ts';
import type { OrderedChange, OrderedCommentThreadChange } from '../change/types.ts';
import { validateOrderedChange } from '../change/validate.ts';
import { retypingBreakRefusal } from '../change/kinds/sections/break.ts';
import { pendingBy } from '../model/document.ts';
import { makerOf } from '../model/person.ts';
import { errorName } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  caretOfBatch,
  OrderedSelectionCarry,
  type HeldSelection,
  type OrderedPlacedPoint,
  type OrderedPlacedRange,
} from './caret.ts';
import { OrderedHistory } from './history.ts';
import { OrderedMedia } from './media.ts';
import { markOfNode, shownParagraphOfNode } from './paragraphs.ts';
import {
  OrderedPresence,
  orderedPresenceState,
  type OrderedPresenceParagraph,
} from './presence.ts';
import { OrderedDirty, OrderedPublisher } from './publish.ts';
import { OrderedClientRefusal, type OrderedClientRefusalReport } from './refusal.ts';
import { OrderedReplica } from './replica.ts';
import { translateOps, type TranslatedBatch } from './translate.ts';
import { redoOwnStep, undoOwnStep, type OrderedHistoryWriter } from './undo.ts';
import {
  isEditBody,
  type OrderedChangeBody,
  type OrderedReviewContext,
  type WithoutIds,
} from './ops/types.ts';

/** A comment change without the fields the session adds (`v`, `id`). */
type OrderedCommentBody = WithoutIds<OrderedCommentThreadChange>;

/** What the session needs once the room's Word file is open. */
export interface OrderedSessionCore {
  readonly replica: OrderedReplica;
  readonly base: OoxmlPackage;
  readonly defaultParagraphStyle: string | null;
}

/** The session's settings. */
export interface OrderedSessionSettings {
  readonly documentId: string;
  readonly sessionId: string;
  readonly identity: CollaborationIdentity;
  /** Edits while the connection is down are kept and sent on reconnect. Default true. */
  readonly offlineEditing: boolean;
  readonly onRefused?: (report: OrderedClientRefusalReport) => void;
  /** This person only suggests: the gate refuses a break that retypes the next section. */
  readonly suggests?: boolean;
  /** Send changes and presence; false while the room is not reachable. */
  readonly send: (changes: readonly OrderedChange[]) => boolean;
  readonly sendPresence: (state: ReturnType<typeof orderedPresenceState> | null) => void;
}

/** A change the gate applied, and the person's own pending suggestions it made. */
interface AdmittedChange {
  readonly change: OrderedChange;
  readonly own: readonly string[];
}

/** An edit the gate admitted, until the editor commits it. */
interface Admitted {
  /** The local commit count the editor's commit makes. */
  readonly commit: number;
  readonly changes: readonly AdmittedChange[];
  readonly selection: CollaborationLocalSelection | null;
  readonly caret: OrderedPlacedPoint | null;
  readonly direct: boolean;
  /** The editor's undo group of the commit: commits of one group are one step. */
  readonly group: HistoryGroup | undefined;
}

/**
 * The markers the store writes for a new comment, as ops the gate translates into one `comment`
 * change. The store mints the comment's own id when it commits; the list mints its own.
 */
function commentMarkerOps(comment: NonNullable<CollaborationReviewWrite['comment']>): TreeDocOp[] {
  const end = comment.endParagraphId ?? comment.paragraphId;
  const marker = (paragraphId: string, offset: number, kind: 'start' | 'end' | 'reference') => ({
    op: 'insertCommentMarker' as const,
    paragraphId,
    offset,
    commentId: 'new',
    marker: kind,
  });
  return [
    marker(comment.paragraphId, comment.start, 'start'),
    marker(end, comment.end, 'end'),
    marker(end, comment.end, 'reference'),
  ];
}

const writable = (status: CollaborationStatus, offline: boolean): boolean =>
  status === 'ready' || (offline && status === 'disconnected');

/** How many recent own changes and actions the session remembers for its refusal reports. */
const ACTION_MEMORY = 4096;

/** Drop the oldest entry of a set or a map that holds more than `ACTION_MEMORY` entries. */
function trimOldest<K>(collection: Set<K> | Map<K, unknown>): void {
  if (collection.size <= ACTION_MEMORY) {
    return;
  }

  const oldest = collection.keys().next();
  if (!oldest.done) {
    collection.delete(oldest.value);
  }
}

export class OrderedSessionBase implements EditorCollaborationSession {
  readonly documentId: string;
  readonly sessionId: string;
  readonly identity: CollaborationIdentity;
  /**
   * A tracked replacement over several places is one batch of changes, placed by addresses
   * that never move, so concurrent edits cannot pull its pieces apart (design D7).
   */
  readonly trackedRangeReplacement = true;
  protected readonly settings: OrderedSessionSettings;
  protected core: OrderedSessionCore | null = null;
  protected publisher: OrderedPublisher | null = null;
  protected readonly history = new OrderedHistory();
  /** Pictures this copy holds by hash: its own to send, others' read back (design D14). */
  protected readonly media = new OrderedMedia();
  protected readonly presence = new OrderedPresence();
  protected readonly statusState = createCollaborationStatusTracker();
  protected destroyed = false;
  #port: CollaborationDocumentPort | null = null;
  #stopPort: (() => void) | null = null;
  #publishing = false;
  #localCommits = 0;
  #admitted: Admitted | null = null;
  readonly #dirty = new OrderedDirty();
  #ungated = false;
  #flushQueued = false;
  readonly #carry = new OrderedSelectionCarry();
  #publishCount = 0;
  #lastRefusal: OrderedClientRefusalReport | null = null;
  /** The user action each own change was made in: one editor commit, one undo or redo. */
  readonly #actions = new Map<string, number>();
  #nextAction = 0;
  /** The actions a refusal was reported for. */
  readonly #reportedActions = new Set<number>();
  readonly #statusListeners = new Set<
    (status: CollaborationStatus, reason?: CollaborationFailureCode, detail?: string) => void
  >();
  readonly #participantListeners = new Set<
    (participants: readonly CollaborationParticipant[]) => void
  >();
  readonly #selectionListeners = new Set<
    (selections: readonly CollaborationRemoteSelection[]) => void
  >();

  constructor(settings: OrderedSessionSettings) {
    this.settings = settings;
    this.documentId = settings.documentId;
    this.sessionId = settings.sessionId;
    this.identity = settings.identity;
  }

  // ---- status ----

  status(): CollaborationStatus {
    return this.statusState.status();
  }

  statusSnapshot(): CollaborationStatusSnapshot {
    return this.statusState.snapshot();
  }

  subscribeStatus(
    listener: (
      status: CollaborationStatus,
      reason?: CollaborationFailureCode,
      detail?: string
    ) => void
  ): () => void {
    this.#statusListeners.add(listener);
    return () => this.#statusListeners.delete(listener);
  }

  protected setStatus(
    status: CollaborationStatus,
    code?: CollaborationFailureCode,
    detail?: string
  ): void {
    if (this.destroyed && status !== 'destroyed') return;
    if (!this.statusState.set(status, code, detail)) return;
    const snapshot = this.statusState.snapshot();
    for (const listener of [...this.#statusListeners])
      listener(snapshot.status, snapshot.reason?.code, snapshot.reason?.detail);
  }

  /** The last refusal the session reported. */
  lastRefusal(): OrderedClientRefusalReport | null {
    return this.#lastRefusal;
  }

  /**
   * The room's refusal of an own change, reported once for the user action that made it: the
   * first refusal names the cause; the action's later changes, refused because they need the
   * first one, are taken back without a report of their own. The redo steps go: a refused undo
   * took nothing back, so its redo would make the words a second time.
   */
  protected reportRefusal(report: OrderedClientRefusalReport): void {
    if (report.change !== undefined) {
      this.history.clearRedo();
    }

    const action = report.change === undefined ? undefined : this.#actions.get(report.change);
    if (action !== undefined) {
      if (this.#reportedActions.has(action)) {
        return;
      }

      this.#reportedActions.add(action);
      trimOldest(this.#reportedActions);
    }

    this.report(report);
  }

  /** A change made in a user action, as the action's. */
  #acted<C extends { readonly id: string }>(change: C, action: number): C {
    this.#actions.set(change.id, action);
    trimOldest(this.#actions);
    return change;
  }

  protected report(report: OrderedClientRefusalReport): void {
    this.#lastRefusal = report;
    try {
      this.settings.onRefused?.(report);
    } catch {
      // A failing host callback must not stop the session
    }
  }

  protected get canWrite(): boolean {
    return (
      !this.destroyed &&
      this.core !== null &&
      writable(this.statusState.status(), this.settings.offlineEditing)
    );
  }

  // ---- attach ----

  get attached(): boolean {
    return this.#port !== null;
  }

  attach(port: CollaborationDocumentPort): () => void {
    if (this.destroyed) return () => {};
    if (this.#port) {
      if (this.#port !== port) this.setStatus('error', 'port-already-attached');
      return () => {};
    }

    if (port.documentId !== this.documentId) {
      this.setStatus('error', 'document-id-mismatch');
      return () => {};
    }

    this.#port = port;
    const stop = port.subscribe((change) => this.#observe(change));
    this.#stopPort = stop;

    // The editor shows the list from the start, under the list's own node ids
    if (this.publisher) this.#publish(null, true);
    return () => {
      if (this.#port !== port) return;
      this.#lastFlush();
      stop();
      this.#port = null;
      this.#stopPort = null;
    };
  }

  #observe(change: TreeModelChange): void {
    if (this.destroyed || change.origin === ORIGIN_IDS.mutationRemote) return;

    // A commit that runs during our publish (a key the host was holding) is the gate's admitted
    // edit: count it, or the edit would be taken back with no word. Any other write during a
    // publish is ours
    if (this.#publishing && this.#admitted === null) {
      return;
    }

    this.#localCommits += 1;
    if (this.#admitted === null) this.#ungated = true;
    this.#noteEngine(change);
    this.#queueFlush();
  }

  /**
   * What the engine's commit wrote: its paragraphs too are drawn again from the list. A
   * paragraph the list does not hold (the engine's own half of a split) is not in the package
   * the list publishes, so the next publish, which then must go out, takes it back.
   */
  #noteEngine(change: TreeModelChange): void {
    const doc = this.core?.replica.doc;
    const story = change.story?.partName;
    const target = story === undefined ? undefined : doc?.stories.get(story);
    if (!doc || story === undefined || !target) return;
    const ids = [...change.dirty, ...change.deleted];
    for (const entry of change.splitJoin)
      if ('split' in entry) ids.push(entry.split.from);
      else ids.push(entry.join.kept, entry.join.removed);
    if (change.created.length > 0) this.#dirty.stale = true;
    for (const id of ids) {
      const mark = markOfNode(doc, target, id);
      if (mark) this.#dirty.mark(story, { space: mark.space, offset: mark.offset });
      else this.#dirty.stale = true;
    }
  }

  // ---- the gate ----

  /**
   * The editor's ops before it commits them. `context.package` is the package they apply to,
   * with the write's own package edits (a new picture's media part and relationships), which a
   * translation reads only when it needs it.
   */
  gateOperations(
    ops: readonly TreeDocOp[],
    _scope: StoryScope,
    context?: CollaborationGateContext
  ): CollaborationFailureCode | null {
    return this.#gate(ops, null, context);
  }

  /**
   * A review write from the editor (task 9.3, design D15): an accept or a reject is checked by
   * its ops (they become `decide` changes), a new comment by its text and range (a `comment`
   * change). Any other review write (a reply, a resolve, a delete: they name no comment of the
   * room) is refused by its name.
   */
  gateReviewWrite(
    write: CollaborationReviewWrite,
    _scope: StoryScope
  ): CollaborationFailureCode | null {
    if (write.comment !== undefined && write.intent === 'comment-add')
      return this.#gate(commentMarkerOps(write.comment), { commentText: write.comment.text });
    if (write.ops !== undefined && write.ops.length > 0) return this.#gate(write.ops, {});
    const unready = this.#unready();
    if (unready !== null) return unready;
    this.report({ code: 'ordered-unsupported-op', reason: write.intent });
    return 'ordered-unsupported-op';
  }

  /** Why the session takes no edit now, or null. */
  #unready(): CollaborationFailureCode | null {
    if (this.destroyed) return 'collaboration-session-destroyed';
    if (!this.#port) return 'collaboration-session-not-attached';
    if (!this.core || !this.canWrite) return 'collaboration-session-not-ready';
    return null;
  }

  /**
   * A local edit: its ops translated into changes and applied to the list in undo records and
   * queued, before the editor commits; or refused by name, with nothing changed.
   */
  #gate(
    ops: readonly TreeDocOp[],
    review: OrderedReviewContext | null,
    context?: CollaborationGateContext
  ): CollaborationFailureCode | null {
    const unready = this.#unready();
    if (unready !== null) return unready;
    const core = this.core;
    const port = this.#port;
    if (!core || !port) return 'collaboration-session-not-ready';
    const packageWrites = context?.packageWrites;
    if (ops.length === 0 && packageWrites === undefined) {
      this.report({ code: 'ordered-opaque-write', reason: 'opaque-write' });
      return 'ordered-opaque-write';
    }

    this.#settleAdmitted();
    const replica = core.replica;
    const applied: AdmittedChange[] = [];
    const left: OrderedClientRefusalReport[] = [];
    try {
      const batch = translateOps(ops, {
        doc: replica.doc,
        person: replica.person,
        // An indexed lookup; a port without one reads its body paragraphs
        editorText: (nodeId) =>
          port.paragraphTextOf
            ? port.paragraphTextOf(nodeId)
            : (port.paragraphByNodeId(nodeId)?.text ?? null),
        defaultParagraphStyle: core.defaultParagraphStyle,
        shown: () => this.publisher?.current ?? null,
        hyperlinkTarget: (part, id) => port.hyperlinkTarget?.(part, id) ?? null,
        ...(review === null ? {} : { review }),
        // Read when a translation asks: the editor may build the package only then
        gated: () => context?.package ?? null,
        hold: (bytes, contentType) => this.media.hold(bytes, contentType),
        leave: (report) => left.push(report),
        ...(packageWrites === undefined ? {} : { packageWrites }),
      });
      this.#applyBatch(batch, replica, applied);
      this.#admitted = this.#admit(applied, batch, replica.doc, context);

      // What the batch leaves is said by name, and the publish takes the editor's write there back
      for (const report of left) {
        this.report(report);
      }

      if (left.length > 0) {
        this.#dirty.stale = true;
      }

      this.#queueFlush();
      return null;
    } catch (error) {
      // Nothing of the batch stays: the editor does not commit it
      replica.dropLast(applied.length);
      if (!(error instanceof OrderedClientRefusal)) {
        throw error;
      }

      this.report(error.report());
      return error.code;
    }
  }

  /**
   * Apply each step's change bodies to the list, in order, as one user action. `applied` gets
   * each change as it applies, so a throw halfway leaves the count to take back.
   */
  #applyBatch(batch: TranslatedBatch, replica: OrderedReplica, applied: AdmittedChange[]): void {
    let into: string | undefined;
    const action = this.#nextAction++;
    batch.steps.forEach((step, k) => {
      // Each op's changes are made as its own op says (a tracked and a direct op in one batch)
      const mode = batch.modes[k] ?? batch.mode;
      for (const body of step()) {
        const change = this.#acted(
          this.#validated(body, mode, mode === 'suggest' ? into : undefined),
          action
        );
        this.#refuseRetypingBreak(change, replica.doc);
        const { outcome, dirty } = replica.applyLocal(change);
        this.#dirty.add(dirty);
        const own = this.ownPending(outcome.suggestions);
        applied.push({ change, own });

        // The batch's later suggestions add to its first one
        if (mode === 'suggest' && into === undefined) {
          into = own.find((id) => id.startsWith(`S:${change.id}.`)) ?? own[0];
        }
      }
    });
  }

  /** The admitted record of a batch the gate applied, until the editor commits it. */
  #admit(
    applied: readonly AdmittedChange[],
    batch: TranslatedBatch,
    doc: OrderedDocument,
    context: CollaborationGateContext | undefined
  ): Admitted {
    return {
      commit: this.#localCommits + 1,
      changes: applied,
      selection: this.#carry.last,
      caret: caretOfBatch(
        applied.map(({ change }) => change),
        doc
      ),
      direct: batch.modes.includes('edit') && !batch.review,
      group: context?.historyGroup,
    };
  }

  /**
   * A suggesting person's break that retypes the next section, refused before it is sent, with
   * the room's reason: only this one party rule, which reads nothing but the section types.
   */
  #refuseRetypingBreak(change: OrderedChange, doc: OrderedDocument): void {
    if (this.settings.suggests !== true || change.kind !== 'section-break') return;
    const refusal = retypingBreakRefusal(change, doc);
    if (refusal) throw new OrderedClientRefusal('ordered-refused', refusal.reason, refusal.detail);
  }

  #validated(
    body: OrderedChangeBody,
    mode: 'suggest' | 'edit',
    into: string | undefined
  ): OrderedChange {
    const replica = this.#replica();

    // A decision, a take-back or a comment has no mode and adds to no suggestion
    const candidate = isEditBody(body)
      ? { v: 1, id: replica.nextId(), mode, ...(into === undefined ? {} : { into }), ...body }
      : { v: 1, id: replica.nextId(), ...body };
    return this.validate(candidate);
  }

  /**
   * A comment change the host asked for, as one user action: applied to the copy at once,
   * published, and sent in the next batch. It never enters the undo history. Returns the ids of
   * the comments it added, or null when the session takes no edit now or the copy refuses it
   * (reported through `onRefused`). A refusal from the room takes it back later.
   */
  protected writeComment(body: OrderedCommentBody): readonly string[] | null {
    if (this.#unready() !== null) {
      return null;
    }

    this.#settleAdmitted();
    const replica = this.#replica();
    const known = new Set(replica.doc.comments.keys());
    try {
      const change = this.#acted(
        this.validate({ v: 1, id: replica.nextId(), ...body }),
        this.#nextAction++
      );
      this.#dirty.add(replica.applyLocal(change).dirty);
    } catch (error) {
      if (!(error instanceof OrderedClientRefusal)) {
        throw error;
      }

      this.report(error.report());
      return null;
    }

    this.#publish(null);
    this.sendQueued();
    return [...replica.doc.comments.keys()].filter((id) => !known.has(id));
  }

  /** A change as the room's validator reads it; refused by name when it does not. */
  protected validate(candidate: Record<string, unknown>): OrderedChange {
    const replica = this.#replica();
    const checked = validateOrderedChange(candidate, {
      namespace: replica.namespace,
      parts: replica.doc.stories,
    });
    if (!checked.ok)
      throw new OrderedClientRefusal(
        'ordered-refused',
        'invalid-change',
        `${checked.error.code} ${checked.error.path}`
      );
    return checked.value;
  }

  /** The person's own pending suggestions among these. */
  protected ownPending(ids: readonly string[]): string[] {
    const doc = this.#replica().doc;
    const maker = makerOf(this.#replica().person);
    return ids.filter((id) => pendingBy(doc, id, maker) !== null);
  }

  #replica(): OrderedReplica {
    const core = this.core;
    if (!core) throw new Error('The room is not open yet.');
    return core.replica;
  }

  /** An admitted edit the editor committed goes to the history; one it did not, comes back. */
  #settleAdmitted(): void {
    const admitted = this.#admitted;
    const core = this.core;
    if (!admitted || !core) return;
    this.#admitted = null;
    if (this.#localCommits >= admitted.commit) {
      this.history.recordCommit(
        admitted.changes,
        admitted.direct,
        admitted.selection,
        admitted.group
      );
      return;
    }

    // By id: a room message between the gate and now may have refused some of them already
    const ids = new Set(admitted.changes.map(({ change }) => change.id));
    const cycle = core.replica.refuse(ids);
    this.#dirty.add(cycle.dirty);
    for (const refused of cycle.refused) this.reportRefusal(refused);
  }

  // ---- after a commit ----

  /** After the editor's commit, in a microtask: publish what the edit touched and send it. */
  #queueFlush(): void {
    if (this.#flushQueued) return;
    this.#flushQueued = true;
    void Promise.resolve().then(() => {
      if (this.#flushQueued) this.flushPendingJournals();
    });
  }

  /**
   * The list's version of what the editor just committed, before the editor lays it out: the
   * flush queued by the commit still sends. A commit that runs during our publish (a key the
   * host was holding) waits for that flush, never publishing inside the publish.
   */
  settleLocalCommit(): void {
    if (this.destroyed || !this.core || this.#publishing) return;
    const admitted = this.#admitted;
    this.#settleAdmitted();
    if (!this.#dirty.empty) this.#publish(admitted?.caret ?? null);
  }

  /** Publish the list's version of what changed and send what is queued. */
  flushPendingJournals(): void {
    this.#flushQueued = false;
    if (this.destroyed || !this.core) return;
    const admitted = this.#admitted;
    this.#settleAdmitted();
    if (this.#ungated) {
      this.#ungated = false;
      this.report({ code: 'ordered-refused', reason: 'ungated-write' });
    }

    if (!this.#dirty.empty) this.#publish(admitted?.caret ?? null);
    this.sendQueued();
  }

  /** Send the changes not sent yet. */
  protected sendQueued(): void {
    const replica = this.core?.replica;
    if (!replica) return;
    const unsent = replica.takeUnsent((change) => this.mediaReady(change));
    if (unsent.length > 0 && !this.settings.send(unsent)) replica.markAllUnsent();
  }

  /**
   * Whether the room holds the media a change names, so it may go now. A change waits until
   * then, and every later change with it; a session that sends media starts its upload here.
   */
  protected mediaReady(_change: OrderedChange): boolean {
    return true;
  }

  /** Add what a write of the session changed, for the next publish. */
  protected dirty(dirty: OrderedDirty): void {
    this.#dirty.add(dirty);
  }

  /**
   * Give the editor the list's version of what changed. `caret` is where the edit being
   * published leaves the caret, for a selection the editor holds in a paragraph the list does
   * not know.
   */
  protected publishNow(caret: OrderedPlacedPoint | null = null, held?: HeldSelection | null): void {
    // A flush still queued sends what the gate admitted: publishing does not send
    this.#settleAdmitted();
    this.#publish(caret, false, held);
  }

  /**
   * The local selection as places of the list, read before a room's changes apply: then the
   * editor and the list still agree, and each endpoint keeps to its units exactly.
   */
  protected holdSelection(): HeldSelection | null {
    const doc = this.core?.replica.doc;
    const port = this.#port;
    return doc && port ? this.#carry.hold(doc, port, null) : null;
  }

  /** The selection the editor last published, as a range of the list now; null for a caret. */
  protected selectionRange(): OrderedPlacedRange | null {
    const doc = this.core?.replica.doc;
    const port = this.#port;
    return doc && port ? this.#carry.range(doc, port) : null;
  }

  #publish(caret: OrderedPlacedPoint | null, full = false, held?: HeldSelection | null): void {
    const port = this.#port;
    const publisher = this.publisher;
    const doc = this.core?.replica.doc;
    if (!port || !publisher || !doc || (!full && this.#dirty.empty)) return;
    const hold = held === undefined ? this.#carry.hold(doc, port, caret) : held;
    let pkg: OoxmlPackage;
    try {
      pkg = full ? publisher.full() : publisher.next(this.#dirty);
    } catch (error) {
      // A defect while the list's version is composed: told in the status, as a refused
      // publish is. What it names stays named, so the next publish tries it again and the
      // session heals when it goes out; a teardown's flush no longer throws it again
      this.setStatus('error', 'remote-apply-failed', `compose: ${errorName(error)}`);
      return;
    }

    this.#dirty.clear();

    // Where the selection goes is known from the list before the editor takes its version: a
    // host read during the apply lays the pages out, and that layout must find the move
    this.#carry.carry(doc, hold);
    let result = this.#apply(port, pkg);
    if (!result.ok) {
      pkg = publisher.full();
      result = this.#apply(port, pkg);
    }

    if (!result.ok) {
      this.#carry.carry(doc, null);
      this.setStatus('error', 'remote-apply-failed', result.reason.slice(0, 200));
      return;
    }

    if (this.statusState.snapshot().reason?.code === 'remote-apply-failed') {
      this.setStatus(this.reconnecting ? 'disconnected' : 'ready');
    }

    this.notifySelections();
  }

  /** Whether the link is waiting to connect again: a healed publish returns to that status. */
  protected get reconnecting(): boolean {
    return this.statusState.status() === 'disconnected';
  }

  #apply(port: CollaborationDocumentPort, pkg: OoxmlPackage) {
    this.#publishCount += 1;
    this.#publishing = true;
    try {
      return port.applyRemotePackage(pkg, {
        origin: ORIGIN_IDS.mutationRemote,
        actorId: 'ordered-room',
        operationId: `${this.sessionId}:ordered:${this.#publishCount}`,
      });
    } finally {
      this.#publishing = false;
    }
  }

  remoteSelectionMove(): CollaborationSelectionMove | null {
    return this.#carry.move;
  }

  // ---- presence ----

  /** Presence names paragraphs by node id, which every copy shares. */
  setLocalSelection(selection: CollaborationLocalSelection | null): void {
    if (this.destroyed) return;
    this.#carry.track(selection);
    const port = this.#port;
    const doc = this.core?.replica.doc;
    if (!selection || !port || !doc) {
      this.settings.sendPresence(orderedPresenceState(this.identity, null));
      return;
    }

    const point = (address: CollaborationLocalSelection['anchor']) => {
      const node = port.paragraphByStableId(address.paragraphId)?.nodeId;

      // Only a paragraph the list holds: an engine paragraph waits for the publish
      return node !== undefined && shownParagraphOfNode(doc, node) !== null
        ? { paragraph: node, offset: address.offset }
        : null;
    };
    const anchor = point(selection.anchor);
    const head = point(selection.head);
    if (anchor && head)
      this.settings.sendPresence(orderedPresenceState(this.identity, { anchor, head }));
  }

  participants(): readonly CollaborationParticipant[] {
    return this.presence.participants(this.identity);
  }

  subscribeParticipants(
    listener: (participants: readonly CollaborationParticipant[]) => void
  ): () => void {
    this.#participantListeners.add(listener);
    return () => this.#participantListeners.delete(listener);
  }

  remoteSelections(): readonly CollaborationRemoteSelection[] {
    const port = this.#port;
    const doc = this.core?.replica.doc;
    if (!port || !doc) return [];
    return this.presence.selections((nodeId) => this.#presenceParagraph(port, doc, nodeId));
  }

  #presenceParagraph(
    port: CollaborationDocumentPort,
    doc: OrderedDocument,
    nodeId: string
  ): OrderedPresenceParagraph | null {
    for (const story of doc.stories.values()) {
      const mark = markOfNode(doc, story, nodeId);
      const shell = mark?.node;
      if (!shell || shell.kind === 'textValue') continue;
      const paraId = shell.attributes.find((a) => a.localName === 'paraId')?.value;
      const text = port.paragraphTextOf?.(nodeId) ?? null;
      if (paraId === undefined || text === null) return null;
      return { paragraphId: paraId.toUpperCase(), nodeId, text };
    }

    return null;
  }

  subscribeRemoteSelections(
    listener: (selections: readonly CollaborationRemoteSelection[]) => void
  ): () => void {
    this.#selectionListeners.add(listener);
    return () => this.#selectionListeners.delete(listener);
  }

  protected notifySelections(): void {
    if (this.#selectionListeners.size === 0) return;
    const selections = this.remoteSelections();
    for (const listener of [...this.#selectionListeners]) listener(selections);
  }

  protected notifyParticipants(): void {
    const participants = this.participants();
    for (const listener of [...this.#participantListeners]) listener(participants);
  }

  // ---- undo and redo ----

  writeRefusal(): CollaborationFailureCode | null {
    return this.#unready();
  }

  canUndo(): boolean {
    return this.canWrite && this.history.canUndo;
  }

  canRedo(): boolean {
    return this.canWrite && this.history.canRedo;
  }

  historySelection(): CollaborationLocalSelection | null {
    return this.history.restored;
  }

  undo(): boolean {
    return this.#step(undoOwnStep);
  }

  redo(): boolean {
    return this.#step(redoOwnStep);
  }

  /** Run an undo or a redo as one user action, then publish and send what it changed. */
  #step(run: (writer: OrderedHistoryWriter) => boolean): boolean {
    if (!this.canWrite) return false;
    this.flushPendingJournals();
    const action = this.#nextAction++;
    const done = run({
      history: this.history,
      replica: this.#replica(),
      validate: (candidate) => this.#acted(this.validate(candidate), action),
      ownPending: (ids) => this.ownPending(ids),
      written: (dirty) => this.dirty(dirty),
      refused: (report) => this.reportRefusal(report),
    });
    this.publishNow();
    this.sendQueued();
    return done;
  }

  // ---- teardown ----

  /**
   * The flush before the session lets go of the editor. A throw there is told in the status,
   * and the teardown goes on: the host must be able to close the page and join again. Nothing
   * pending is dropped; `unsentChanges` still names it.
   */
  #lastFlush(): void {
    try {
      this.flushPendingJournals();
    } catch (error) {
      this.setStatus('error', 'remote-apply-failed', `flush: ${errorName(error)}`);
    }
  }

  destroy(): void {
    if (this.destroyed) return;
    this.#lastFlush();
    this.#stopPort?.();
    this.#stopPort = null;
    this.#port = null;
    this.destroyed = true;
    this.settings.sendPresence(null);
    this.setStatus('destroyed');
    this.#statusListeners.clear();
    this.#participantListeners.clear();
    this.#selectionListeners.clear();
  }
}
