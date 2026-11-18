/**
 * `@docx-editor.dev/core/collaboration` — provider-neutral collaboration contracts.
 *
 * This lane names the attachment between the canonical tree and an optional replication
 * implementation. It performs no networking and imports no CRDT.
 *
 * @packageDocumentation
 * @public
 */

import type { HistoryGroup } from '../store/store/history-group.ts';
import type { StoryScope } from '../store/store/tree-package-store.ts';
import type { TreeDocOp } from '../store/store/tree-op-types.ts';
import type { OoxmlPackage } from '../store/package/ooxml-package.ts';
import type { CollaborationDocumentPort } from './replication.ts';

export { safeParticipantColor } from './participant-color.ts';
export {
  createCollaborationStatusTracker,
  isCollaborationFailureCode,
  type CollaborationFailure,
  type CollaborationFailureCode,
  type CollaborationStatus,
  type CollaborationStatusSnapshot,
} from './failure.ts';
import type {
  CollaborationFailureCode,
  CollaborationStatus,
  CollaborationStatusSnapshot,
} from './failure.ts';

/** Human or automation identity attached to authored collaboration transactions. @public */
export interface CollaborationIdentity {
  /** Stable author id, 1 to 256 characters after trimming. Recorded on authored changes. */
  readonly actorId: string;
  /** Display name, 1 to 256 characters after trimming. */
  readonly name: string;
  /** Presence color, at most 64 characters. An unsafe CSS color falls back to the accent. */
  readonly color?: string;
  /** `agent` for automation. Defaults to `human`. */
  readonly role?: 'human' | 'agent';
}

/** One validated identity visible through ephemeral collaboration presence. @public */
export interface CollaborationParticipant extends CollaborationIdentity {
  /** True for this replica's own participant. */
  readonly isLocal: boolean;
}

/**
 * One endpoint of a published selection: a stable paragraph id and a UTF-16 offset.
 *
 * The wire payload carries only these two endpoints. The receiver walks its own canonical
 * tree to find the paragraphs between them, so a select-all does not grow with document size.
 *
 * @public
 */
export interface CollaborationSelectionAddress {
  /** The stable `w14:paraId` of the paragraph. */
  readonly paragraphId: string;
  /** A UTF-16 offset in the paragraph's text. */
  readonly offset: number;
}

/**
 * How a published selection covers the document.
 *
 * Absent or omitted means a character range. `cells` means the table rectangle whose
 * corner cells contain the two endpoints. The payload still carries only those endpoints,
 * so a large table selection does not grow with the number of selected cells.
 *
 * @public
 */
export type CollaborationSelectionKind = 'cells';

/**
 * One endpoint of a remote selection resolved into this replica's canonical addresses.
 *
 * `paragraphId` is the stable `w14:paraId`. `nodeId` is replica-local and is used to paint.
 *
 * @public
 */
export interface CollaborationRemoteSelectionAddress {
  /** The stable `w14:paraId` of the paragraph. */
  readonly paragraphId: string;
  /** The paragraph's node id in this replica's editor tree. */
  readonly nodeId: string;
  /** A UTF-16 offset in the paragraph's text. */
  readonly offset: number;
}

/**
 * Stable remote selection resolved into this replica's canonical paragraph addresses.
 *
 * Anchor and head may name different paragraphs. A collapsed caret is the same address twice.
 *
 * @public
 */
export interface CollaborationRemoteSelection {
  /** The participant's `actorId`. */
  readonly actorId: string;
  /** The participant's display name. */
  readonly name: string;
  /** The participant's presence color, when it published a safe one. */
  readonly color?: string;
  /** Where the selection starts. */
  readonly anchor: CollaborationRemoteSelectionAddress;
  /** Where the selection ends: the remote caret. */
  readonly head: CollaborationRemoteSelectionAddress;
  /** `cells` for a table cell rectangle; absent for a character range. */
  readonly kind?: CollaborationSelectionKind;
}

/**
 * Selection published by the local editor through ephemeral awareness.
 *
 * Publish an anchor address and a head address. Do not materialize every covered paragraph.
 *
 * @public
 */
export interface CollaborationLocalSelection {
  /** Where the selection starts. */
  readonly anchor: CollaborationSelectionAddress;
  /** Where the selection ends: the caret. */
  readonly head: CollaborationSelectionAddress;
  /** `cells` for a table cell rectangle; absent for a character range. */
  readonly kind?: CollaborationSelectionKind;
}

/**
 * Where a collaboration session carried the local selection across the last remote change.
 *
 * The editor uses `to` only while its selection still equals `from`: a selection the user
 * moved since is the user's. @public
 */
export interface CollaborationSelectionMove {
  /** The selection the session carried, as the editor last published it. */
  readonly from: CollaborationEditorSelection;
  /** Where it stands now: beside the same characters, wherever they show. */
  readonly to: CollaborationEditorSelection;
}

/** A selection in this replica's editor. @public */
export interface CollaborationEditorSelection {
  /** Where the selection starts: it stays put when the selection extends. */
  readonly anchor: CollaborationEditorPosition;
  /** Where the selection ends: the caret. */
  readonly head: CollaborationEditorPosition;
}

/**
 * A position in this replica's editor tree. Unlike a stable paragraph id, which a remote
 * change can rename on one replica, a node id names one paragraph of this editor. @public
 */
export interface CollaborationEditorPosition {
  /** The node id of the paragraph in this replica's editor tree. */
  readonly nodeId: string;
  /** A UTF-16 offset in the paragraph's text. */
  readonly offset: number;
}

/**
 * Which review write the editor commits: an accept or a reject, a change of suggestions'
 * author, a comment write, or a package-scoped write (such as a custom node). @public
 */
export type CollaborationReviewWriteIntent =
  | 'revision-resolve'
  | 'revision-attribution'
  | 'comment-add'
  | 'comment-reply'
  | 'comment-resolve'
  | 'comment-delete'
  | 'package-scoped';

/**
 * A review write as the editor is about to commit it, for a session that checks review writes
 * by what they do (`gateReviewWrite`).
 *
 * @public
 */
export interface CollaborationReviewWrite {
  /** Which review write it is. */
  readonly intent: CollaborationReviewWriteIntent;
  /** The ops the write commits, when the editor plans them before the commit (an accept). */
  readonly ops?: readonly TreeDocOp[];
  /**
   * A new comment: its text and its range, in this replica's paragraph node ids. The store
   * mints the comment's id and markers when it commits.
   */
  readonly comment?: {
    readonly text: string;
    readonly paragraphId: string;
    readonly start: number;
    readonly end: number;
    readonly endParagraphId?: string;
  };
}

/**
 * What one of a write's package edits does, for a session that carries a package edit by its
 * meaning rather than by the parts it rewrites. A package edit with no such description is
 * opaque to it.
 *
 * - `document-properties`: core document properties written (`docProps/core.xml`), by name.
 * - `bookmark-removal`: a bookmark's markers taken out of a story part the write's ops do not
 *   address.
 *
 * @public
 */
export type CollaborationPackageWrite =
  | {
      readonly kind: 'document-properties';
      readonly values: Readonly<Record<string, string>>;
    }
  | { readonly kind: 'bookmark-removal'; readonly partName: string; readonly name: string };

/**
 * What the editor knows about a write besides its ops, for a session that reads what the
 * write adds outside the story text.
 *
 * A rich paste merges styles, numbering, media, relationships and notes into the package in
 * the same transaction as its `insertFragment`, and a list or a link can create the style it
 * needs. Those package edits are not ops, so a session that sees only the ops cannot tell what
 * the blocks name. The context carries the package that holds them.
 *
 * @public
 */
export interface CollaborationGateContext {
  /**
   * The package the ops apply to, after the write's own package edits (a paste's merged
   * styles, numbering, media and relationships; a style a list or a link needs). The ops'
   * blocks name those parts by the ids this package holds.
   *
   * Absent when the editor has no package at hand for the write. The editor may compute it
   * only when it is read, so read it only when the session needs it.
   */
  readonly package?: OoxmlPackage;
  /**
   * The editor's undo group of the write. The writes of one command that commit apart (a
   * table of contents' rows, then its page numbers, which need the layout the rows made)
   * carry the same group, and the editor alone undoes them as one step. A session that keeps
   * its own undo steps joins them the same way. Absent when the write has no group.
   */
  readonly historyGroup?: HistoryGroup;
  /**
   * What the write's package edits do, in order. Present only when every one of them says what
   * it does (`CollaborationPackageWrite`); a session that carries package edits by meaning
   * refuses a write with an edit that does not.
   */
  readonly packageWrites?: readonly CollaborationPackageWrite[];
}

/**
 * Optional replication session attached to an editor or headless automation host.
 *
 * Implementations own replication state only. The attached document port remains the authored
 * authority for reads, layout, paint, and save.
 *
 * @public
 */
export interface EditorCollaborationSession {
  /** The room's document id. An attached port must carry the same id. */
  readonly documentId: string;
  /** Unique identity for this attachment lifetime. It prevents operation ID reuse after reconnect. */
  readonly sessionId: string;
  /** The validated local identity. */
  readonly identity: CollaborationIdentity;
  /** Current lifecycle state. */
  status(): CollaborationStatus;
  /**
   * Cached status, current reason, and last failure.
   *
   * Same reference until any of those values change. A host that mounts after a
   * recovered error still reads {@link CollaborationStatusSnapshot.lastFailure}.
   */
  statusSnapshot(): CollaborationStatusSnapshot;
  /** Call `listener` on each status or reason change. Returns the unsubscribe function. */
  subscribeStatus(
    listener: (
      status: CollaborationStatus,
      reason?: CollaborationFailureCode,
      detail?: string
    ) => void
  ): () => void;
  /**
   * Attach the editor's document port and start replicating. Returns the detach function.
   * A second, different port moves the session to `error` with `port-already-attached`.
   */
  attach(port: CollaborationDocumentPort): () => void;
  /**
   * Whether an editor has attached its document port to this replica.
   *
   * False means edits do not replicate, whatever `status()` says — the usual cause is a host
   * that did not remount the editor when the session appeared, so `collaborationModule`
   * never attached. That mistake used to be reachable only as a `console.warn`, which a
   * production build discards; this is the same fact as a value a host can render.
   */
  readonly attached: boolean;
  /**
   * Check whether `ops` may commit now. Returns `null` to admit them, or the refusal code.
   * Call it before every store transaction a custom port commits.
   *
   * `context` tells what the write adds to the package besides the ops. The editor passes it
   * for a rich paste, which it asks inside the store transaction once the clipboard's
   * resources have merged and before the blocks land, and for a write that carries package
   * edits. A session that needs only the ops ignores it.
   */
  gateOperations(
    ops: readonly TreeDocOp[],
    scope: StoryScope,
    context?: CollaborationGateContext
  ): CollaborationFailureCode | null;
  /**
   * Whether the editing API may send this session a tracked write that changes text at several
   * places in one batch: a tracked replacement across paragraphs, a tracked table or cell value
   * write, and a tracked content control around existing text. The session's gate then admits
   * or refuses each write by what it holds. Optional; without it the editing API refuses these
   * writes in collaboration (`unsupported-capability`), because a session that places each
   * piece against concurrent edits by itself cannot keep them together.
   */
  readonly trackedRangeReplacement?: boolean;
  /**
   * Check a review write (an accept, a reject, a comment) before it commits, by what it does.
   * Returns `null` to admit it, or the refusal code. Optional: without it the editor asks
   * `gateOperations([], scope)`, which answers readiness only.
   */
  gateReviewWrite?(
    write: CollaborationReviewWrite,
    scope: StoryScope
  ): CollaborationFailureCode | null;
  /**
   * Why the session takes no write now, such as `collaboration-session-not-ready` while it
   * joins, waits, or is offline with offline editing turned off. Returns `null` when it takes
   * writes. Reports nothing and changes nothing. Optional; with it the editor gives this
   * reason for an undo or redo that the session cannot run, rather than "nothing to undo".
   */
  writeRefusal?(): CollaborationFailureCode | null;
  /** Whether the shared undo history holds a local step to undo. */
  canUndo(): boolean;
  /** Whether the shared undo history holds a local step to redo. */
  canRedo(): boolean;
  /** Undo this participant's last step; others' edits stay. Returns false when refused. */
  undo(): boolean;
  /** Redo this participant's last undone step. Returns false when refused. */
  redo(): boolean;
  /**
   * The selection the last `undo` or `redo` restored: the one its author had before the
   * undone change. Null when the step recorded none. Optional; a session without it leaves
   * the editor to place the caret at the change.
   */
  historySelection?(): CollaborationLocalSelection | null;
  /**
   * Where the last remote change carried the local selection, on the shared characters
   * beside its endpoints, and the selection it was carried from, both in the editor's
   * paragraph node ids. Null when the session could not carry it. Optional; without it the
   * editor maps the selection by text.
   */
  remoteSelectionMove?(): CollaborationSelectionMove | null;
  /** Publish the local selection through presence. `null` clears it. */
  setLocalSelection(selection: CollaborationLocalSelection | null): void;
  /** Participants in presence, this replica included. Capped at 256. */
  participants(): readonly CollaborationParticipant[];
  /** Call `listener` when the participant list changes. Returns the unsubscribe function. */
  subscribeParticipants(
    listener: (participants: readonly CollaborationParticipant[]) => void
  ): () => void;
  /** Other participants' selections, resolved into this replica's addresses. */
  remoteSelections(): readonly CollaborationRemoteSelection[];
  /** Call `listener` when a remote selection changes. Returns the unsubscribe function. */
  subscribeRemoteSelections(
    listener: (selections: readonly CollaborationRemoteSelection[]) => void
  ): () => void;
  /**
   * Give the editor this session's own version of the local edit the editor just committed,
   * before the editor lays that edit out, so one layout pass shows both. The editor calls it
   * at the end of each command it commits. Sending to shared state may still wait for
   * `flushPendingJournals`. Optional: a session whose shared state is the editor's own commit
   * has nothing to give.
   */
  settleLocalCommit?(): void;
  /**
   * Publish queued local journals to shared state.
   *
   * Local typing does not wait for replication. Undo, destroy, and page hide call this
   * so a queued journal is never dropped.
   */
  flushPendingJournals(): void;
  /** Stop replicating and release listeners. The status becomes `destroyed`. */
  destroy(): void;
}

/**
 * What a collaboration module contributes: the replica the surface attaches.
 *
 * @public
 */
export interface CollaborationModuleContribution {
  /**
   * A ready session. The host creates the Yjs room, then wraps it with
   * `collaborationModule({ session })`.
   *
   * Core does not build a session from a document id: the opened package has
   * no collaboration room identity. That identity is chosen before mount.
   */
  readonly session: EditorCollaborationSession;
}
