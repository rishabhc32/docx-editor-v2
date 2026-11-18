// `TreeDocOp` translation (design D7 and D13, tasks 6.1 and 9.1). The editor asks the gate
// before it commits a batch of ops; the client turns each op into changes of the change format.

import { HYPERLINK_RELATIONSHIP_TYPE } from '../../../store/package/hyperlink.ts';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { resolveContentTypeOf } from '../../../store/package/package-edit.ts';
import { resolveImageRelationship } from '../../../store/package/relationships.ts';
import type { TreeDocOp } from '../../../store/store/tree-op-types.ts';
import type { CollaborationPackageWrite } from '../../index.ts';
import type { OrderedChangeMedia } from '../change/resources.ts';
import type { OrderedChangeMode } from '../change/types.ts';
import { fileMediaAt } from '../model/file-media.ts';
import type { OrderedPerson } from '../model/person.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { orderedOpEntry } from './ops/index.ts';
import { packageWriteStep } from './ops/package-writes.ts';
import { splitStep } from './ops/paragraphs.ts';
import type {
  OrderedReviewContext,
  OrderedTranslateBatch,
  OrderedTranslateStep,
} from './ops/types.ts';
import {
  joinedParagraphOfNode,
  paragraphOfNode,
  viewText,
  type ListParagraph,
} from './paragraphs.ts';
import { OrderedClientRefusal, type OrderedClientRefusalReport } from './refusal.ts';

export type { OrderedChangeBody, OrderedEditBody } from './ops/types.ts';

/** What a translation reads besides the ops. */
export interface TranslateHost {
  readonly doc: OrderedDocument;
  /** The person the room stamps the batch with. */
  readonly person: OrderedPerson;
  /** The editor's text of a paragraph now, before the commit; null when it shows none. */
  readonly editorText: (nodeId: string) => string | null;
  /** The document's default paragraph style id, for an Enter that names no style for the tail. */
  readonly defaultParagraphStyle: string | null;
  /** The package the editor shows, as the client last published it. */
  readonly shown?: () => OoxmlPackage | null;
  /** The review write the batch is, when the editor gave one (a comment's text). */
  readonly review?: OrderedReviewContext;
  /** A link relationship's address on the editor's package (a link the batch makes). */
  readonly hyperlinkTarget?: (part: string, relationshipId: string) => string | null;
  /** The package the ops apply to, as the editor gave it to the gate. */
  readonly gated?: () => OoxmlPackage | null;
  /** Told what the batch leaves as it is, by name, without refusing it. */
  readonly leave?: (report: OrderedClientRefusalReport) => void;
  /** Hold a new picture's bytes under their hash: the session's media store. */
  readonly hold?: (bytes: Uint8Array, contentType: string) => OrderedChangeMedia;
  /** What the write's package edits do, applied before its ops (`CollaborationPackageWrite`). */
  readonly packageWrites?: readonly CollaborationPackageWrite[];
}

/** A batch the gate admits: how its edits are made, and its steps. */
export interface TranslatedBatch {
  /** How the batch's edits are made; `suggest` when any of them is a suggestion. */
  readonly mode: OrderedChangeMode;
  /** The batch is review actions (decisions, markers), not edits. */
  readonly review: boolean;
  /** One step for each op (or for a split and its mark revision), in order. */
  readonly steps: readonly OrderedTranslateStep[];
  /**
   * Each step's own mode: a batch of tracked and direct ops (an armed format where formatting
   * is not tracked, Word's `w:doNotTrackFormatting`) makes each change as its own op says.
   */
  readonly modes: readonly OrderedChangeMode[];
}

/** An op's own mode, or null when it says none (it takes the batch's). */
function opMode(op: TreeDocOp): OrderedChangeMode | null {
  const tracking = orderedOpEntry(op.op).tracking ?? 'attribution';
  if (tracking === 'review' || tracking === 'either') return null;
  return tracking === 'direct' || !('revision' in op) || op.revision === undefined
    ? 'edit'
    : 'suggest';
}

/**
 * The armed format of typed text (a format chosen before typing, where formatting is not
 * tracked, `w:doNotTrackFormatting`): an untracked format over exactly the text a tracked
 * insertion typed. It falls on the typist's own pending text, which a format changes with no
 * record in either mode, so it takes the insertion's mode (the room's rule allows it too).
 */
function armedFormat(op: TreeDocOp, before: TreeDocOp | undefined): boolean {
  return (
    op.op === 'setRunProperties' &&
    op.revision === undefined &&
    op.targetRunIds === undefined &&
    before?.op === 'insertText' &&
    before.revision !== undefined &&
    before.paragraphId === op.paragraphId &&
    op.start === before.offset &&
    op.end === before.offset + before.text.length
  );
}

/** The batch's review state, and the mode of an op that says none: the first one said. */
function batchMode(ops: readonly TreeDocOp[]): { mode: OrderedChangeMode; review: boolean } {
  let mode: OrderedChangeMode | null = null;
  let review = 0;
  for (const op of ops) {
    if ((orderedOpEntry(op.op).tracking ?? 'attribution') === 'review') review += 1;
    mode ??= opMode(op);
  }

  if (review > 0 && review < ops.length)
    throw new OrderedClientRefusal('ordered-unsupported-op', 'mixed-review');
  return { mode: mode ?? 'edit', review: review > 0 };
}

/** One batch's reads: the paragraphs it names, checked against the editor's text once each. */
class Batch implements OrderedTranslateBatch {
  readonly #host: TranslateHost;
  readonly ops: readonly TreeDocOp[];
  /** Node ids whose editor text was checked once; later reads find them across a join. */
  readonly #known = new Set<string>();

  constructor(host: TranslateHost, ops: readonly TreeDocOp[]) {
    this.#host = host;
    this.ops = ops;
  }

  paragraph(nodeId: string): ListParagraph {
    // A paragraph an earlier step joined into the next one keeps its id in the editor
    if (this.#known.has(nodeId)) return joinedParagraphOfNode(this.#host.doc, nodeId);
    const found = paragraphOfNode(this.#host.doc, nodeId);
    const editor = this.#host.editorText(nodeId);
    if (editor === null || editor !== viewText(found.view))
      throw new OrderedClientRefusal('ordered-refused', 'text-differs', nodeId.slice(0, 160));
    this.#known.add(nodeId);
    return found;
  }

  shown(): OoxmlPackage | null {
    return this.#host.shown?.() ?? null;
  }

  hyperlinkTarget(part: string, relationshipId: string): string | null {
    return this.#host.hyperlinkTarget?.(part, relationshipId) ?? null;
  }

  gated(): OoxmlPackage | null {
    return this.#host.gated?.() ?? null;
  }

  get review(): OrderedReviewContext | null {
    return this.#host.review ?? null;
  }

  media(part: string, relationshipId: string): OrderedChangeMedia | null {
    const pkg = this.gated();
    if (pkg === null) {
      return null;
    }

    // An internal image part only: an external picture is never fetched
    const found = resolveImageRelationship(pkg.relationships.get(part), part, relationshipId);
    if (found.mode !== 'internal') {
      return null;
    }

    const contentType = resolveContentTypeOf(pkg, found.partName);
    if (contentType === null) {
      return null;
    }

    return this.partMedia(found.partName, pkg.partBytes.get(found.partName), contentType);
  }

  partMedia(
    part: string,
    bytes: Uint8Array | undefined,
    contentType: string
  ): OrderedChangeMedia | null {
    // The file's media, unless the editor holds other bytes under its name
    const file = fileMediaAt(this.#host.doc, part);
    const same =
      file !== undefined &&
      (bytes === undefined ||
        (bytes.byteLength === file.size && orderedBlobHash(bytes) === file.sha256));
    if (same) {
      return { sha256: file.sha256, contentType: file.contentType, size: file.size };
    }

    return bytes === undefined ? null : this.heldMedia(bytes, contentType);
  }

  heldMedia(bytes: Uint8Array, contentType: string): OrderedChangeMedia {
    if (!this.#host.hold)
      throw new OrderedClientRefusal('ordered-refused', 'invalid-media', 'no media store');
    return this.#host.hold(bytes, contentType);
  }

  link(part: string, relationshipId: string): string | null {
    const found = this.gated()?.externalTargets.find(
      (target) =>
        target.ownerPart === part &&
        target.id === relationshipId &&
        target.type === HYPERLINK_RELATIONSHIP_TYPE
    );
    return found?.sinkSafe ? found.rawTarget : null;
  }

  leave(report: OrderedClientRefusalReport): void {
    this.#host.leave?.(report);
  }

  get doc(): OrderedDocument {
    return this.#host.doc;
  }

  get person(): OrderedPerson {
    return this.#host.person;
  }

  get defaultParagraphStyle(): string | null {
    return this.#host.defaultParagraphStyle;
  }
}

/** Whether a split's tracked mark revision is the op after it, on the same paragraph. */
function pairsWithSplit(
  split: Extract<TreeDocOp, { readonly op: 'splitParagraph' }>,
  next: TreeDocOp | undefined
): boolean {
  return (
    next?.op === 'setParagraphMarkRevision' &&
    next.kind === 'ins' &&
    next.paragraphId === split.paragraphId
  );
}

/**
 * Translate one batch of the editor's ops into steps, or refuse it by name. Every op is checked
 * for support before any step runs, so a refusal leaves the list as it was.
 *
 * @throws OrderedClientRefusal for an op the format does not say.
 */
export function translateOps(ops: readonly TreeDocOp[], host: TranslateHost): TranslatedBatch {
  const batch = new Batch(host, ops);

  // The package edits apply before the ops, as the store applies them
  const writes = (host.packageWrites ?? []).map((write) => packageWriteStep(write, batch));
  const steps: OrderedTranslateStep[] = [...writes];
  const own: (OrderedChangeMode | null)[] = writes.map(() => 'edit');
  for (let k = 0; k < ops.length; k += 1) {
    const op = ops[k];
    if (op === undefined) {
      continue;
    }

    if (op.op === 'splitParagraph' && pairsWithSplit(op, ops[k + 1])) {
      steps.push(splitStep(batch, op.paragraphId, op.offset, op.tailStyleId));
      own.push(opMode(ops[k + 1] ?? op));
      k += 1;
      continue;
    }

    const entry = orderedOpEntry(op.op);

    // A planned op is refused by its name until its group builds it (design D13)
    if (entry.status !== 'translated')
      throw new OrderedClientRefusal('ordered-unsupported-op', op.op);
    steps.push(entry.translate(op, batch));
    own.push(armedFormat(op, ops[k - 1]) ? 'suggest' : opMode(op));
  }

  // Every op is one the format says before the batch's tracking is read
  const { mode, review } = batchMode(ops);
  const modes = own.map((m) => m ?? mode);
  return { mode: modes.includes('suggest') ? 'suggest' : mode, review, steps, modes };
}
