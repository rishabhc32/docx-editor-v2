import type { OrderedMediaType } from '../model/resources.ts';
import { isOrderedBlobRefusalCode, type OrderedBlobRefusalCode } from '../protocol/blobs.ts';
import { ORDERED_RECONNECT_DELAYS, type OrderedSchedule } from './link.ts';
import { ORDERED_MEDIA_AT_ONCE } from './media-reads.ts';

/** One picture for the room's storage. @public */
export interface OrderedMediaUpload {
  /** `welcome.blobs`, the place media reads come from; the host appends the hash. */
  readonly location: string;
  /** The bytes' SHA-256: 64 lower-case hex digits. */
  readonly sha256: string;
  readonly contentType: OrderedMediaType;
  readonly bytes: Uint8Array;
  /** Aborted when the session ends. */
  readonly signal: AbortSignal;
}

/** The storage's answer to an upload: stored, or refused by name. @public */
export type OrderedMediaUploadResult =
  | { readonly stored: true }
  | { readonly stored: false; readonly code: OrderedBlobRefusalCode };

/**
 * How a host stores a picture the room does not hold yet, at the place `welcome.blobs` names. A
 * rejected promise is a transient failure: the session tries again later.
 *
 * @public
 */
export type OrderedMediaUploader = (
  upload: OrderedMediaUpload
) => Promise<OrderedMediaUploadResult>;

/**
 * Media this copy holds, to upload: what `OrderedMedia.ready` hands its upload callback.
 *
 * @public
 */
export type OrderedHeldBlob = Pick<OrderedMediaUpload, 'sha256' | 'contentType' | 'bytes'>;

/** A blob to upload, and how many tries failed before. */
interface Queued {
  readonly blob: OrderedHeldBlob;
  readonly tries: number;
}

/** A blob whose last try failed, with the timer of its next try. */
interface Failed {
  readonly queued: Queued;
  cancel: (() => void) | null;
}

/** The answer as the session takes it, or null for anything else (a transient failure). */
function answerOf(value: unknown): OrderedMediaUploadResult | null {
  if (typeof value !== 'object' || value === null) {
    return null;
  }

  const answer: { readonly stored?: unknown; readonly code?: unknown } = value;
  if (answer.stored === true) {
    return { stored: true };
  }

  if (answer.stored === false && isOrderedBlobRefusalCode(answer.code)) {
    return { stored: false, code: answer.code };
  }

  return null;
}

/**
 * The uploads of one session (design D14): each blob the host's uploader stores, a few at a
 * time, only to the location `welcome` names. The storage checks the size, the content type and
 * the hash, and stores the bytes under the hash; the room refuses a change whose media it does
 * not hold (`missing-blob`). A failed try goes again after the reconnect delays, and at once
 * after a rejoin. The session ends them all with `stop`.
 */
export class OrderedMediaUploads {
  readonly #upload: OrderedMediaUploader;
  readonly #schedule: OrderedSchedule | null;
  readonly #delays: readonly number[];
  readonly #answered: (sha256: string, result: OrderedMediaUploadResult) => void;
  readonly #abort = new AbortController();
  readonly #queue: Queued[] = [];
  readonly #running = new Set<string>();
  readonly #failed = new Map<string, Failed>();
  #location: string | null = null;

  constructor(
    upload: OrderedMediaUploader,
    schedule: OrderedSchedule | null,
    delays: readonly number[] | undefined,
    answered: (sha256: string, result: OrderedMediaUploadResult) => void
  ) {
    this.#upload = upload;
    this.#schedule = schedule;
    this.#delays = delays !== undefined && delays.length > 0 ? delays : ORDERED_RECONNECT_DELAYS;
    this.#answered = answered;
  }

  /** The room's media location, from `welcome`: failed uploads go again now. */
  connect(location: string | null): void {
    this.#location = location;
    for (const failed of this.#failed.values()) {
      failed.cancel?.();
      this.#queue.push(failed.queued);
    }

    this.#failed.clear();
    this.#pump();
  }

  /** Upload a blob the room does not hold yet. */
  start(blob: OrderedHeldBlob): void {
    this.#queue.push({ blob, tries: 0 });
    this.#pump();
  }

  /** The session ended: abort the uploads in flight and try nothing more. */
  stop(): void {
    this.#abort.abort();
    for (const failed of this.#failed.values()) {
      failed.cancel?.();
    }

    this.#failed.clear();
    this.#queue.length = 0;
  }

  #pump(): void {
    const location = this.#location;
    if (location === null || this.#abort.signal.aborted) {
      return;
    }

    while (this.#running.size < ORDERED_MEDIA_AT_ONCE) {
      const next = this.#queue.shift();
      if (!next) {
        return;
      }

      const { sha256 } = next.blob;
      this.#running.add(sha256);
      void this.#try(next, location).finally(() => {
        this.#running.delete(sha256);
        this.#pump();
      });
    }
  }

  async #try(queued: Queued, location: string): Promise<void> {
    let answer: OrderedMediaUploadResult | null;
    try {
      answer = answerOf(
        await this.#upload({ location, ...queued.blob, signal: this.#abort.signal })
      );
    } catch {
      answer = null;
    }

    if (this.#abort.signal.aborted) {
      return;
    }

    if (answer === null) {
      this.#retry({ blob: queued.blob, tries: queued.tries + 1 });
      return;
    }

    this.#answered(queued.blob.sha256, answer);
  }

  /** Try again after the reconnect delay of this try, or at the next rejoin without timers. */
  #retry(queued: Queued): void {
    const failed: Failed = { queued, cancel: null };
    const { sha256 } = queued.blob;
    this.#failed.set(sha256, failed);
    const delay = this.#delays[Math.min(queued.tries, this.#delays.length) - 1] ?? 0;
    failed.cancel =
      this.#schedule?.(() => {
        if (this.#failed.get(sha256) !== failed) {
          return;
        }

        this.#failed.delete(sha256);
        this.#queue.push(queued);
        this.#pump();
      }, delay) ?? null;
  }
}
