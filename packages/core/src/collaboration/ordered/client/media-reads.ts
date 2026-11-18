import type { OrderedChangeMedia } from '../change/resources.ts';
import { isFileMediaHash } from '../model/file-media.ts';
import { orderedResourceShows } from '../model/resources.ts';
import type { OrderedDocument } from '../model/types.ts';
import type { OrderedMedia } from './media.ts';

/** How a host reads media the room stores: from `location`, by hash. @public */
export type OrderedMediaReader = (location: string, sha256: string) => Promise<Uint8Array>;

/** Reads, or uploads, at once, at most. */
export const ORDERED_MEDIA_AT_ONCE = 4;

/** A read to make: media a change named, or the file's media a checkpoint names. */
interface MediaRead extends OrderedChangeMedia {
  readonly fileMedia: boolean;
}

/** What a read ended with, for the session. */
export type OrderedMediaReadResult =
  | { readonly ok: true; readonly sha256: string }
  | {
      readonly ok: false;
      readonly sha256: string;
      readonly reason: 'missing-blob' | 'invalid-media';
    };

/**
 * The reads of one session (design D14): each picture the copy shows and does not hold, once, a
 * few at a time, only from the location `welcome` names. `OrderedMedia.keep` checks the bytes.
 */
export class OrderedMediaReads {
  readonly #media: OrderedMedia;
  readonly #done: (result: OrderedMediaReadResult) => void;
  readonly #running = new Map<string, Promise<void>>();
  readonly #queue: MediaRead[] = [];
  /** Hashes that did not read: tried again on the next connection. */
  readonly #failed = new Set<string>();
  #reader: OrderedMediaReader | null = null;
  #location: string | null = null;

  constructor(media: OrderedMedia, done: (result: OrderedMediaReadResult) => void) {
    this.#media = media;
    this.#done = done;
  }

  /** The room's media location, from `welcome`, and the host's reader: a new connection. */
  connect(location: string | null, reader: OrderedMediaReader | null): void {
    this.#location = location;
    this.#reader = reader;
    this.#failed.clear();
  }

  /**
   * Read every picture `doc` shows whose bytes the copy does not hold: each of the file's media
   * that a checkpoint names, then media changes added. A change that copies the file's media
   * names the same hash, which the file's read serves.
   */
  readMissing(doc: OrderedDocument): void {
    for (const item of doc.resources.fileMedia.values()) {
      this.#queueRead(item, true);
    }

    for (const resource of doc.resources.added.values()) {
      if (resource.kind !== 'media' || !orderedResourceShows(doc, resource)) {
        continue;
      }

      if (isFileMediaHash(doc, resource.sha256)) {
        continue;
      }

      this.#queueRead(resource, false);
    }

    this.#pump();
  }

  /** Queue a read of media whose bytes the copy does not hold, once. */
  #queueRead(media: OrderedChangeMedia, fileMedia: boolean): void {
    const { sha256 } = media;
    if (this.#media.bytes(sha256) !== undefined || this.#running.has(sha256)) {
      return;
    }

    if (this.#failed.has(sha256) || this.#queue.some((m) => m.sha256 === sha256)) {
      return;
    }

    this.#queue.push({ sha256, contentType: media.contentType, size: media.size, fileMedia });
  }

  /** Every read started or queued so far, settled: a read ending starts the next in the queue. */
  async settled(): Promise<void> {
    while (this.#running.size > 0) {
      await Promise.all(this.#running.values());
    }
  }

  #pump(): void {
    while (this.#running.size < ORDERED_MEDIA_AT_ONCE) {
      const next = this.#queue.shift();
      if (!next) return;
      const read = this.#read(next).finally(() => {
        this.#running.delete(next.sha256);
        this.#pump();
      });
      this.#running.set(next.sha256, read);
    }
  }

  async #read(media: MediaRead): Promise<void> {
    const reader = this.#reader;
    const location = this.#location;
    if (reader === null || location === null) {
      this.#fail(media.sha256, 'missing-blob');
      return;
    }

    let bytes: Uint8Array;
    try {
      bytes = await reader(location, media.sha256);
    } catch {
      this.#fail(media.sha256, 'missing-blob');
      return;
    }

    // A view from any window: the host's reader may run in another one
    const view = ArrayBuffer.isView(bytes)
      ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      : null;
    if (view === null || !this.#media.keep(view, media, media.fileMedia)) {
      this.#fail(media.sha256, 'invalid-media');
      return;
    }

    this.#done({ ok: true, sha256: media.sha256 });
  }

  #fail(sha256: string, reason: 'missing-blob' | 'invalid-media'): void {
    this.#failed.add(sha256);
    this.#done({ ok: false, sha256, reason });
  }
}
