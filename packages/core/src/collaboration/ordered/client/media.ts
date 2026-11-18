import { orderedMediaCap, type OrderedChangeMedia } from '../change/resources.ts';
import { orderedMediaType, type OrderedMediaType } from '../model/resources.ts';
import { orderedBlobHash, type OrderedBlobRefusalCode } from '../protocol/blobs.ts';
import { isOrderedMediaBytes } from './media-check.ts';
import type { OrderedHeldBlob } from './media-uploads.ts';
import { OrderedClientRefusal } from './refusal.ts';

/** Where the upload of one blob stands. */
type Upload = 'uploading' | 'stored' | { readonly refused: OrderedBlobRefusalCode };

/**
 * Media by SHA-256, and where each upload stands. One for each session. A change never carries
 * bytes: it names media by hash. Every byte is untrusted, the editor's and the room's alike.
 *
 * @public
 */
export class OrderedMedia {
  readonly #bytes = new Map<
    string,
    { readonly bytes: Uint8Array; readonly type: OrderedMediaType }
  >();
  readonly #uploads = new Map<string, Upload>();

  /**
   * Hold bytes the editor gives for a new picture, checked by content, and return the change's
   * media value. Refused by name when the bytes are over the room's cap for their type, or are
   * not an image of `contentType` the room stores.
   *
   * @throws OrderedClientRefusal `ordered-refused` with `media-too-large` (detail the type) or
   * `invalid-media`.
   */
  hold(bytes: Uint8Array, contentType: string): OrderedChangeMedia {
    const type = orderedMediaType(contentType);
    if (type === null) {
      throw new OrderedClientRefusal('ordered-refused', 'invalid-media', contentType.slice(0, 64));
    }

    if (bytes.byteLength > orderedMediaCap(type)) {
      throw new OrderedClientRefusal('ordered-refused', 'media-too-large', type);
    }

    if (!isOrderedMediaBytes(bytes, type)) {
      throw new OrderedClientRefusal('ordered-refused', 'invalid-media', contentType.slice(0, 64));
    }

    const sha256 = orderedBlobHash(bytes);

    // A copy of its own: the caller's array may change after the edit
    if (!this.#bytes.has(sha256)) this.#bytes.set(sha256, { bytes: bytes.slice(), type });
    return { sha256, contentType: type, size: bytes.byteLength };
  }

  /**
   * Keep bytes read back from the room for media a change named, or for the file's media a
   * checkpoint names (`fileMedia`). False when they are not the bytes it named: another hash,
   * another size, or not an image of its type. The file's own media is not checked by content:
   * it is kept as the file had it.
   */
  keep(bytes: Uint8Array, media: OrderedChangeMedia, fileMedia = false): boolean {
    if (bytes.byteLength !== media.size) {
      return false;
    }

    if (!fileMedia && !isOrderedMediaBytes(bytes, media.contentType)) {
      return false;
    }

    if (orderedBlobHash(bytes) !== media.sha256) return false;
    this.#bytes.set(media.sha256, { bytes, type: media.contentType });

    // The room's storage holds what it served
    this.#uploads.set(media.sha256, 'stored');
    return true;
  }

  /** The bytes of media this copy holds, or undefined. */
  bytes(sha256: string): Uint8Array | undefined {
    return this.#bytes.get(sha256)?.bytes;
  }

  /** Whether the storage holds the blob. */
  stored(sha256: string): boolean {
    return this.#uploads.get(sha256) === 'stored';
  }

  /** The storage's refusal of the blob, or null. */
  refused(sha256: string): OrderedBlobRefusalCode | null {
    const state = this.#uploads.get(sha256);
    return typeof state === 'object' ? state.refused : null;
  }

  /**
   * Whether every blob in `hashes` is stored, starting the host's upload of each one that is not
   * and is not uploading. A blob this copy does not hold or the storage refused stays not ready.
   * The session holds the first change that names a blob, and every later one, until the upload
   * says it is stored.
   */
  ready(hashes: readonly string[], upload: (blob: OrderedHeldBlob) => void): boolean {
    let ready = true;
    for (const sha256 of hashes) {
      const state = this.#uploads.get(sha256);
      if (state === 'stored') {
        continue;
      }

      ready = false;
      const held = this.#bytes.get(sha256);
      if (state !== undefined || !held) {
        continue;
      }

      this.#uploads.set(sha256, 'uploading');
      upload({ sha256, contentType: held.type, bytes: held.bytes });
    }

    return ready;
  }

  /** The storage holds the blob: a change may name it now. */
  markStored(sha256: string): void {
    this.#uploads.set(sha256, 'stored');
  }

  /** The storage refused the blob: no change may name it. */
  markRefused(sha256: string, code: OrderedBlobRefusalCode): void {
    this.#uploads.set(sha256, { refused: code });
  }
}
