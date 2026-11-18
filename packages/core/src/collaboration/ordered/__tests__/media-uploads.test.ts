// Media bytes through the room's storage (design D14): the host uploads a picture over its own
// link, the storage stores it by SHA-256 only when the bytes are what the host named, and the
// protocol carries no media frames.

import { describe, expect, test } from 'bun:test';
import { OrderedLocalRoom } from '../client/local-room.ts';
import { OrderedMedia } from '../client/media.ts';
import {
  OrderedMediaUploads,
  type OrderedMediaUpload,
  type OrderedMediaUploadResult,
} from '../client/media-uploads.ts';
import type { OrderedSchedule } from '../client/link.ts';
import { ORDERED_RESOURCE_LIMITS } from '../change/resources.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  decodeOrderedClientMessage,
  decodeOrderedServerMessage,
  encodeOrderedServerMessage,
} from '../protocol/codec.ts';
import type { OrderedServerMessage } from '../protocol/messages.ts';
import { APPLY_FILE } from './change-apply-fixture.ts';
import { RED_PNG } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const LOCATION = '/ordered/room-1/blobs/';
const LOGIN = {
  person: testPerson('Ben'),
  suggests: true,
  decides: false,
  readOnly: false,
  comments: true,
};

/** A real PNG with one more byte at the end: other bytes, another hash. */
const png = (seed: number) => new Uint8Array([...RED_PNG, seed]);

const blobOf = (bytes: Uint8Array) =>
  ({ sha256: orderedBlobHash(bytes), contentType: 'image/png', bytes }) as const;

const localRoom = () =>
  new OrderedLocalRoom({ room: 'room-1', wordFile: APPLY_FILE, file: 1, code: 'test-code' });

/** Timers the test runs itself. */
function timers() {
  const queued: { run: () => void; delayMs: number; cancelled: boolean }[] = [];
  const schedule: OrderedSchedule = (run, delayMs) => {
    const timer = { run, delayMs, cancelled: false };
    queued.push(timer);
    return () => {
      timer.cancelled = true;
    };
  };
  const runAll = () => {
    for (const timer of queued.splice(0)) {
      if (!timer.cancelled) {
        timer.run();
      }
    }
  };
  return { queued, schedule, runAll };
}

/** An uploader whose answers the test gives, one call at a time. */
function heldUploader() {
  const calls: {
    readonly upload: OrderedMediaUpload;
    readonly answer: (result: OrderedMediaUploadResult) => void;
    readonly fail: () => void;
  }[] = [];
  const upload = (request: OrderedMediaUpload) =>
    new Promise<OrderedMediaUploadResult>((resolve, reject) => {
      calls.push({ upload: request, answer: resolve, fail: () => reject(new Error('offline')) });
    });
  return { calls, upload };
}

const settle = async () => {
  for (let k = 0; k < 6; k += 1) {
    await Promise.resolve();
  }
};

describe('the protocol carries no media', () => {
  test('a blob frame from a client, and a blob answer from the room, are unknown types', () => {
    const client = decodeOrderedClientMessage(
      JSON.stringify({
        v: 3,
        type: 'blob',
        sha256: 'a'.repeat(64),
        contentType: 'image/png',
        size: 1,
        offset: 0,
        data: 'AA==',
      })
    );
    expect(client.ok ? 'ok' : client.error.code).toBe('unknown-type');
    for (const type of ['blob-stored', 'blob-refused']) {
      const server = decodeOrderedServerMessage(
        JSON.stringify({ v: 3, type, sha256: 'a'.repeat(64), code: 'blob-hash' })
      );
      expect(server.ok ? 'ok' : server.error.code).toBe('unknown-type');
    }
  });

  test('the welcome names where media is read and stored, as a safe location', () => {
    const welcome: OrderedServerMessage = {
      v: 3,
      type: 'welcome',
      file: 1,
      location: '/ordered/r/1.docx',
      n: 0,
      saved: 0,
      changes: [],
      blobs: '/ordered/r/blobs/',
    };
    const decoded = decodeOrderedServerMessage(encodeOrderedServerMessage(welcome));
    expect(decoded.ok && decoded.value).toEqual(welcome);
    const unsafe = decodeOrderedServerMessage(
      JSON.stringify({ ...welcome, blobs: 'javascript:alert(1)' })
    );
    expect(unsafe.ok ? 'ok' : unsafe.error.code).toBe('invalid-href');
  });
});

describe("the local room's storage", () => {
  test('stores media by hash only when the bytes are what the host named', async () => {
    const room = localRoom();
    const upload = room.uploader(LOGIN);
    const signal = new AbortController().signal;
    const good = png(1);
    expect(await upload({ location: LOCATION, ...blobOf(good), signal })).toEqual({
      stored: true,
    });
    expect(room.blob(orderedBlobHash(good))).toEqual(good);

    // Bytes whose hash is not the one named
    const named = { ...blobOf(png(2)), sha256: 'b'.repeat(64) };
    expect(await upload({ location: LOCATION, ...named, signal })).toEqual({
      stored: false,
      code: 'blob-hash',
    });
    expect(room.blob('b'.repeat(64))).toBeUndefined();

    // Bytes that are not the content type named
    const text = new TextEncoder().encode('<svg onload="x"/>');
    expect(await upload({ location: LOCATION, ...blobOf(text), signal })).toEqual({
      stored: false,
      code: 'blob-type',
    });

    // Bytes over the media cap
    const large = new Uint8Array(ORDERED_RESOURCE_LIMITS.mediaSize + 1);
    large.set(RED_PNG);
    expect(await upload({ location: LOCATION, ...blobOf(large), signal })).toEqual({
      stored: false,
      code: 'blob-too-large',
    });
  });

  test("checks an embedded object's part and an ink part by signature", async () => {
    const upload = localRoom().uploader(LOGIN);
    const signal = new AbortController().signal;
    const compound = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    const ole = 'application/vnd.openxmlformats-officedocument.oleObject';
    const ink = new TextEncoder().encode('<inkml:ink/>');
    const take = (bytes: Uint8Array, contentType: typeof ole | 'application/inkml+xml') =>
      upload({ location: LOCATION, sha256: orderedBlobHash(bytes), contentType, bytes, signal });
    expect(await take(compound, ole)).toEqual({ stored: true });
    expect(await take(RED_PNG, ole)).toEqual({ stored: false, code: 'blob-type' });
    expect(await take(ink, 'application/inkml+xml')).toEqual({ stored: true });
    expect(await take(compound, 'application/inkml+xml')).toEqual({
      stored: false,
      code: 'blob-type',
    });
  });

  test('a read-only login, or another location, fails as a transient fault', async () => {
    const room = localRoom();
    const signal = new AbortController().signal;
    const blob = blobOf(png(3));
    const readOnly = room.uploader({ ...LOGIN, readOnly: true });
    await expect(readOnly({ location: LOCATION, ...blob, signal })).rejects.toThrow();
    const other = room.uploader(LOGIN);
    await expect(other({ location: '/ordered/room-2/blobs/', ...blob, signal })).rejects.toThrow();
    expect(room.blob(blob.sha256)).toBeUndefined();
  });
});

describe("a session's media and uploads", () => {
  test('a blob is ready once stored; each one starts one upload; a refused one stays not ready', () => {
    const media = new OrderedMedia();
    const first = media.hold(png(1), 'image/png');
    const second = media.hold(png(2), 'image/png');
    const started: string[] = [];
    const start = (blob: { readonly sha256: string }) => started.push(blob.sha256);
    expect(media.ready([first.sha256, second.sha256], start)).toBe(false);
    expect(media.ready([first.sha256, second.sha256], start)).toBe(false);
    expect(started).toEqual([first.sha256, second.sha256]);
    media.markStored(first.sha256);
    media.markRefused(second.sha256, 'blob-quota');
    expect(media.ready([first.sha256], start)).toBe(true);
    expect(media.ready([second.sha256], start)).toBe(false);
    expect(media.refused(second.sha256)).toBe('blob-quota');
    expect(started).toHaveLength(2);

    // Media this copy does not hold has nothing to upload
    expect(media.ready(['c'.repeat(64)], start)).toBe(false);
    expect(started).toHaveLength(2);
  });

  test('uploads run four at a time, to the location the welcome names', async () => {
    const { calls, upload } = heldUploader();
    const answers: [string, OrderedMediaUploadResult][] = [];
    const uploads = new OrderedMediaUploads(upload, null, undefined, (sha256, result) =>
      answers.push([sha256, result])
    );
    const blobs = Array.from({ length: 6 }, (_, k) => blobOf(png(k)));
    for (const blob of blobs) {
      uploads.start(blob);
    }

    // No location yet: nothing goes
    expect(calls).toHaveLength(0);
    uploads.connect(LOCATION);
    expect(calls.map((call) => call.upload.sha256)).toEqual(blobs.slice(0, 4).map((b) => b.sha256));
    expect(calls.every((call) => call.upload.location === LOCATION)).toBe(true);
    calls[0]?.answer({ stored: true });
    calls[1]?.answer({ stored: false, code: 'blob-quota' });
    await settle();
    expect(calls).toHaveLength(6);
    expect(answers).toEqual([
      [blobs[0]?.sha256 ?? '', { stored: true }],
      [blobs[1]?.sha256 ?? '', { stored: false, code: 'blob-quota' }],
    ]);
  });

  test('a failed upload goes again after the reconnect delays, and at once after a rejoin', async () => {
    const { calls, upload } = heldUploader();
    const clock = timers();
    const answers: string[] = [];
    const uploads = new OrderedMediaUploads(upload, clock.schedule, [100, 400], (sha256) =>
      answers.push(sha256)
    );
    uploads.connect(LOCATION);
    const blob = blobOf(png(1));
    uploads.start(blob);
    for (const delay of [100, 400, 400]) {
      calls.at(-1)?.fail();
      await settle();
      expect(clock.queued.map((timer) => timer.delayMs)).toEqual([delay]);
      clock.runAll();
    }

    // A rejoin tries again at once, and cancels the wait
    calls.at(-1)?.fail();
    await settle();
    expect(calls).toHaveLength(4);
    const wait = clock.queued[0];
    uploads.connect(LOCATION);
    expect(calls).toHaveLength(5);
    expect(wait?.cancelled).toBe(true);

    // An answer that is not one the storage gives is a transient failure too
    calls.at(-1)?.answer({ stored: false, code: 'made-up' } as unknown as OrderedMediaUploadResult);
    await settle();
    expect(answers).toEqual([]);
    clock.runAll();
    calls.at(-1)?.answer({ stored: true });
    await settle();
    expect(answers).toEqual([blob.sha256]);
  });

  test('stop aborts the uploads in flight and tries nothing more', async () => {
    const { calls, upload } = heldUploader();
    const clock = timers();
    const answers: string[] = [];
    const uploads = new OrderedMediaUploads(upload, clock.schedule, undefined, (sha256) =>
      answers.push(sha256)
    );
    uploads.connect(LOCATION);
    uploads.start(blobOf(png(1)));
    uploads.start(blobOf(png(2)));
    calls[1]?.fail();
    await settle();
    expect(clock.queued).toHaveLength(1);
    uploads.stop();
    expect(calls[0]?.upload.signal.aborted).toBe(true);
    expect(clock.queued[0]?.cancelled).toBe(true);
    calls[0]?.answer({ stored: true });
    await settle();
    uploads.connect(LOCATION);
    uploads.start(blobOf(png(3)));
    expect(answers).toEqual([]);
    expect(calls).toHaveLength(2);
  });
});
