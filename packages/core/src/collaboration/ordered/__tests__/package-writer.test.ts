// The streaming package writer: the same bytes as `orderedPackageBytes`, written front to back,
// with each added picture read through `media` when its entry is written, one at a time.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { unzipSync } from 'fflate';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes, writeOrderedPackage } from '../compose/package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { apply, at, EDIT } from './change-apply-fixture.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, t } from './word-fixtures.ts';

const FIXTURES = join(import.meta.dir, '../../../../../../e2e/fixtures');

/** Repository files with pictures of their own, in the body, headers and sections. */
const FILES = [
  'images-formats.docx',
  'images-tiff.docx',
  'header-vml-emf.docx',
  'issue-319-sections.docx',
  'demo.docx',
];

interface Opened {
  readonly doc: OrderedDocument;
  readonly base: OoxmlPackage;
}

const opened = (bytes: Uint8Array): Opened => {
  const file = openOrderedWordFile(bytes, { file: 1 });
  return { doc: file.document, base: file.base };
};

/** A document with one loaded picture between two words, and one plain paragraph. */
const withLoadedPicture = () =>
  opened(docx(p(t('Alpha ') + pictureRun(5) + t(' Omega')) + p(t('Plain words.')), MATRIX_EXTRAS));

/** PNG bytes of `size` that compress poorly, different for each `seed`. */
function pictureBytes(seed: number, size = 4096): Uint8Array {
  const bytes = new Uint8Array(size);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  let state = seed * 2_654_435_761 + 1;
  for (let k = 8; k < size; k += 1) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    bytes[k] = state & 255;
  }

  return bytes;
}

/** Each picture inserted as a suggestion at the start of the first paragraph, by its own change. */
function addPictures(doc: OrderedDocument, pictures: readonly Uint8Array[]): void {
  pictures.forEach((bytes, k) => {
    const media = {
      sha256: orderedBlobHash(bytes),
      contentType: 'image/png' as const,
      size: bytes.length,
    };
    const object = { type: 'picture', media, extent: { cx: 914_400, cy: 457_200 } } as const;
    apply(doc, `ben-1:${k + 1}`, { ...EDIT, kind: 'insert-object', place: at(doc, 0, 0), object });
  });
}

/** The bytes of `pictures` by SHA-256, as a room's storage holds them. */
const storeOf = (pictures: readonly Uint8Array[]) =>
  new Map(pictures.map((bytes) => [orderedBlobHash(bytes), bytes]));

/** The file `writeOrderedPackage` writes, with the SHA-256 of each `media` call in order. */
async function streamed(
  { doc, base }: Opened,
  store: ReadonlyMap<string, Uint8Array>
): Promise<{ bytes: Uint8Array; calls: string[] }> {
  const chunks: Uint8Array[] = [];
  const calls: string[] = [];
  await writeOrderedPackage(doc, base, {
    media: async (sha256) => {
      calls.push(sha256);
      const bytes = store.get(sha256);
      if (!bytes) {
        throw new Error(`no blob ${sha256}`);
      }

      return bytes;
    },
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return { bytes: joined(chunks), calls };
}

function joined(chunks: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }

  return out;
}

/** The SHA-256 of each added picture, in the order of its entry in the file. */
function addedInFileOrder(bytes: Uint8Array, store: ReadonlyMap<string, Uint8Array>): string[] {
  const out: string[] = [];
  for (const data of Object.values(unzipSync(bytes))) {
    const sha256 = orderedBlobHash(data);
    if (store.has(sha256)) {
      out.push(sha256);
    }
  }

  return out;
}

/** Equal bytes, without printing two whole files when they differ. */
function expectSameBytes(actual: Uint8Array, expected: Uint8Array): void {
  expect(actual.length).toBe(expected.length);
  expect(Buffer.compare(actual, expected)).toBe(0);
}

describe('writeOrderedPackage writes what orderedPackageBytes gives', () => {
  test('a file with no pictures, after an edit, reads no media', async () => {
    const file = opened(docx(p(t('Plain words.')) + p(t('More words.'))));
    apply(file.doc, 'ana-1:1', {
      ...EDIT,
      kind: 'insert',
      place: at(file.doc, 1, 4),
      text: ' new',
    });
    const { bytes, calls } = await streamed(file, new Map());
    expectSameBytes(bytes, orderedPackageBytes(file.doc, file.base));
    expect(calls).toEqual([]);
  });

  test('pictures from the file are written from the file, with no media read', async () => {
    const file = withLoadedPicture();
    const { bytes, calls } = await streamed(file, new Map());
    expectSameBytes(bytes, orderedPackageBytes(file.doc, file.base));
    expect(calls).toEqual([]);
  });

  test('added pictures are read once each, in the order of their entries', async () => {
    const file = withLoadedPicture();
    const pictures = [1, 2, 3].map((seed) => pictureBytes(seed));
    addPictures(file.doc, pictures);
    const store = storeOf(pictures);
    const { bytes, calls } = await streamed(file, store);
    expectSameBytes(
      bytes,
      orderedPackageBytes(file.doc, file.base, { blobs: (h) => store.get(h) })
    );
    expect(calls).toEqual(addedInFileOrder(bytes, store));
    expect(new Set(calls)).toEqual(new Set(store.keys()));
  });

  test('a large picture, deflated across many slices', async () => {
    const file = withLoadedPicture();
    const pictures = [pictureBytes(7, 6 * 1024 * 1024 + 13)];
    addPictures(file.doc, pictures);
    const store = storeOf(pictures);
    const { bytes, calls } = await streamed(file, store);
    expectSameBytes(
      bytes,
      orderedPackageBytes(file.doc, file.base, { blobs: (h) => store.get(h) })
    );
    expect(calls).toHaveLength(1);
  });

  for (const name of FILES) {
    test(`${name}, as loaded and with two added pictures`, async () => {
      const source = new Uint8Array(readFileSync(join(FIXTURES, name)));
      const file = opened(source);
      expectSameBytes(
        (await streamed(file, new Map())).bytes,
        orderedPackageBytes(file.doc, file.base)
      );

      const pictures = [11, 12].map((seed) => pictureBytes(seed));
      addPictures(file.doc, pictures);
      const store = storeOf(pictures);
      const { bytes, calls } = await streamed(file, store);
      expectSameBytes(
        bytes,
        orderedPackageBytes(file.doc, file.base, { blobs: (h) => store.get(h) })
      );
      expect(calls).toEqual(addedInFileOrder(bytes, store));
    });
  }
});

describe('writeOrderedPackage reads and writes one thing at a time', () => {
  test('a slow write is awaited, and no media read or write overlaps another', async () => {
    const file = withLoadedPicture();
    const pictures = [1, 2, 3].map((seed) => pictureBytes(seed, 300_000));
    addPictures(file.doc, pictures);
    const store = storeOf(pictures);
    const chunks: Uint8Array[] = [];
    let busy = 0;
    let most = 0;
    const slow = async <T>(work: () => T): Promise<T> => {
      busy += 1;
      most = Math.max(most, busy);
      await new Promise((done) => setTimeout(done, 2));
      busy -= 1;
      return work();
    };
    await writeOrderedPackage(file.doc, file.base, {
      media: (sha256) => slow(() => store.get(sha256) ?? new Uint8Array(0)),
      write: (chunk) => slow(() => void chunks.push(chunk)),
    });
    expect(most).toBe(1);
    expectSameBytes(
      joined(chunks),
      orderedPackageBytes(file.doc, file.base, { blobs: (h) => store.get(h) })
    );
  });

  test('the writer holds no added picture after its entry is written', async () => {
    const file = withLoadedPicture();
    const seeds = [21, 22, 23, 24];
    addPictures(
      file.doc,
      seeds.map((seed) => pictureBytes(seed, 1024 * 1024))
    );
    const bySha = new Map(
      seeds.map((seed) => [orderedBlobHash(pictureBytes(seed, 1024 * 1024)), seed])
    );
    const given: WeakRef<Uint8Array>[] = [];
    const heldAtRead: number[] = [];
    await writeOrderedPackage(file.doc, file.base, {
      media: async (sha256) => {
        // Collected from a new task, so no stale pointer on the stack keeps a picture
        await new Promise((done) => setTimeout(done, 0));
        Bun.gc(true);
        heldAtRead.push(given.filter((ref) => ref.deref() !== undefined).length);

        // A new copy at each read, so only the writer can keep it
        const bytes = pictureBytes(bySha.get(sha256) ?? 0, 1024 * 1024);
        given.push(new WeakRef(bytes));
        return bytes;
      },
      write: async () => {},
    });
    expect(heldAtRead).toEqual([0, 0, 0, 0]);
  });
});

describe('writeOrderedPackage refuses missing media', () => {
  const twoPictures = () => {
    const file = withLoadedPicture();
    const pictures = [1, 2].map((seed) => pictureBytes(seed));
    addPictures(file.doc, pictures);
    return { file, store: storeOf(pictures) };
  };

  async function refusalOf(write: () => Promise<void>): Promise<string> {
    try {
      await write();
    } catch (error) {
      if (error instanceof OrderedRefusal) {
        return `${error.reason} ${error.message}`;
      }

      throw error;
    }

    return 'written';
  }

  test('a rejected read is missing-blob, and nothing is written after it', async () => {
    const { file, store } = twoPictures();
    const [first, second] = [...store.keys()];
    let written = 0;
    let after = 0;
    let failed = false;
    const reason = await refusalOf(() =>
      writeOrderedPackage(file.doc, file.base, {
        media: async (sha256) => {
          if (sha256 === second) {
            failed = true;
            throw new Error('storage is down');
          }

          return store.get(sha256) ?? new Uint8Array(0);
        },
        write: async () => {
          written += 1;
          if (failed) {
            after += 1;
          }
        },
      })
    );
    expect(reason).toStartWith('missing-blob');
    expect(reason).toContain(second);
    expect(reason).not.toContain(first);
    expect(written).toBeGreaterThan(0);
    expect(after).toBe(0);
  });

  test('bytes of another size are missing-blob', async () => {
    const { file, store } = twoPictures();
    const reason = await refusalOf(() =>
      writeOrderedPackage(file.doc, file.base, {
        media: async (sha256) => (store.get(sha256) ?? new Uint8Array(0)).subarray(1),
        write: async () => {},
      })
    );
    expect(reason).toStartWith('missing-blob');
  });

  test('a rejected write stops the writer with its error', async () => {
    const { file, store } = twoPictures();
    let reads = 0;
    const failure = writeOrderedPackage(file.doc, file.base, {
      media: async (sha256) => {
        reads += 1;
        return store.get(sha256) ?? new Uint8Array(0);
      },
      write: async () => {
        throw new Error('upload stopped');
      },
    });
    await expect(failure).rejects.toThrow('upload stopped');
    expect(reads).toBe(0);
  });
});
