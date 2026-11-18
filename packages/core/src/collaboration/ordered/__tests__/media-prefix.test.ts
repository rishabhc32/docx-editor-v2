// The check of media from its first bytes, held to the whole check: real files of every type, cut
// at every byte, with three different rests past each cut.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { zipSync } from 'fflate';
import { validateJpegHeader, validateTiffHeader } from '../../../store/package/image-resources.ts';
import { resolveImageResourceLimits } from '../../../store/runtime/limits.ts';
import { ORDERED_RESOURCE_LIMITS } from '../change/resources.ts';
import {
  checkOrderedMediaPrefix,
  isOrderedMediaBytes,
  ORDERED_MEDIA_PREFIX,
} from '../client/media-check.ts';
import { OrderedMedia } from '../client/media.ts';
import { OrderedClientRefusal } from '../client/refusal.ts';
import type { OrderedMediaType } from '../model/resources.ts';
import { seededRandom } from './random.ts';
import { RED_PNG } from './word-fixtures.ts';

type Answer = ReturnType<typeof checkOrderedMediaPrefix>;

/** The first bytes the check reads a signature or a fixed header from (`SIGNATURE_BYTES`). */
const SIGNATURE_BYTES = 8_192;

const OLE = 'application/vnd.openxmlformats-officedocument.oleObject';
const DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const INK = 'application/inkml+xml';

const fixture = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(new URL(`./fixtures/media/${name}`, import.meta.url)));

const text = (value: string): Uint8Array => new TextEncoder().encode(value);

function noise(size: number, seed: number): Uint8Array {
  const random = seededRandom(seed);
  return Uint8Array.from({ length: size }, () => Math.floor(random() * 256));
}

/** Bytes that stand in for the rest of a file past a cut: zeros, 0xff bytes and noise. */
const rests = (size: number): Uint8Array[] => [
  new Uint8Array(size),
  new Uint8Array(size).fill(0xff),
  noise(size, size),
];

/**
 * The prefix check at every cut of `bytes`. Each answer other than "whole" must be the whole
 * check's answer for the file of that size, whatever bytes follow the cut. Returns the answers
 * by cut, and every cut where the whole check disagrees.
 */
function everyCut(
  bytes: Uint8Array,
  type: OrderedMediaType
): { answers: Answer[]; wrong: string[] } {
  const size = bytes.byteLength;
  const wholes = rests(size);
  const answers: Answer[] = [];
  const wrong: string[] = [];
  for (let cut = 0; cut <= size; cut += 1) {
    if (cut > 0) {
      for (const whole of wholes) {
        whole[cut - 1] = bytes[cut - 1]!;
      }
    }

    const answer = checkOrderedMediaPrefix(bytes.subarray(0, cut), { contentType: type, size });
    answers.push(answer);
    if (answer === 'whole') {
      continue;
    }

    for (const [k, whole] of wholes.entries()) {
      if (isOrderedMediaBytes(whole, type) !== (answer === 'valid')) {
        wrong.push(`cut ${cut}, rest ${k}: ${answer}`);
      }
    }
  }

  return { answers, wrong };
}

/** The first cut that gives `answer`, after which every cut gives it, and only "whole" before. */
function decidedFrom(answers: readonly Answer[], answer: Answer): number {
  const first = answers.indexOf(answer);
  expect(first).toBeGreaterThan(0);
  expect(answers.slice(0, first).every((each) => each === 'whole')).toBe(true);
  expect(answers.slice(first).every((each) => each === answer)).toBe(true);
  return first;
}

/** Where each JPEG segment starts (its 0xff) and ends, up to the scan. */
function jpegSegments(bytes: Uint8Array): { marker: number; at: number; end: number }[] {
  const out: { marker: number; at: number; end: number }[] = [];
  let at = 2;
  while (bytes[at + 1] !== 0xda) {
    const end = at + 2 + ((bytes[at + 2]! << 8) | bytes[at + 3]!);
    out.push({ marker: bytes[at + 1]!, at, end });
    at = end;
  }

  return out;
}

/** A TIFF's first image directory: where it starts, its entry count and the byte order. */
function tiffDirectory(bytes: Uint8Array): { view: DataView; at: number; little: boolean } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const little = bytes[0] === 0x49;
  return { view, at: view.getUint32(4, little), little };
}

/** A TIFF whose first image directory is copied past `gap` more bytes, at the end. */
function tiffWithDirectoryAfter(bytes: Uint8Array, gap: number): Uint8Array {
  const { view, at, little } = tiffDirectory(bytes);
  const length = 2 + view.getUint16(at, little) * 12 + 4;
  const out = new Uint8Array(bytes.byteLength + gap + length);
  out.set(bytes);
  out.set(bytes.subarray(at, at + length), bytes.byteLength + gap);

  // The header points at the copy; the first directory stays behind as unread bytes
  new DataView(out.buffer).setUint32(4, bytes.byteLength + gap, little);
  return out;
}

/** A JPEG with `count` empty application segments of the most bytes one holds after its start. */
function jpegWithSegmentsFirst(bytes: Uint8Array, count: number): Uint8Array {
  const segment = new Uint8Array(2 + 65_535);
  segment.set([0xff, 0xef, 0xff, 0xff]);
  const out = new Uint8Array(bytes.byteLength + count * segment.byteLength);
  out.set(bytes.subarray(0, 2));
  for (let k = 0; k < count; k += 1) {
    out.set(segment, 2 + k * segment.byteLength);
  }

  out.set(bytes.subarray(2), 2 + count * segment.byteLength);
  return out;
}

/** An SVG document of exactly `size` bytes. */
function svgOfSize(size: number): Uint8Array {
  const open = '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><!--';
  const close = '--></svg>';
  return text(open + 'x'.repeat(size - open.length - close.length) + close);
}

const prefixOf = (bytes: Uint8Array): Uint8Array => bytes.subarray(0, ORDERED_MEDIA_PREFIX);

describe('the check of media from its first bytes, at every cut of real files', () => {
  const fixedHeaders: [string, OrderedMediaType][] = [
    ['noise.png', 'image/png'],
    ['noise.gif', 'image/gif'],
    ['noise.bmp', 'image/bmp'],
    ['lossy.webp', 'image/webp'],
    ['lossless.webp', 'image/webp'],
    ['word.emf', 'image/x-emf'],
    ['polylines.wmf', 'image/x-wmf'],
  ];

  for (const [name, type] of fixedHeaders) {
    test(`${name}: decided from the first 8 KiB, and always as the whole check`, () => {
      const bytes = fixture(name);
      expect(isOrderedMediaBytes(bytes, type)).toBe(true);
      const { answers, wrong } = everyCut(bytes, type);
      expect(wrong).toEqual([]);
      expect(decidedFrom(answers, 'valid')).toBe(SIGNATURE_BYTES);
    });
  }

  test('a JPEG whose frame header lies past 8 KiB: "whole" until the header is in', () => {
    const bytes = fixture('late-frame.jpeg');
    expect(isOrderedMediaBytes(bytes, 'image/jpeg')).toBe(true);
    const { answers, wrong } = everyCut(bytes, 'image/jpeg');
    expect(wrong).toEqual([]);
    const first = decidedFrom(answers, 'valid');
    expect(first).toBeGreaterThan(SIGNATURE_BYTES);
    expect(validateJpegHeader(bytes.subarray(0, first - 1))).toBeNull();
    expect(validateJpegHeader(bytes.subarray(0, first))).not.toBeNull();
  });

  for (const name of ['late-directory.tiff', 'big-endian.tiff']) {
    test(`${name}: its first image directory at the end, "whole" until it is in`, () => {
      const bytes = fixture(name);
      expect(isOrderedMediaBytes(bytes, 'image/tiff')).toBe(true);
      const { answers, wrong } = everyCut(bytes, 'image/tiff');
      expect(wrong).toEqual([]);
      const first = decidedFrom(answers, 'valid');
      expect(first).toBeGreaterThan(SIGNATURE_BYTES);
      expect(validateTiffHeader(bytes.subarray(0, first - 1))).toBeNull();
      expect(validateTiffHeader(bytes.subarray(0, first))).not.toBeNull();
    });
  }

  test('an EMF header record that ends past the first bytes is measured against the file', () => {
    const bytes = fixture('word.emf');
    new DataView(bytes.buffer).setUint32(4, bytes.byteLength - 4, true);
    expect(isOrderedMediaBytes(bytes, 'image/x-emf')).toBe(true);
    const { answers, wrong } = everyCut(bytes, 'image/x-emf');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'valid')).toBe(SIGNATURE_BYTES);
  });

  test('an SVG is "whole" at every cut but the last', () => {
    const bytes = fixture('paths.svg');
    const { answers, wrong } = everyCut(bytes, 'image/svg+xml');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'valid')).toBe(bytes.byteLength);
  });

  test("an object's part and an ink part are decided by their signatures", () => {
    const zip = zipSync({ 'word/media/noise.bin': noise(12_000, 1) }, { level: 0 });
    const compound = noise(12_000, 2);
    compound.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    const ink = text(
      `﻿${' '.repeat(4_000)}<inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML">` +
        '<inkml:trace>10 0, 9 14, 8 28</inkml:trace>'.repeat(200) +
        '</inkml:ink>'
    );
    const cases: [Uint8Array, OrderedMediaType][] = [
      [zip, DOCX],
      [compound, OLE],
      [ink, INK],
    ];
    for (const [bytes, type] of cases) {
      const { answers, wrong } = everyCut(bytes, type);
      expect(wrong).toEqual([]);
      expect(decidedFrom(answers, 'valid')).toBe(SIGNATURE_BYTES);
    }

    const asObject = everyCut(zip, OLE);
    expect(asObject.wrong).toEqual([]);
    expect(decidedFrom(asObject.answers, 'invalid')).toBe(SIGNATURE_BYTES);
  });
});

describe('"invalid" only when the first bytes decide it', () => {
  test('a wrong signature', () => {
    const { answers, wrong } = everyCut(fixture('noise.png'), 'image/gif');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'invalid')).toBe(SIGNATURE_BYTES);
  });

  test('a PNG header over the pixel caps', () => {
    const bytes = fixture('noise.png');
    const view = new DataView(bytes.buffer);
    view.setUint32(16, 40_000);
    view.setUint32(20, 40_000);
    const { answers, wrong } = everyCut(bytes, 'image/png');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'invalid')).toBe(SIGNATURE_BYTES);
  });

  test('a JPEG frame header over the pixel caps, once the header is in', () => {
    const bytes = fixture('late-frame.jpeg');
    const frame = jpegSegments(bytes).find((segment) => segment.marker === 0xc0)!;
    const view = new DataView(bytes.buffer);
    view.setUint16(frame.at + 5, 20_000);
    view.setUint16(frame.at + 7, 20_000);
    const { answers, wrong } = everyCut(bytes, 'image/jpeg');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'invalid')).toBe(frame.end);
  });

  test('a TIFF directory over the pixel caps, once the directory is in', () => {
    const bytes = fixture('late-directory.tiff');
    const { view, at, little } = tiffDirectory(bytes);
    const count = view.getUint16(at, little);
    for (let k = 0; k < count; k += 1) {
      const entry = at + 2 + k * 12;
      if (view.getUint16(entry, little) === 256) {
        view.setUint16(entry + 2, 4, little);
        view.setUint32(entry + 8, 100_000, little);
      }
    }

    const { answers, wrong } = everyCut(bytes, 'image/tiff');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'invalid')).toBe(at + 2 + count * 12);
  });

  test('an EMF header record longer than the file', () => {
    const bytes = fixture('word.emf');
    new DataView(bytes.buffer).setUint32(4, bytes.byteLength + 4, true);
    const { answers, wrong } = everyCut(bytes, 'image/x-emf');
    expect(wrong).toEqual([]);
    expect(decidedFrom(answers, 'invalid')).toBe(SIGNATURE_BYTES);
  });

  test('a size outside the caps, or first bytes longer than the size', () => {
    const png = fixture('noise.png');
    const check = (size: number, type: OrderedMediaType = 'image/png') =>
      checkOrderedMediaPrefix(png.subarray(0, SIGNATURE_BYTES), { contentType: type, size });
    expect(check(0)).toBe('invalid');
    expect(check(-1)).toBe('invalid');
    expect(check(SIGNATURE_BYTES + 0.5)).toBe('invalid');
    expect(check(SIGNATURE_BYTES - 1)).toBe('invalid');
    expect(check(ORDERED_RESOURCE_LIMITS.mediaSize + 1)).toBe('invalid');
    expect(check(ORDERED_RESOURCE_LIMITS.mediaSize)).toBe('valid');
    expect(check(ORDERED_RESOURCE_LIMITS.svgSize + 1, 'image/svg+xml')).toBe('invalid');
  });

  test("the raster check's own byte cap never decides inside the media cap", () => {
    expect(resolveImageResourceLimits().maxEncodedBytes).toBeGreaterThanOrEqual(
      ORDERED_RESOURCE_LIMITS.mediaSize
    );
  });
});

describe('media past the first MiB', () => {
  test('a JPEG whose frame header lies past the first MiB needs the whole file', () => {
    const bytes = jpegWithSegmentsFirst(fixture('late-frame.jpeg'), 17);
    expect(bytes.byteLength).toBeGreaterThan(ORDERED_MEDIA_PREFIX);
    const media = { contentType: 'image/jpeg', size: bytes.byteLength } as const;
    expect(checkOrderedMediaPrefix(prefixOf(bytes), media)).toBe('whole');
    expect(checkOrderedMediaPrefix(bytes, media)).toBe('valid');
    expect(isOrderedMediaBytes(bytes, 'image/jpeg')).toBe(true);
  });

  test('a TIFF whose first image directory lies past the first MiB needs the whole file', () => {
    const bytes = tiffWithDirectoryAfter(fixture('late-directory.tiff'), ORDERED_MEDIA_PREFIX);
    const media = { contentType: 'image/tiff', size: bytes.byteLength } as const;
    expect(checkOrderedMediaPrefix(prefixOf(bytes), media)).toBe('whole');
    expect(checkOrderedMediaPrefix(bytes, media)).toBe('valid');
    expect(isOrderedMediaBytes(bytes, 'image/tiff')).toBe(true);
  });

  test('when the first bytes are the whole file, they get the whole check', () => {
    expect(
      checkOrderedMediaPrefix(RED_PNG, { contentType: 'image/png', size: RED_PNG.byteLength })
    ).toBe('valid');
    const cut = RED_PNG.subarray(0, 20);
    expect(checkOrderedMediaPrefix(cut, { contentType: 'image/png', size: 20 })).toBe('invalid');
  });
});

describe('the SVG cap (2 MiB)', () => {
  const { svgSize } = ORDERED_RESOURCE_LIMITS;

  test('an SVG at the cap passes, one byte over it does not', () => {
    const atCap = svgOfSize(svgSize);
    const over = svgOfSize(svgSize + 1);
    expect(isOrderedMediaBytes(atCap, 'image/svg+xml')).toBe(true);
    expect(isOrderedMediaBytes(over, 'image/svg+xml')).toBe(false);
    expect(
      checkOrderedMediaPrefix(prefixOf(atCap), { contentType: 'image/svg+xml', size: svgSize })
    ).toBe('whole');
    expect(
      checkOrderedMediaPrefix(prefixOf(over), { contentType: 'image/svg+xml', size: svgSize + 1 })
    ).toBe('invalid');
  });

  test('a copy holds an SVG at the cap and refuses one over it as too large', () => {
    const media = new OrderedMedia();
    expect(media.hold(svgOfSize(svgSize), 'image/svg+xml')).toMatchObject({ size: svgSize });
    expect(() => media.hold(svgOfSize(svgSize + 1), 'image/svg+xml')).toThrow(
      new OrderedClientRefusal('ordered-refused', 'media-too-large', 'image/svg+xml')
    );
  });
});
