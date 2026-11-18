// Media bytes a copy holds or keeps (group 9.10, design D14): checked by content as the content
// type the change names. A raster image takes the engine's media gate; an SVG must be a bounded,
// well-formed UTF-8 SVG document; TIFF, EMF and WMF must have their headers.

import { describe, expect, test } from 'bun:test';
import { isOrderedMediaBytes } from '../client/media-check.ts';
import { OrderedMedia } from '../client/media.ts';
import { OrderedClientRefusal } from '../client/refusal.ts';
import { emfBytes, RED_PNG, RED_SVG, wmfBytes } from './word-fixtures.ts';

const svg = (text: string) => new TextEncoder().encode(text);
const SVG_ROOT = '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4">';

/** A TIFF of one image whose header gives `width` by `height` pixels. */
function tiffBytes(width: number, height: number): Uint8Array {
  const bytes = new Uint8Array(8 + 2 + 2 * 12 + 4);
  const view = new DataView(bytes.buffer);
  bytes.set([0x49, 0x49, 0x2a, 0x00]);
  view.setUint32(4, 8, true);
  view.setUint16(8, 2, true);
  for (const [k, tag, value] of [
    [0, 256, width],
    [1, 257, height],
  ] as const) {
    const at = 10 + k * 12;
    view.setUint16(at, tag, true);
    view.setUint16(at + 2, 4, true);
    view.setUint32(at + 4, 1, true);
    view.setUint32(at + 8, value, true);
  }
  return bytes;
}

describe('media bytes by content', () => {
  test('an image of the type named, and nothing else', () => {
    expect(isOrderedMediaBytes(RED_PNG, 'image/png')).toBe(true);
    expect(isOrderedMediaBytes(RED_PNG, 'image/svg+xml')).toBe(false);
    expect(isOrderedMediaBytes(RED_SVG, 'image/png')).toBe(false);
    expect(isOrderedMediaBytes(new Uint8Array(), 'image/png')).toBe(false);
  });

  test('an SVG: a well-formed UTF-8 document whose root is svg in the SVG namespace', () => {
    expect(isOrderedMediaBytes(RED_SVG, 'image/svg+xml')).toBe(true);
    expect(
      isOrderedMediaBytes(svg(`<?xml version="1.0"?>\n${SVG_ROOT}</svg>`), 'image/svg+xml')
    ).toBe(true);
    // Painted as an image, a script never runs: the document is still an SVG
    expect(
      isOrderedMediaBytes(svg(`${SVG_ROOT}<script>alert(1)</script></svg>`), 'image/svg+xml')
    ).toBe(true);
    expect(isOrderedMediaBytes(svg('<svg width="4"></svg>'), 'image/svg+xml')).toBe(false);
    expect(isOrderedMediaBytes(svg(`${SVG_ROOT}<g></svg>`), 'image/svg+xml')).toBe(false);
    expect(
      isOrderedMediaBytes(svg(`<!DOCTYPE svg [<!ENTITY x "y">]>${SVG_ROOT}</svg>`), 'image/svg+xml')
    ).toBe(false);
    expect(
      isOrderedMediaBytes(
        svg(`${SVG_ROOT}<!ENTITY x SYSTEM "file:///etc/passwd"></svg>`),
        'image/svg+xml'
      )
    ).toBe(false);
    const latin1 = Uint8Array.from([...svg(`${SVG_ROOT}<text>`), 0xff, ...svg('</text></svg>')]);
    expect(isOrderedMediaBytes(latin1, 'image/svg+xml')).toBe(false);
  });

  test('a TIFF within the pixel caps, an EMF and a WMF by their headers', () => {
    expect(isOrderedMediaBytes(tiffBytes(4, 4), 'image/tiff')).toBe(true);
    expect(isOrderedMediaBytes(tiffBytes(100_000, 100_000), 'image/tiff')).toBe(false);
    expect(isOrderedMediaBytes(emfBytes(), 'image/x-emf')).toBe(true);
    const unsigned = emfBytes();
    unsigned[40] = 0;
    expect(isOrderedMediaBytes(unsigned, 'image/x-emf')).toBe(false);
    expect(isOrderedMediaBytes(emfBytes().slice(0, 80), 'image/x-emf')).toBe(false);
    expect(isOrderedMediaBytes(wmfBytes(), 'image/x-wmf')).toBe(true);
    expect(isOrderedMediaBytes(wmfBytes(false), 'image/x-wmf')).toBe(true);
    const version = wmfBytes();
    version[26] = 2;
    expect(isOrderedMediaBytes(version, 'image/x-wmf')).toBe(false);
  });

  test("an embedded object's part and an ink part by their signatures, never opened", () => {
    const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0, 0, 0, 0]);
    const compound = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0]);
    const ink = svg('\uFEFF <inkml:ink xmlns:inkml="http://www.w3.org/2003/InkML"/>');
    const cases: [Uint8Array, string, boolean][] = [
      [compound, 'application/vnd.openxmlformats-officedocument.oleObject', true],
      [zip, 'application/vnd.openxmlformats-officedocument.oleObject', false],
      [zip, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', true],
      [zip, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', true],
      [compound, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', false],
      [ink, 'application/inkml+xml', true],
      [svg('ink'), 'application/inkml+xml', false],
      [RED_PNG, 'application/inkml+xml', false],
    ];
    for (const [bytes, type, ok] of cases)
      expect({ type, ok: isOrderedMediaBytes(bytes, type as never) }).toEqual({ type, ok });
  });

  test('the media store holds an SVG version and refuses bytes that are not its type', () => {
    const media = new OrderedMedia();
    expect(media.hold(RED_SVG, 'image/svg+xml')).toMatchObject({
      contentType: 'image/svg+xml',
      size: RED_SVG.byteLength,
    });
    expect(() => media.hold(RED_SVG, 'image/png')).toThrow(OrderedClientRefusal);
    expect(() => media.hold(RED_SVG, 'image/svg')).toThrow(OrderedClientRefusal);
  });
});
