// A picture over the room's size cap, with two real editors and one room: the session refuses it
// by its own reason, `media-too-large`, before the editor commits; nothing reaches the room and
// the next edit goes. Bytes that are not a picture keep `invalid-media`.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  validateRasterHeader,
  type ImageDecodePort,
} from '../../../../store/package/image-resources.ts';
import { ORDERED_RESOURCE_LIMITS } from '../../change/resources.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { RED_PNG } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedSvgPicture } from './paste-harness.ts';

afterEach(closeAll);

const { mediaSize, svgSize } = ORDERED_RESOURCE_LIMITS;

/** Decodes by the raster header, as the browser port does, without a browser. */
const decodePort = (): ImageDecodePort => ({
  async decode(bytes, mime) {
    const header = validateRasterHeader(bytes, mime);
    if (!header) {
      throw new Error('not an image');
    }

    return { pixelWidth: header.pixelWidth, pixelHeight: header.pixelHeight, dpiX: 96, dpiY: 96 };
  },
});

/** The organiser, and a party who adds the pictures. */
async function room() {
  const r = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const imageDecodePort = decodePort();
  const olive = await join(r, 'Olive', { mode: 'edit', login: ORGANISER, imageDecodePort });
  const ben = await join(r, 'Ben', { login: PARTY, imageDecodePort });
  return { r, olive, ben };
}

/** The red PNG with padding after its end, `size` bytes in all. */
function pngOfSize(size: number): Uint8Array {
  const bytes = new Uint8Array(size);
  bytes.set(RED_PNG);
  return bytes;
}

/** An SVG of `size` bytes: a comment fills it. */
function svgOfSize(size: number): Uint8Array {
  const open = '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="4"><!--';
  const close = '--></svg>';
  return new TextEncoder().encode(open + 'x'.repeat(size - open.length - close.length) + close);
}

/** Insert a PNG picture in the second paragraph; whether the editor took it. */
async function insertPicture(person: Person, bytes: Uint8Array): Promise<boolean> {
  caretAt(person, 1, 3);
  const result = await person.editor.executeImageCommand({
    type: 'insertImage',
    data: bytes,
    mime: 'image/png',
    widthPoints: 24,
    heightPoints: 24,
  });
  await settle(16);
  return result.ok;
}

/** Paste a picture whose SVG version is `svg`, with its PNG image as the fallback. */
async function pasteSvgPicture(person: Person, svg: Uint8Array): Promise<void> {
  const { text, html } = await copiedSvgPicture(svg);
  caretAt(person, 1, 3);
  surface(person).pasteRich(text, html);
  await settle(16);
}

const reasons = (person: Person) => person.refusals.map((x) => [x.code, x.reason, x.detail]);

const pictures = (person: Person) => editorXml(person).match(/<w:drawing[ >]/g)?.length ?? 0;

/** Nothing of the refused edit is in the room or on a page; the next edit goes. */
async function leftNothing(
  r: ReturnType<typeof localRoom>,
  olive: Person,
  ben: Person
): Promise<void> {
  expect(r.room.order).toBe(0);
  for (const person of [olive, ben]) {
    expect(pictures(person)).toBe(0);
    expect(sameAsRoom(person, r)).toBe(true);
  }

  caretAt(ben, 2, 0);
  ben.editor.exec({ type: 'insertText', text: 'Later ' });
  await settle(16);
  expect(r.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert']);
  for (const person of [olive, ben]) {
    expect(editorXml(person)).toContain('Later ');
    expect(sameAsRoom(person, r)).toBe(true);
  }
}

describe('a picture over the room size cap', () => {
  test('an inserted picture over 25 MiB is refused as media-too-large', async () => {
    const { r, olive, ben } = await room();
    expect(await insertPicture(ben, pngOfSize(mediaSize + 1))).toBe(false);
    expect(reasons(ben)).toEqual([['ordered-refused', 'media-too-large', 'image/png']]);
    await leftNothing(r, olive, ben);
    expect(reasons(ben)).toHaveLength(1);
  });

  test('a pasted SVG over 2 MiB is refused as media-too-large', async () => {
    const { r, olive, ben } = await room();
    await pasteSvgPicture(ben, svgOfSize(svgSize + 1));
    expect(reasons(ben)).toEqual([['ordered-refused', 'media-too-large', 'image/svg+xml']]);
    expect(r.blob(orderedBlobHash(RED_PNG))).toBeUndefined();
    await leftNothing(r, olive, ben);
    expect(reasons(ben)).toHaveLength(1);
  });

  test('a picture at the cap is not refused for its size', async () => {
    const { r, olive, ben } = await room();
    await pasteSvgPicture(ben, svgOfSize(svgSize));
    expect(await insertPicture(ben, pngOfSize(mediaSize))).toBe(true);
    expect(ben.refusals).toEqual([]);
    expect(r.room.joinSnapshot(0).map((change) => change.kind)).toEqual([
      'insert-fragment',
      'insert-object',
    ]);
    for (const person of [olive, ben]) {
      expect(pictures(person)).toBe(2);
      expect(sameAsRoom(person, r)).toBe(true);
    }
  });

  test('bytes that are not a picture are still refused as invalid-media', async () => {
    const { r, olive, ben } = await room();
    await pasteSvgPicture(ben, new TextEncoder().encode('<svg width="4"></svg>'));
    expect(reasons(ben)).toEqual([['ordered-refused', 'invalid-media', 'image/svg+xml']]);
    await leftNothing(r, olive, ben);
  });
});
