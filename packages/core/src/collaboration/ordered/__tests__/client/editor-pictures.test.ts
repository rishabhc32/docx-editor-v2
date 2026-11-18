// Pictures with real editors and one room (group 9.4, task 9.11): a party's picture is a tracked
// insertion whose bytes travel once, by hash, through the room; the organiser's property edits
// are direct, undone through `revert`; every copy ends equal to the room and shows the picture.

import { afterEach, describe, expect, test } from 'bun:test';
import { orderedPackageBytes } from '../../compose/package.ts';
import {
  validateRasterHeader,
  type ImageDecodePort,
} from '../../../../store/package/image-resources.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { encodeOrderedClientMessage } from '../../protocol/codec.ts';
import type { OrderedClientMessage } from '../../protocol/messages.ts';
import type { OrderedChange } from '../../change/types.ts';
import { orderedViewParagraph } from '../../compose/compose.ts';
import type { AutomationHandle, AutomationHost } from '../../../../automation/index.ts';
import { createBrowserAutomationHost } from '../../../../editor/automation-host.ts';
import { createOrderedCollaboration } from '../../client/create.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, RED_PNG, t } from '../word-fixtures.ts';
import {
  closeAll,
  CODE,
  HeldTransport,
  HeldUpload,
  join,
  listXml,
  localRoom,
  ORGANISER,
  PARTY,
  ROOM,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

/** A picture between two words, and a plain paragraph. */
const PICTURE_FILE = docx(
  p(t('Alpha ') + pictureRun(5) + t(' Omega')) + p(t('Plain words.')),
  MATRIX_EXTRAS
);

afterEach(closeAll);

/** Decodes by the raster header, as the browser port does, without a browser. */
const decodePort = (): ImageDecodePort => ({
  async decode(bytes, mime) {
    const header = validateRasterHeader(bytes, mime);
    if (!header) throw new Error('not an image');
    return { pixelWidth: header.pixelWidth, pixelHeight: header.pixelHeight, dpiX: 96, dpiY: 96 };
  },
});

async function room() {
  const r = localRoom(PICTURE_FILE, { policy: suggestionsOnlyPolicy() });
  const decode = decodePort();
  const ana = await join(r, 'Ana', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: PICTURE_FILE,
    imageDecodePort: decode,
  });
  const ben = await join(r, 'Ben', {
    login: PARTY,
    wordFile: PICTURE_FILE,
    imageDecodePort: decode,
  });
  return { r, ana, ben };
}

/** Put the caret in a paragraph of a person's editor, by the paragraph's index. */
function caret(person: Person, index: number, offset: number): void {
  const paragraphId = surface(person).session.paragraphIds()[index];
  if (paragraphId === undefined) throw new Error(`no paragraph ${index}`);
  surface(person).setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId, offset },
  });
}

/** Select the first paragraph's picture, as a click on it does. */
async function selectPicture(person: Person, offset = 7): Promise<void> {
  caret(person, 0, offset);
  await settle(4);
  if (!person.editor.snapshot().image) caret(person, 0, offset - 1);
}

/** The media parts of the person's editor package, by name. */
function editorMedia(person: Person): Map<string, Uint8Array> {
  const out = new Map<string, Uint8Array>();
  for (const [name, bytes] of surface(person).session.currentPackage().partBytes)
    if (name.startsWith('/word/media/')) out.set(name, bytes);
  return out;
}

const RED_HASH = orderedBlobHash(RED_PNG);

describe('pictures in an ordered room', () => {
  test("a party's picture is a tracked insertion; its bytes travel once and every copy shows it", async () => {
    const { r, ana, ben } = await room();
    caret(ben, 1, 6);
    const result = await ben.editor.executeImageCommand({
      type: 'insertImage',
      data: RED_PNG,
      mime: 'image/png',
      widthPoints: 24,
      heightPoints: 24,
    });
    expect(result.ok).toBe(true);
    await settle(16);
    expect(ben.refusals).toEqual([]);
    expect(r.blob(RED_HASH)).toEqual(RED_PNG);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toMatch(/<w:ins [^>]*w:author="Ben"[^>]*><w:r><w:drawing[ >]/);
      // The other copy read the bytes back from the room, by hash, and shows them
      expect(
        [...editorMedia(person).values()].some((bytes) => bytes.length === RED_PNG.length)
      ).toBe(true);
    }
    // Saved with the room's blobs: the media part is there while the picture is pending
    const saved = orderedPackageBytes(r.room.document, readPackage(PICTURE_FILE), {
      blobs: (hash) => r.blob(hash),
    });
    const parts = readPackage(saved).partBytes;
    expect(
      [...parts.keys()].filter((name) => name.startsWith('/word/media/ordered_'))
    ).toHaveLength(1);
  });

  test("the organiser's direct picture edits apply on every copy, and undo and redo take them back exactly", async () => {
    const { r, ana, ben } = await room();
    const before = listXml(ana);
    await selectPicture(ana);
    expect(ana.editor.snapshot().image).not.toBeNull();
    expect(
      ana.editor.exec({ type: 'setImageProperties', widthEmu: 228_600, heightEmu: 228_600 })
    ).toEqual({ ok: true, changed: true });
    await settle(16);
    expect(ana.refusals).toEqual([]);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toContain('<wp:extent cx="228600" cy="228600"/>');
    }
    // One command, one undo step
    ana.editor.exec({ type: 'undo' });
    await settle(16);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toBe(before);
    }
    ana.editor.exec({ type: 'redo' });
    await settle(16);
    expect(listXml(ben)).toContain('<wp:extent cx="228600" cy="228600"/>');
    expect(sameAsRoom(ben, r)).toBe(true);
  });

  test('a party may not resize: the editor refuses in suggesting mode, and the room refuses an untracked one', async () => {
    const { r, ben } = await room();
    await selectPicture(ben);
    const refused = ben.editor.exec({
      type: 'setImageProperties',
      widthEmu: 228_600,
      heightEmu: 228_600,
    });
    expect(refused.ok).toBe(false);
    // The same party with tracking off: the host's earlier check refuses it as a package change
    const carl = await join(r, 'Carl', { mode: 'edit', login: PARTY, wordFile: PICTURE_FILE });
    await selectPicture(carl);
    carl.editor.exec({ type: 'setImageProperties', widthEmu: 228_600, heightEmu: 228_600 });
    await settle(16);
    expect(carl.refusals.map((x) => x.reason)).toContain('package');
    expect(sameAsRoom(carl, r)).toBe(true);
    expect(listXml(carl)).toContain('<wp:extent cx="457200" cy="457200"/>');
  });
});

function readPackage(bytes: Uint8Array) {
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

// ---- more scenarios ----

/** A one-pixel GIF: other bytes than the red PNG, for a new image. */
const GIF = Uint8Array.from(atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'), (c) =>
  c.charCodeAt(0)
);

async function insertPicture(
  person: Person,
  index: number,
  offset: number,
  data: Uint8Array = RED_PNG,
  mime: 'image/png' | 'image/gif' = 'image/png'
) {
  caret(person, index, offset);
  const result = await person.editor.executeImageCommand({
    type: 'insertImage',
    data,
    mime,
    widthPoints: 24,
    heightPoints: 24,
  });
  expect(result.ok).toBe(true);
  await settle(16);
}

/** The pending suggestions of an author, in the room's copy. */
function pendingOf(r: ReturnType<typeof localRoom>, author: string): string[] {
  return [...r.room.document.suggestions.values()]
    .filter((s) => s.status === 'pending' && s.attribution.author === author)
    .map((s) => s.id);
}

function saved(r: ReturnType<typeof localRoom>) {
  return readPackage(
    orderedPackageBytes(r.room.document, readPackage(PICTURE_FILE), {
      blobs: (hash) => r.blob(hash),
    })
  );
}

const mediaNames = (pkg: ReturnType<typeof readPackage>) =>
  [...pkg.partBytes.keys()].filter((name) => name.startsWith('/word/media/')).sort();

describe('pictures in an ordered room: every edit, decisions, faults', () => {
  test('the organiser wraps, turns, describes, replaces and deletes a picture; each undo is exact', async () => {
    const { r, ana, ben } = await room();
    const states = [listXml(ana)];
    const commands = [
      { type: 'setImageWrapType', target: 'square' },
      { type: 'transformImage', action: 'rotateCW' },
      { type: 'setImageProperties', alt: 'A red dot' },
    ] as const;
    for (const command of commands) {
      await selectPicture(ana);
      expect(ana.editor.exec(command)).toEqual({ ok: true, changed: true });
      await settle(16);
      expect(ana.refusals).toEqual([]);
      expect(sameAsRoom(ben, r)).toBe(true);
      states.push(listXml(ana));
    }
    expect(states.at(-1)).toContain('descr="A red dot"');
    expect(states.at(-1)).toContain('<wp:wrapSquare');
    // A new image: its bytes go to the room once, and the other copy shows them
    await selectPicture(ana);
    const replaced = await ana.editor.executeImageCommand({
      type: 'replaceImage',
      data: GIF,
      mime: 'image/gif',
    });
    expect(replaced.ok).toBe(true);
    await settle(16);
    expect(r.blob(orderedBlobHash(GIF))).toEqual(GIF);
    expect([...editorMedia(ben).values()].some((bytes) => bytes.length === GIF.length)).toBe(true);
    states.push(listXml(ana));
    await selectPicture(ana);
    expect(ana.editor.exec({ type: 'deleteImage' }).ok).toBe(true);
    await settle(16);
    expect(listXml(ben)).not.toContain('<w:drawing');
    expect(sameAsRoom(ben, r)).toBe(true);
    // Undo walks back through every state, on every copy
    for (let k = states.length - 1; k >= 0; k -= 1) {
      ana.editor.exec({ type: 'undo' });
      await settle(16);
      expect([k, listXml(ben)]).toEqual([k, states[k]]);
      expect(sameAsRoom(ben, r)).toBe(true);
    }
    expect(ana.refusals).toEqual([]);
    // Back to the file: the new image's part is not in the saved file
    expect(mediaNames(saved(r))).toEqual(['/word/media/red.png']);
  });

  test("a party's picture rejected leaves no part; accepted, it reopens as the file's own", async () => {
    const { r, ana, ben } = await room();
    await insertPicture(ben, 1, 0);
    await insertPicture(ben, 1, 13, GIF, 'image/gif');
    // Two places, two suggestions: each is decided on its own
    const pending = pendingOf(r, 'Ben');
    expect(pending).toHaveLength(2);
    expect(mediaNames(saved(r))).toHaveLength(3);
    for (const id of pending) ana.handle.session.requestDecision(id, 'rejected');
    await settle(16);
    expect(sameAsRoom(ben, r)).toBe(true);
    expect(mediaNames(saved(r))).toEqual(['/word/media/red.png']);
    await insertPicture(ben, 1, 0);
    const [again] = pendingOf(r, 'Ben');
    ana.handle.session.requestDecision(again ?? '', 'accepted');
    await settle(16);
    const pkg = saved(r);
    expect(mediaNames(pkg)).toHaveLength(2);
    const body = pkg.parts.get('/word/document.xml');
    expect(body).toBeDefined();
    expect(listXml(ana)).not.toContain('<w:ins');
  });

  test("a party's deletion of a picture, accepted, takes its media out of the saved file", async () => {
    const { r, ana, ben } = await room();
    await selectPicture(ben);
    expect(ben.editor.exec({ type: 'deleteImage' }).ok).toBe(true);
    await settle(16);
    expect(listXml(ana)).toMatch(/<w:del [^>]*w:author="Ben"[^>]*><w:r><w:drawing/);
    expect(mediaNames(saved(r))).toEqual(['/word/media/red.png']);
    ana.handle.session.requestDecision(pendingOf(r, 'Ben')[0] ?? '', 'accepted');
    await settle(16);
    expect(sameAsRoom(ben, r)).toBe(true);
    expect(mediaNames(saved(r))).toEqual([]);
  });

  test('one person deletes a picture while another resizes it: both hold, on every copy', async () => {
    const r = localRoom(PICTURE_FILE, { policy: suggestionsOnlyPolicy() });
    const held = new HeldTransport(r.transport({ person: testPerson('Ben'), ...PARTY }));
    const ana = await join(r, 'Ana', { mode: 'edit', login: ORGANISER, wordFile: PICTURE_FILE });
    const ben = await join(r, 'Ben', { login: PARTY, wordFile: PICTURE_FILE, transport: held });
    held.held = true;
    await selectPicture(ana);
    ana.editor.exec({ type: 'setImageProperties', widthEmu: 228_600, heightEmu: 228_600 });
    await selectPicture(ben);
    ben.editor.exec({ type: 'deleteImage' });
    await settle(16);
    held.held = false;
    held.release();
    await settle(24);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toMatch(/<w:del [^>]*w:author="Ben"[^>]*><w:r><w:drawing/);
      expect(listXml(person)).toContain('<wp:extent cx="228600" cy="228600"/>');
    }
    expect([...ana.refusals, ...ben.refusals]).toEqual([]);
  });

  test('two people insert the same bytes: the room stores them once, and both pictures show', async () => {
    const r = localRoom(PICTURE_FILE, { policy: suggestionsOnlyPolicy() });
    const decode = decodePort();
    const ben = await join(r, 'Ben', {
      login: PARTY,
      wordFile: PICTURE_FILE,
      imageDecodePort: decode,
    });
    const cara = await join(r, 'Cara', {
      login: PARTY,
      wordFile: PICTURE_FILE,
      imageDecodePort: decode,
    });
    await Promise.all([insertPicture(ben, 1, 0), insertPicture(cara, 1, 6)]);
    await settle(24);
    const media = [...r.room.document.resources.added.values()].filter((x) => x.kind === 'media');
    expect(media.map((x) => (x.kind === 'media' ? x.sha256 : ''))).toEqual([RED_HASH, RED_HASH]);
    for (const person of [ben, cara]) expect(sameAsRoom(person, r)).toBe(true);
    expect(mediaNames(saved(r))).toHaveLength(3);
  });

  test('a connection lost during an upload: the upload goes on, and the picture applies once', async () => {
    const r = localRoom(PICTURE_FILE, { policy: suggestionsOnlyPolicy() });
    const login = { person: testPerson('Ben'), ...PARTY };
    const held = new HeldUpload(r.uploader(login));
    const ana = await join(r, 'Ana', { mode: 'edit', login: ORGANISER, wordFile: PICTURE_FILE });
    const ben = await join(r, 'Ben', {
      login: PARTY,
      wordFile: PICTURE_FILE,
      upload: held.upload,
      imageDecodePort: decodePort(),
    });
    await insertPicture(ben, 1, 0);
    expect(held.waiting).toHaveLength(1);

    // The connection drops while the upload is on its way; the upload ends while it is down
    r.disconnect(ben.handle.session.namespace);
    await settle(4);
    expect(ben.handle.session.status()).toBe('disconnected');
    held.goAll();
    await settle(4);
    expect(r.blob(RED_HASH)).toEqual(RED_PNG);
    expect(r.room.order).toBe(0);
    ben.handle.session.reconnect();
    await settle(24);
    expect(held.waiting).toHaveLength(0);
    expect(ben.refusals).toEqual([]);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person).match(/<w:drawing/g)).toHaveLength(2);
    }
  });

  test('the room refuses a change that names media its storage does not hold (missing-blob)', async () => {
    const r = localRoom(PICTURE_FILE, { policy: suggestionsOnlyPolicy() });
    const frames: string[] = [];
    const connection = r.transport({ person: testPerson('Dan'), ...PARTY }).connect({
      open: () => {},
      close: () => {},
      message: (frame) => frames.push(frame),
    });
    const send = (message: OrderedClientMessage) =>
      connection.send(encodeOrderedClientMessage(message));
    send({ v: 3, type: 'hello', room: ROOM, file: 1, last: 0, code: CODE, namespace: 'dan-1' });
    await settle(4);
    const view = orderedViewParagraph(r.room.document, '/word/document.xml', 1);
    const place = view?.chars[0]?.address;
    if (!place) throw new Error('no place');
    const change = {
      v: 1,
      id: 'dan-1:1',
      kind: 'insert-object',
      part: '/word/document.xml',
      mode: 'suggest',
      place,
      object: {
        type: 'picture',
        media: { sha256: RED_HASH, contentType: 'image/png', size: RED_PNG.length },
        extent: { cx: 304_800, cy: 304_800 },
      },
    } as unknown as OrderedChange;
    send({ v: 3, type: 'changes', changes: [change] });
    await settle(4);
    expect(frames.some((f) => f.includes('"missing-blob"'))).toBe(true);
    expect(r.room.order).toBe(0);
    // Once the bytes are stored, the same change applies
    const stored = await r.uploader({ person: testPerson('Dan'), ...PARTY })({
      location: `/ordered/${ROOM}/blobs/`,
      sha256: RED_HASH,
      contentType: 'image/png',
      bytes: RED_PNG,
      signal: new AbortController().signal,
    });
    expect(stored).toEqual({ stored: true });
    send({ v: 3, type: 'changes', changes: [change] });
    await settle(4);
    expect(r.room.order).toBe(1);
    connection.close();
  });
});

describe('pictures in an ordered room: take-back, text boxes, the other screen', () => {
  test("a party's undo takes the picture back on every copy, and redo makes it again", async () => {
    const { r, ana, ben } = await room();
    await insertPicture(ben, 1, 0);
    expect(listXml(ana).match(/<w:drawing/g)).toHaveLength(2);
    ben.editor.exec({ type: 'undo' });
    await settle(16);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person).match(/<w:drawing/g)).toHaveLength(1);
    }
    expect(mediaNames(saved(r))).toEqual(['/word/media/red.png']);
    ben.editor.exec({ type: 'redo' });
    await settle(16);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toMatch(/<w:ins [^>]*w:author="Ben"[^>]*><w:r><w:drawing[ >]/);
    }
    expect(ben.refusals).toEqual([]);
  });

  test('the other copy loads the new picture, and the picture keeps its id there', async () => {
    const { ana, ben } = await room();
    await insertPicture(ben, 1, 0);
    await settle(16);
    caret(ana, 1, 1);
    for (let k = 0; k < 20 && ana.editor.snapshot().image?.resourceStatus === 'pending'; k += 1)
      await settle(2);
    const image = ana.editor.snapshot().image;
    expect(image?.resourceStatus).toBe('ready');
    caret(ben, 1, 1);
    expect(ben.editor.snapshot().image?.id).toBe(image?.id);
  });

  test('a page that joins reads every picture of the catch-up before its first document', async () => {
    const { r, ben } = await room();
    // Six pictures, each with its own bytes: more than the reads that run at once
    const pictures = Array.from({ length: 6 }, (_, k) => new Uint8Array([...RED_PNG, k]));
    // Apart, so the caret never touches the picture before
    for (const [k, data] of pictures.entries()) {
      await insertPicture(ben, 1, 2 * k, data);
    }
    expect(ben.refusals).toEqual([]);
    const handle = await createOrderedCollaboration({
      transport: r.transport({ person: testPerson('Cara'), ...ORGANISER }),
      room: ROOM,
      code: CODE,
      identity: { actorId: 'cara', name: 'Cara' },
      wordFile: async () => PICTURE_FILE,
      upload: r.uploader({ person: testPerson('Cara'), ...ORGANISER }),
      // A read that takes a moment, as a network read does
      media: async (_location, sha256) => {
        await new Promise((resolve) => setTimeout(resolve, 5));
        const bytes = r.blob(sha256);
        if (!bytes) throw new Error(`the room holds no blob ${sha256}`);
        return bytes;
      },
      schedule: null,
    });
    const held = new Set([...readPackage(handle.document).partBytes.values()].map(orderedBlobHash));
    handle.destroy();
    expect(pictures.filter((data) => held.has(orderedBlobHash(data)))).toHaveLength(6);
  });

  test("a party's text box is a tracked insertion on every copy: undone, rejected or accepted", async () => {
    const { r, ana, ben } = await room();
    const gone = () => {
      for (const person of [ana, ben]) {
        expect(sameAsRoom(person, r)).toBe(true);
        expect(listXml(person)).not.toContain('<wps:txbx>');
        expect(listXml(person).match(/<w:drawing[ >]/g)).toHaveLength(1);
      }
    };
    const suggest = async () => {
      caret(ben, 1, 3);
      expect(ben.editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
      await settle(16);
      expect(ben.refusals).toEqual([]);
      for (const person of [ana, ben]) {
        expect(sameAsRoom(person, r)).toBe(true);
        expect(listXml(person)).toMatch(
          /<w:ins [^>]*w:author="Ben"[^>]*><w:r><w:drawing[ >][\s\S]*<wps:txbx>/
        );
      }
    };
    await suggest();
    // The party takes it back
    ben.editor.exec({ type: 'undo' });
    await settle(16);
    gone();
    // The newest suggestion holds the text box (a taken-back one stays, emptied)
    await suggest();
    ana.handle.session.requestDecision(pendingOf(r, 'Ben').at(-1) ?? '', 'rejected');
    await settle(16);
    gone();
    await suggest();
    ana.handle.session.requestDecision(pendingOf(r, 'Ben').at(-1) ?? '', 'accepted');
    await settle(16);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toContain('<wps:txbx>');
      expect(listXml(person)).not.toContain('<w:ins');
    }
  });

  test('the organiser inserts a text box: one direct object on every copy, undone exactly', async () => {
    const { r, ana, ben } = await room();
    const before = listXml(ana);
    caret(ana, 1, 3);
    expect(ana.editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
    await settle(16);
    expect(ana.refusals).toEqual([]);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person)).toContain('<wps:txbx>');
    }
    ana.editor.exec({ type: 'undo' });
    await settle(16);
    expect(listXml(ben)).toBe(before);
    expect(sameAsRoom(ben, r)).toBe(true);
  });
});

/** A paragraph of the body, by index, as a script names it. */
function scriptParagraph(host: AutomationHost, index: number): AutomationHandle {
  const handleOf = (response: ReturnType<AutomationHost['execute']>): AutomationHandle => {
    const result = response.results[0];
    if (result?.status !== 'ok' || result.value.kind !== 'handle') throw new Error('no handle');
    return result.value.handle;
  };
  const document = handleOf(host.execute({ operations: [{ op: 'getDocument' }] }));
  const body = handleOf(host.execute({ operations: [{ op: 'getBody', document }] }));
  const listed = host.execute({ operations: [{ op: 'getParagraphs', body }] }).results[0];
  if (listed?.status !== 'ok' || listed.value.kind !== 'handles') throw new Error('no paragraphs');
  const paragraph = listed.value.handles[index];
  if (paragraph === undefined) throw new Error(`no paragraph ${index}`);
  return paragraph;
}

describe('pictures in an ordered room: a script', () => {
  test("the organiser's scripted picture (Office.js insertInlinePictureFromBase64) travels as Insert Picture's does", async () => {
    const { r, ana, ben } = await room();
    const before = listXml(ben);
    // The file holds the red picture as its own part; the room holds no blob of it yet
    expect(r.blob(RED_HASH)).toBeUndefined();
    const host = createBrowserAutomationHost(ana.editor);
    const response = host.execute({
      operations: [
        {
          op: 'insertInlinePicture',
          span: { paragraph: scriptParagraph(host, 1) },
          base64: Buffer.from(RED_PNG).toString('base64'),
          location: 'End',
        },
      ],
    });
    expect(response.ok).toBe(true);
    await settle(16);
    expect(ana.refusals).toEqual([]);
    // The bytes went to the room once, by hash, and every copy shows the picture
    expect(r.blob(RED_HASH)).toEqual(RED_PNG);
    for (const person of [ana, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person).match(/<w:drawing/g)).toHaveLength(2);
    }
    expect(
      mediaNames(saved(r)).filter((name) => name.startsWith('/word/media/ordered_'))
    ).toHaveLength(1);
    // A direct edit, undone exactly
    ana.editor.exec({ type: 'undo' });
    await settle(16);
    expect(listXml(ben)).toBe(before);
    expect(sameAsRoom(ben, r)).toBe(true);
    expect(mediaNames(saved(r))).toEqual(['/word/media/red.png']);
  });
});
