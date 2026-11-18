// A room whose file is a checkpoint (design D14), with two real editors: each page reads the file's
// pictures through its media reader and shows them, a copy of one names its hash with no upload,
// and the room's document saves as a whole Word file with every picture.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OrderedMediaUploader } from '../../client/media-uploads.ts';
import { openOrderedCheckpoint, openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageMedia, writeOrderedCheckpoint } from '../../compose/checkpoint.ts';
import { orderedPackageBytes, writeOrderedPackage } from '../../compose/package.ts';
import { orderedBlobHash } from '../../protocol/blobs.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, RED_PNG, t } from '../word-fixtures.ts';
import { OBJECT_SOURCE, WORKBOOK } from './paste-objects-harness.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  join,
  listXml,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  settle,
  storiesOf,
  surface,
  type Person,
} from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const RED_HASH = orderedBlobHash(RED_PNG);
const RED_PART = '/word/media/red.png';

/** A picture of the file between two words, and a plain paragraph. */
const SOURCE = docx(
  p(t('Alpha ') + pictureRun(5) + t(' Omega')) + p(t('Plain words.')),
  MATRIX_EXTRAS
);

/** The bytes a streaming writer writes. */
async function collected(
  run: (write: (chunk: Uint8Array) => Promise<void>) => Promise<void>
): Promise<Uint8Array> {
  const chunks: Uint8Array[] = [];
  await run(async (chunk) => {
    chunks.push(chunk);
  });
  return new Uint8Array(Buffer.concat(chunks));
}

const noMedia = async (sha256: string): Promise<Uint8Array> => {
  throw new Error(`a checkpoint reads no media: ${sha256}`);
};

/**
 * A room made from the Word file as a host makes it: the file's media moved into the storage
 * once, then the first checkpoint. Every upload is counted.
 */
async function checkpointRoom(source = SOURCE) {
  const word = openOrderedWordFile(source, { file: 1 });
  const stored = new Map(orderedPackageMedia(word.base).map((m) => [m.sha256, m.bytes]));
  const checkpoint = await collected((write) =>
    writeOrderedCheckpoint(word.document, word.base, { media: noMedia, write })
  );
  const r = localRoom(checkpoint, { blobs: stored });
  const uploads: string[] = [];
  const counted = (author: string, login: typeof ORGANISER): OrderedMediaUploader => {
    const upload = r.uploader({ person: testPerson(author), ...login });
    return async (blob) => {
      uploads.push(blob.sha256);
      return upload(blob);
    };
  };
  const olive = await join(r, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: checkpoint,
    upload: counted('Olive', ORGANISER),
  });
  const ben = await join(r, 'Ben', {
    login: PARTY,
    wordFile: checkpoint,
    upload: counted('Ben', PARTY),
  });
  return { r, olive, ben, uploads, checkpoint };
}

/** The media parts of the person's editor package that hold the red picture. */
function redParts(bytes: Uint8Array | null, person?: Person): string[] {
  const pkg = person ? surface(person).session.currentPackage() : readPackage(bytes);
  return [...pkg.partBytes]
    .filter(([name, data]) => name.startsWith('/word/media/') && orderedBlobHash(data) === RED_HASH)
    .map(([name]) => name);
}

function readPackage(bytes: Uint8Array | null) {
  const read = readOoxmlPackage(bytes ?? new Uint8Array(), {});
  if (!read.ok) throw new Error(read.reason);
  return read.package;
}

const drawings = (xml: string) => xml.match(/<w:drawing>/g)?.length ?? 0;

/** Copy the first paragraph's picture in a person's editor, and paste it into the second. */
async function copyPicture(person: Person): Promise<void> {
  caretAt(person, 0, 6, 7);
  const { text, html } = surface(person).copyFlavours();
  caretAt(person, 1, 0);
  expect(surface(person).pasteRich(text, html)).toBe(true);
  await settle(24);
}

describe('a room whose file is a checkpoint', () => {
  test("every page reads the file's picture and shows it, from its first document on", async () => {
    const { r, olive, ben, uploads, checkpoint } = await checkpointRoom();
    expect(Object.keys(readPackage(checkpoint).partBytes)).not.toContain(RED_PART);
    for (const person of [olive, ben]) {
      expect(redParts(person.handle.document)).toEqual([RED_PART]);
      expect(redParts(null, person)).toEqual([RED_PART]);
      expect(drawings(editorXml(person))).toBe(1);
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, r)).toBe(true);
    }

    expect(uploads).toEqual([]);
  });

  test('a copy of the picture names its hash with no upload, and the room saves it whole', async () => {
    const { r, olive, ben, uploads, checkpoint } = await checkpointRoom();
    await copyPicture(ben);
    expect(ben.refusals).toEqual([]);
    expect(uploads).toEqual([]);
    expect(JSON.stringify(r.room.joinSnapshot(0))).toContain(RED_HASH);
    for (const person of [olive, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(drawings(listXml(person))).toBe(2);
      expect(redParts(null, person)).toHaveLength(2);
    }

    // The room's close: a real Word file with every picture, read from the storage
    const base = openOrderedCheckpoint(checkpoint, { file: 1 }).base;
    const reads: string[] = [];
    const saved = await collected((write) =>
      writeOrderedPackage(r.room.document, base, {
        media: async (sha256) => {
          reads.push(sha256);
          return r.blob(sha256) ?? new Uint8Array();
        },
        write,
      })
    );
    expect(reads).toEqual([RED_HASH, RED_HASH]);
    const blobs = (sha256: string) => r.blob(sha256);
    expect(Buffer.compare(saved, orderedPackageBytes(r.room.document, base, { blobs }))).toBe(0);
    expect(redParts(saved)).toHaveLength(2);
    const reopened = openOrderedWordFile(saved, { file: 2 });
    expect(drawings(storiesOf(reopened.document).join(''))).toBe(2);
  });

  test("a chart's copy names its workbook's hash with no upload", async () => {
    const { r, olive, ben, uploads } = await checkpointRoom(OBJECT_SOURCE);
    const workbook = orderedBlobHash(WORKBOOK);
    caretAt(olive, 0, 0, 'Chart \uFFFC here.'.length);
    const { text, html } = surface(olive).copyFlavours();
    caretAt(olive, 2, 0);
    expect(surface(olive).pasteRich(text, html)).toBe(true);
    await settle(24);
    expect(olive.refusals).toEqual([]);
    expect(uploads).toEqual([]);
    expect(JSON.stringify(r.room.joinSnapshot(0))).toContain(workbook);
    for (const person of [olive, ben]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(listXml(person).match(/<c:chart /g)).toHaveLength(2);
    }
  });

  test('a page whose read fails reports it by name, keeps the picture, and shows it after a rejoin', async () => {
    const { r, olive, checkpoint } = await checkpointRoom();
    let down = true;
    const cara = await join(r, 'Cara', {
      login: PARTY,
      wordFile: checkpoint,
      media: async (_location, sha256) => {
        if (down) {
          throw new Error('storage is down');
        }

        return r.blob(sha256) ?? new Uint8Array();
      },
    });
    expect(cara.refusals.map((x) => [x.reason, x.detail])).toEqual([['missing-blob', RED_HASH]]);
    expect(redParts(null, cara)).toEqual([]);
    expect(drawings(editorXml(cara))).toBe(1);

    // Edits go on around the picture, which stays in the room's document
    caretAt(cara, 1, 0);
    cara.editor.exec({ type: 'insertText', text: 'New ' });
    await settle(16);
    for (const person of [olive, cara]) {
      expect(sameAsRoom(person, r)).toBe(true);
      expect(drawings(listXml(person))).toBe(1);
    }

    // The storage is back: a rejoin reads the picture again, and the page shows it
    down = false;
    cara.handle.session.reconnect();
    await settle(24);
    expect(cara.refusals).toHaveLength(1);
    expect(redParts(null, cara)).toEqual([RED_PART]);
    expect(sameAsRoom(cara, r)).toBe(true);
  });
});
