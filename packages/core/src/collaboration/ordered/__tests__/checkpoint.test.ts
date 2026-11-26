// Room checkpoints (design D14): a checkpoint holds no bytes of the media the room's storage holds,
// and names them by reference. It loads with no media bytes, and a document loaded from it
// composes and saves as one loaded from the Word file does, byte for byte. Version 5 is the Word
// file without its media parts and a manifest, read only (`versionFiveCase`); version 6 the item
// list (`checkpoint-snapshot`).

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedRoomFile,
  openOrderedWordFile,
  OrderedWordFileError,
  type OrderedWordFile,
} from '../client/open.ts';
import { orderedPackageMedia, writeOrderedCheckpoint } from '../compose/checkpoint.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import {
  composeOrderedPackage,
  orderedPackageBytes,
  writeOrderedPackage,
} from '../compose/package.ts';
import { translateOps } from '../client/translate.ts';
import { buildInlinePictureDrawing } from '../../../store/package/drawing-package-edit.ts';
import type { OoxmlDrawingNode } from '../../../store/package/ooxml-tree.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChangeMedia } from '../change/resources.ts';
import type { OrderedChange } from '../change/types.ts';
import { allowWritersPolicy } from '../room/policy.ts';
import { OrderedRoom } from '../room/room.ts';
import { apply, at, BODY, EDIT } from './change-apply-fixture.ts';
import { versionFiveCase } from './checkpoint-fixture.ts';
import { folded, heldEntries } from './fold-check.ts';
import { actor } from './room-harness.ts';
import { EMBEDDED_SOURCE, EMBEDDED_WORKBOOK, INK } from './client/paste-embedded-harness.ts';
import { OBJECT_SOURCE } from './client/paste-objects-harness.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, R, RED_PNG, t, W } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const FIXTURES = join(import.meta.dir, '../../../../../../e2e/fixtures');
const MANIFEST = 'ordered-checkpoint.json';

/** Room files with pictures of their own, charts and their workbooks, and embedded objects. */
const ROOM_FILES: ReadonlyArray<readonly [string, () => Uint8Array]> = [
  ['a picture', () => docx(p(t('Alpha ') + pictureRun(5) + t(' Omega')), MATRIX_EXTRAS)],
  ['charts and a diagram', () => OBJECT_SOURCE],
  ['an embedded workbook and ink', () => EMBEDDED_SOURCE],
  ...['images-formats.docx', 'images-tiff.docx', 'header-vml-emf.docx', 'demo.docx'].map(
    (name) => [name, () => new Uint8Array(readFileSync(join(FIXTURES, name)))] as const
  ),
];

/**
 * The room files with a version 5 case: each kind of media. demo.docx's case would hold its fonts
 * (1.2 MB), and its media are kinds the others hold.
 */
const VERSION_FIVE_FILES = ROOM_FILES.filter(([name]) => name !== 'demo.docx');

/** The name of a room file's version 5 case. */
const caseName = (name: string): string => name.replace(/\.docx$/, '').replaceAll(' ', '-');

/** PNG bytes of `size`, different for each `seed`. */
function pictureBytes(seed: number, size = 2048): Uint8Array {
  const bytes = new Uint8Array(size);
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  for (let k = 8; k < size; k += 1) {
    bytes[k] = (k * 31 + seed * 17) & 255;
  }

  return bytes;
}

/** Collect what a streaming writer writes. */
async function collected(run: (write: (chunk: Uint8Array) => Promise<void>) => Promise<void>) {
  const chunks: Uint8Array[] = [];
  await run(async (chunk) => {
    chunks.push(chunk);
  });
  return new Uint8Array(Buffer.concat(chunks));
}

/** A store of bytes by SHA-256, and a reader over it that counts its reads. */
function storeOf(media: Iterable<Uint8Array>) {
  const bytes = new Map([...media].map((data) => [orderedBlobHash(data), data]));
  const reads: string[] = [];
  return {
    bytes,
    reads,
    blobs: (sha256: string) => bytes.get(sha256),
    read: async (sha256: string) => {
      reads.push(sha256);
      const data = bytes.get(sha256);
      if (!data) {
        throw new Error(`no blob ${sha256}`);
      }

      return data;
    },
  };
}

/** A document's checkpoint, written with media read from `store`. */
function checkpointOf(file: OrderedWordFile, store = storeOf([])): Promise<Uint8Array> {
  return collected((write) =>
    writeOrderedCheckpoint(file.document, file.base, { media: store.read, write })
  );
}

/** Each version a room opens. */
const VERSIONS = [5, 6] as const;

/** The version 5 case of the Word file with its own embedded workbook and ink. */
const EMBEDDED_CASE = 'an-embedded-workbook-and-ink';

/**
 * A document's checkpoint of `version`, written with media read from `store`; of version 5, the
 * case `name` that the last version 5 writer wrote of the same document.
 */
const checkpointIn = async (
  version: (typeof VERSIONS)[number],
  name: string,
  file: OrderedWordFile,
  store = storeOf([])
): Promise<Uint8Array> => (version === 5 ? versionFiveCase(name) : checkpointOf(file, store));

/** The Word file `writeOrderedPackage` writes. */
const streamed = (file: OrderedWordFile, store: ReturnType<typeof storeOf>) =>
  collected((write) => writeOrderedPackage(file.document, file.base, { media: store.read, write }));

const ANA = { id: 'ana', name: 'Ana', org: '' };

/** A listed suggestion of Ana's that the file writes no mark for. */
const ANA_SUGGESTION = { author: 'Ana', attributes: [], person: ANA, revisions: [] };

function expectSameBytes(actual: Uint8Array, expected: Uint8Array): void {
  expect(actual.length).toBe(expected.length);
  expect(Buffer.compare(actual, expected)).toBe(0);
}

function reasonOf(open: () => unknown): string {
  try {
    open();
  } catch (error) {
    if (error instanceof OrderedWordFileError || error instanceof OrderedRefusal) {
      return error.reason;
    }

    throw error;
  }

  return 'opened';
}

describe('orderedPackageMedia lists the parts a checkpoint names', () => {
  test('pictures, workbooks and embedded files, with hash, type, size and bytes', () => {
    const { base } = openOrderedWordFile(EMBEDDED_SOURCE, { file: 1 });
    expect(orderedPackageMedia(base).map(({ bytes, ...rest }) => [rest, bytes.length])).toEqual([
      [
        {
          part: '/word/media/preview.png',
          contentType: 'image/png',
          sha256: orderedBlobHash(RED_PNG),
          size: RED_PNG.length,
        },
        RED_PNG.length,
      ],
      [
        {
          part: '/word/embeddings/Microsoft_Excel_Worksheet.xlsx',
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          sha256: orderedBlobHash(EMBEDDED_WORKBOOK),
          size: EMBEDDED_WORKBOOK.length,
        },
        EMBEDDED_WORKBOOK.length,
      ],
    ]);
  });

  test('fonts, the thumbnail, a .wdp, ink, an empty picture and other types stay inside', async () => {
    const source = docx(p(t('Alpha ') + pictureRun(5)), {
      parts: {
        'word/media/red.png': RED_PNG,
        'word/media/empty.png': new Uint8Array(0),
        'word/media/photo.wdp': pictureBytes(1),
        'word/fonts/font1.odttf': pictureBytes(2),
        'word/printerSettings/printerSettings1.bin': pictureBytes(3),
        'word/ink/ink1.xml': INK,
        'docProps/thumbnail.jpeg': pictureBytes(4),
      },
      contentTypes:
        '<Default Extension="png" ContentType="image/png"/>' +
        '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
        '<Default Extension="wdp" ContentType="image/vnd.ms-photo"/>' +
        '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/>' +
        '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.printerSettings"/>' +
        '<Override PartName="/word/ink/ink1.xml" ContentType="application/inkml+xml"/>',
      relationships: `<Relationship Id="rIdImg1" Type="${R}/image" Target="media/red.png"/>`,
    });
    const withThumbnail = withRootRelationship(
      source,
      '<Relationship Id="rIdThumb" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>'
    );
    const file = openOrderedWordFile(withThumbnail, { file: 1 });
    expect(orderedPackageMedia(file.base).map((media) => media.part)).toEqual([
      '/word/media/red.png',
    ]);

    const inside = Object.keys(heldEntries(await checkpointOf(file)));
    for (const name of [
      'word/media/empty.png',
      'word/media/photo.wdp',
      'word/fonts/font1.odttf',
      'word/printerSettings/printerSettings1.bin',
      'word/ink/ink1.xml',
      'docProps/thumbnail.jpeg',
    ]) {
      expect(inside).toContain(name);
    }

    expect(inside).not.toContain('word/media/red.png');
  });
});

/** `bytes` with one more relationship from the package root. */
function withRootRelationship(bytes: Uint8Array, relationship: string): Uint8Array {
  const entries = unzipSync(bytes);
  const rels = strFromU8(entries['_rels/.rels'] ?? new Uint8Array());
  entries['_rels/.rels'] = strToU8(
    rels.replace('</Relationships>', `${relationship}</Relationships>`)
  );
  return zipSync(entries);
}

for (const version of VERSIONS) {
  describe(`a document loaded from a checkpoint of version ${version}`, () => {
    for (const [name, source] of version === 5 ? VERSION_FIVE_FILES : ROOM_FILES) {
      test(`${name}: loads with no bytes, and saves as the Word file does`, async () => {
        const file = openOrderedWordFile(source(), { file: 1 });
        const store = storeOf(orderedPackageMedia(file.base).map((media) => media.bytes));
        const word = orderedPackageBytes(file.document, file.base);
        const checkpoint = await checkpointIn(version, caseName(name), file);
        const fromCheckpoint = openOrderedCheckpoint(checkpoint, { file: 2 });
        const fromWord = openOrderedWordFile(word, { file: 2 });
        const fileMedia = [...fromCheckpoint.document.resources.fileMedia.values()];
        expect(fileMedia.map((media) => media.part)).toEqual(
          orderedPackageMedia(fromWord.base).map((media) => media.part)
        );

        // The same Word file through either writer, and the media read once each, one at a time
        const expected = orderedPackageBytes(fromWord.document, fromWord.base);
        expectSameBytes(
          orderedPackageBytes(fromCheckpoint.document, fromCheckpoint.base, { blobs: store.blobs }),
          expected
        );
        expectSameBytes(await streamed(fromCheckpoint, store), expected);
        expect(store.reads).toEqual(fileMedia.map((media) => media.sha256));

        // The next close writes version 6, which opens to the same document
        await folded(fromCheckpoint.document, fromCheckpoint.base, 3, store.blobs);
      });
    }

    test('edits and added pictures compose alike, and added pictures are referenced', async () => {
      const source = docx(
        p(t('Alpha ') + pictureRun(5) + t(' Omega')) + p(t('Plain.')),
        MATRIX_EXTRAS
      );
      const file = openOrderedWordFile(source, { file: 1 });
      const added = [pictureBytes(7), pictureBytes(8)];
      const store = storeOf([RED_PNG, ...added]);
      added.forEach((bytes, k) => {
        const media = {
          sha256: orderedBlobHash(bytes),
          contentType: 'image/png' as const,
          size: bytes.length,
        };
        const object = { type: 'picture', media, extent: { cx: 914_400, cy: 457_200 } } as const;
        apply(file.document, `ben-1:${k + 1}`, {
          ...EDIT,
          kind: 'insert-object',
          place: at(file.document, 1, 0),
          object,
        });
      });
      apply(file.document, 'ben-1:3', {
        ...EDIT,
        kind: 'insert',
        place: at(file.document, 0, 2),
        text: 'xy',
      });
      const checkpoint = await checkpointIn(version, 'added-pictures', file, store);
      expect(store.reads).toEqual([]);
      const held = new Set(Object.values(heldEntries(checkpoint)).map(orderedBlobHash));
      expect([...store.bytes.keys()].filter((sha256) => held.has(sha256))).toEqual([]);

      const fromCheckpoint = openOrderedCheckpoint(checkpoint, { file: 2 });
      const fromWord = openOrderedWordFile(
        orderedPackageBytes(file.document, file.base, { blobs: store.blobs }),
        { file: 2 }
      );
      expectSameBytes(
        await streamed(fromCheckpoint, store),
        orderedPackageBytes(fromWord.document, fromWord.base, { blobs: store.blobs })
      );
    });

    test('with no bytes, a save refuses missing-blob and a screen leaves the picture out', async () => {
      const checkpoint = await checkpointIn(
        version,
        EMBEDDED_CASE,
        openOrderedWordFile(EMBEDDED_SOURCE, { file: 1 })
      );
      const file = openOrderedCheckpoint(checkpoint, { file: 2 });
      expect(reasonOf(() => orderedPackageBytes(file.document, file.base))).toBe('missing-blob');
      const screen = unzipSync(orderedPackageBytes(file.document, file.base, { missing: 'skip' }));
      expect(Object.keys(screen)).not.toContain('word/media/preview.png');
      expect(strFromU8(screen['word/_rels/document.xml.rels'] ?? new Uint8Array())).toContain(
        'media/preview.png'
      );
    });

    test('a change may name the media a checkpoint references: the room storage holds it', async () => {
      const file = openOrderedWordFile(docx(p(t('Alpha ') + pictureRun(5)), MATRIX_EXTRAS), {
        file: 1,
      });
      const checkpoint = await checkpointIn(version, 'a-picture-alone', file);
      const { document } = openOrderedCheckpoint(checkpoint, { file: 2 });
      const room = new OrderedRoom({
        document,
        applyChange: applyOrderedChange,
        policy: allowWritersPolicy,
        now: () => Date.UTC(2026, 10, 6),
        blobs: new Set<string>(),
      });
      const media = {
        sha256: orderedBlobHash(RED_PNG),
        contentType: 'image/png' as const,
        size: RED_PNG.length,
      };
      const ana = actor('Ana', 'ana-1', 'organiser');
      const picture = (id: string, sha256: string): OrderedChange => ({
        v: 1,
        id,
        ...EDIT,
        kind: 'insert-object',
        place: at(document, 0, 0),
        object: {
          type: 'picture',
          media: { ...media, sha256 },
          extent: { cx: 457_200, cy: 457_200 },
        },
      });
      expect(room.submit(ana, [picture('ana-1:1', media.sha256)]).refused).toEqual([]);
      const other = orderedBlobHash(pictureBytes(9));
      const refused = room.submit(ana, [picture('ana-1:2', other)]).refused;
      expect(refused.map((refusal) => refusal.code)).toEqual(['missing-blob']);
    });
  });

  describe(`a copy of the file's media names its reference, version ${version}`, () => {
    /** A checkpoint's document, and an `insertDrawing` of the file's picture into its second line. */
    async function copied(screenBytes: Uint8Array | undefined) {
      const source = docx(p(t('Alpha ') + pictureRun(5)) + p(t('Plain.')), MATRIX_EXTRAS);
      const word = openOrderedWordFile(source, { file: 1 });
      const checkpoint = await checkpointIn(version, 'a-picture-and-a-line', word);
      const { document: doc, base } = openOrderedCheckpoint(checkpoint, { file: 2 });
      const screen = composeOrderedPackage(doc, base, {
        blobs: () => screenBytes,
        missing: 'skip',
      });
      const line = orderedViewParagraph(doc, BODY, 1);
      if (!line) {
        throw new Error('no line');
      }

      const drawing = buildInlinePictureDrawing({
        docPrId: 900,
        relationshipId: 'rIdImg1',
        extentEmu: { cx: 457_200, cy: 457_200 },
      }) as OoxmlDrawingNode;
      const held: string[] = [];
      const batch = translateOps(
        [{ op: 'insertDrawing', paragraphId: line.id, offset: 0, drawing }],
        {
          doc,
          person: testPerson('Ana'),
          editorText: (nodeId) => (nodeId === line.id ? 'Plain.' : null),
          defaultParagraphStyle: null,
          gated: () => screen,
          hold: (bytes) => {
            held.push(orderedBlobHash(bytes));
            return { sha256: orderedBlobHash(bytes), contentType: 'image/png', size: bytes.length };
          },
        }
      );
      const [body] = batch.steps.flatMap((step): unknown[] => [...step()]);
      const object = (body as { object?: { media?: OrderedChangeMedia } } | undefined)?.object;
      return { media: object?.media ?? null, held };
    }

    const reference: OrderedChangeMedia = {
      sha256: orderedBlobHash(RED_PNG),
      contentType: 'image/png',
      size: RED_PNG.length,
    };

    test('before the page reads it, and after, with nothing held for an upload', async () => {
      expect(await copied(undefined)).toEqual({ media: reference, held: [] });
      expect(await copied(RED_PNG)).toEqual({ media: reference, held: [] });
    });

    test('other bytes under its part name are new media, held for an upload', async () => {
      const other = pictureBytes(5, RED_PNG.length);
      const { media, held } = await copied(other);
      expect(held).toEqual([orderedBlobHash(other)]);
      expect(media?.sha256).toBe(orderedBlobHash(other));
    });
  });
}

describe('a file of another kind is refused by name: version 5', () => {
  const embedded = () => openOrderedWordFile(EMBEDDED_SOURCE, { file: 1 });

  /** The checkpoint with its manifest replaced. */
  function withManifest(change: (manifest: Record<string, unknown>) => unknown): Uint8Array {
    const entries = unzipSync(versionFiveCase(EMBEDDED_CASE));
    const manifest = JSON.parse(strFromU8(entries[MANIFEST] ?? new Uint8Array())) as Record<
      string,
      unknown
    >;
    entries[MANIFEST] = strToU8(JSON.stringify(change(manifest)));
    return zipSync(entries);
  }

  const open = (bytes: Uint8Array) => () => openOrderedCheckpoint(bytes, { file: 2 });

  test('a checkpoint of another format or version', () => {
    expect(reasonOf(open(withManifest((m) => ({ ...m, format: 'other' }))))).toBe(
      'checkpoint-format'
    );
    for (const version of [1, 2, 3, 4, 6]) {
      expect(reasonOf(open(withManifest((m) => ({ ...m, version }))))).toBe('checkpoint-version');
    }
  });

  test('a manifest that does not read or does not fit the file', () => {
    const media = (m: Record<string, unknown>) => m.media as Record<string, unknown>[];
    const cases: ((m: Record<string, unknown>) => unknown)[] = [
      () => [1, 2],
      (m) => ({ ...m, media: 'all' }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, sha256: 'ABC' })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, size: 0 })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, part: '/word/../x.png' })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, part: '/word/document.xml' })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, contentType: 'image/gif' })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, entry: 0 })) }),
      (m) => ({ ...m, media: media(m).map((x) => ({ ...x, entry: 10_000 })) }),
      (m) => ({ ...m, media: [...media(m), ...media(m)] }),
      ({ suggestions: _, ...m }) => m,
      (m) => ({ ...m, suggestions: 'all' }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, person: { id: 'ana', name: 'Ana' } }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, revisions: 1 }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, revisions: [1] }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, author: 'Ben' }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, person: { ...ANA, org: 'Acme' } }] }),
      ({ comments: _, ...m }) => m,
      (m) => ({
        ...m,
        comments: [
          { id: '7', person: ANA },
          { id: '7', person: ANA },
        ],
      }),
      (m) => ({ ...m, takeovers: [{ label: 0 }] }),
      // A label on no element of the file
      (m) => ({
        ...m,
        suggestions: [ANA_SUGGESTION, ANA_SUGGESTION],
        takeovers: [
          { label: 0, taken: [{ suggestion: 1, properties: `<w:rPr xmlns:w="${W}"/>` }] },
        ],
      }),
      ({ namespaces: _, ...m }) => m,
      // A namespace twice, and no namespace
      (m) => ({
        ...m,
        namespaces: [
          { namespace: 'ana', person: ANA },
          { namespace: 'ana', person: ANA },
        ],
      }),
      (m) => ({ ...m, namespaces: [{ namespace: 'a b', person: ANA }] }),
      (m) => ({ ...m, untracked: [{ label: 0, target: 'p' }] }),
      (m) => ({
        ...m,
        suggestions: [ANA_SUGGESTION],
        untracked: [{ label: 0, target: 'tcPr', records: [{ suggestion: 0 }] }],
      }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, version: -1 }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, version: '2' }] }),
      // A direct edit is decided: never a pending suggestion, and the mark is a boolean
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, direct: true }] }),
      (m) => ({ ...m, suggestions: [{ ...ANA_SUGGESTION, status: 'accepted', direct: 1 }] }),
      (m) => ({ ...m, changeRecords: [{ label: 0, target: 'tcPr', records: [] }] }),
      // A label on no element of the file
      (m) => ({
        ...m,
        suggestions: [ANA_SUGGESTION],
        changeRecords: [{ label: 0, target: 'tcPr', records: [{ suggestion: 0, merge: [] }] }],
      }),
      // A key that reaches a prototype, in a known section or in one a reader skips
      (m) => ({ ...m, comments: [JSON.parse('{"__proto__": {"id": "7"}}')] }),
      (m) => ({ ...m, later: { constructor: 1 } }),
    ];
    for (const change of cases) {
      expect(reasonOf(open(withManifest(change)))).toBe('checkpoint-manifest');
    }

    // A namespace's person opens with no run that names it
    const named = withManifest((m) => ({
      ...m,
      namespaces: [{ namespace: 'ana', person: ANA }],
    }));
    expect(reasonOf(open(named))).toBe('opened');

    const entries = unzipSync(versionFiveCase(EMBEDDED_CASE));
    entries[MANIFEST] = strToU8('{"format":');
    expect(reasonOf(open(zipSync(entries)))).toBe('checkpoint-manifest');
  });

  test('a Word file is not a checkpoint, and a checkpoint is not a Word file', async () => {
    expect(reasonOf(open(EMBEDDED_SOURCE))).toBe('not-checkpoint');
    for (const version of VERSIONS) {
      const bytes = await checkpointIn(version, EMBEDDED_CASE, embedded());
      expect(reasonOf(() => openOrderedWordFile(bytes, { file: 2 }))).toBe('checkpoint');
    }
  });
});

describe("the version 5 manifest's name in another case", () => {
  test('opens as a checkpoint: OPC part names compare without case', () => {
    const bytes = versionFiveCase(EMBEDDED_CASE);
    const entries = unzipSync(bytes);
    entries[MANIFEST.toUpperCase()] = entries[MANIFEST]!;
    delete entries[MANIFEST];
    const renamed = zipSync(entries);
    const expected = openOrderedCheckpoint(bytes, { file: 2 });
    for (const open of [openOrderedCheckpoint, openOrderedRoomFile]) {
      const opened = open(renamed, { file: 2 });
      expect([...opened.base.partBytes.keys()]).toEqual([...expected.base.partBytes.keys()]);
      expect(opened.document.resources.fileMedia).toEqual(expected.document.resources.fileMedia);
    }
  });
});
