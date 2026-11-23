// A checkpoint names the places of its facts by label, and who typed a run on the run (design
// D14): a later engine that reads one more element than this one, before a labelled place, reads
// every fact at its place.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import type { OrderedChangePlace } from '../change/types.ts';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { CHECKPOINT_NAMESPACE } from '../model/checkpoint.ts';
import { replaceItems } from '../model/mutations.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DIRECT,
  EDIT,
  textOf,
} from './change-apply-fixture.ts';
import { paragraph, withRoomFacts } from './checkpoint-fixture.ts';
import { checkFold, checkpointBytes, folded, foldPrint, mediaStore } from './fold-check.ts';
import { bodyOf, docx, MATRIX_ITEMS, p, RED_PNG, t } from './word-fixtures.ts';

/** A package with one of its entries edited as text. */
function crafted(bytes: Uint8Array, entry: string, edit: (text: string) => string): Uint8Array {
  const entries = unzipSync(bytes);
  entries[entry] = strToU8(edit(strFromU8(entries[entry]!)));
  return zipSync(entries);
}

/**
 * A package's main document with a control and a unit at the end of its first paragraph, and a
 * table at the end of its first cell: what a later preparation or unit rule might add before the
 * labelled places. No node before them moves.
 */
function withAdded(bytes: Uint8Array): Uint8Array {
  const control =
    '<w:sdt><w:sdtPr><w:id w:val="9"/><w:text/></w:sdtPr>' +
    `<w:sdtContent>${t('Z')}</w:sdtContent></w:sdt>${t('Z')}`;
  const cell = `<w:tc>${p(t('Z'))}</w:tc>`;
  const table = `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell}</w:tr></w:tbl>`;
  return crafted(bytes, 'word/document.xml', (xml) =>
    xml.replace('</w:p>', `${control}</w:p>`).replace('</w:tc>', `${table}${p('')}</w:tc>`)
  );
}

/** The room of `withRoomFacts` over APPLY_FILE, its checkpoint and its storage. */
async function roomFacts() {
  const room = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const blobs = mediaStore(room.base);
  const doc = withRoomFacts(room.document);
  return { doc, base: room.base, blobs, checkpoint: await checkpointBytes(doc, room.base, blobs) };
}

/** Whether any part of a Word file holds the checkpoint's namespace. */
const holdsNamespace = (bytes: Uint8Array): boolean =>
  Object.values(unzipSync(bytes)).some((entry) => strFromU8(entry).includes(CHECKPOINT_NAMESPACE));

/** The reason a checkpoint does not open, or `opens`. */
function openReason(bytes: Uint8Array): string {
  try {
    openOrderedCheckpoint(bytes, { file: 2 });
    return 'opens';
  } catch (error) {
    return error instanceof OrderedWordFileError ? error.reason : String(error);
  }
}

describe('an element a later engine reads before a labelled place', () => {
  test('moves no fact: the reopened print is the room with that element', async () => {
    const { checkpoint, blobs } = await roomFacts();
    const opened = openOrderedCheckpoint(withAdded(checkpoint), { file: 2 });

    // The same room over the Word file with the same elements
    const added = openOrderedWordFile(withAdded(APPLY_FILE), { file: 1 });
    expect(foldPrint(opened.document, opened.base, blobs)).toEqual(
      foldPrint(withRoomFacts(added.document), added.base, blobs)
    );
  });
});

describe("a checkpoint's namespace", () => {
  test('no Word file holds it: after a fold, after a reopen, and after a further edit', async () => {
    const { checkpoint, blobs } = await roomFacts();
    expect(holdsNamespace(checkpoint)).toBe(true);
    const once = openOrderedCheckpoint(checkpoint, { file: 2 });
    const twice = await folded(once.document, once.base, 3, blobs);
    for (const copy of [once, twice]) {
      expect(holdsNamespace(orderedPackageBytes(copy.document, copy.base, { blobs }))).toBe(false);
      apply(
        copy.document,
        'olive:9',
        { kind: 'insert', ...EDIT, place: at(copy.document, 0, 3), text: 'Q' },
        'Olive'
      );
      expect(holdsNamespace(orderedPackageBytes(copy.document, copy.base, { blobs }))).toBe(false);
    }
  });

  test('a Word file that holds it opens as the same file without it, and folds again', async () => {
    const { doc, base, blobs } = await roomFacts();
    const plain = orderedPackageBytes(doc, base, { blobs });
    const marked = crafted(plain, 'word/document.xml', (xml) =>
      xml
        .replace('<w:document ', `<w:document xmlns:oc="${CHECKPOINT_NAMESPACE}" `)
        .replaceAll('<w:p ', '<w:p oc:label="0" ')
        .replaceAll('<w:r>', '<w:r oc:typer="zed">')
    );
    expect(holdsNamespace(marked)).toBe(true);
    const upload = openOrderedWordFile(marked, { file: 1 });
    const same = openOrderedWordFile(plain, { file: 1 });
    expect(upload.document.namespaces.size).toBe(0);
    expect(foldPrint(upload.document, upload.base, blobs)).toEqual(
      foldPrint(same.document, same.base, blobs)
    );
    await checkFold(upload.document, upload.base, blobs, 'an upload with marks');
  });

  test('a label twice, missing, misplaced or on another container, and an unknown typer, are refused', async () => {
    const { checkpoint } = await roomFacts();
    const inDocument = (edit: (xml: string) => string) =>
      openReason(crafted(checkpoint, 'word/document.xml', edit));
    expect(openReason(checkpoint)).toBe('opens');
    expect({
      // A cell of the new table's label on a cell of the file's table too
      twice: inDocument((xml) => xml.replace('<w:tc>', '<w:tc oc:label="1">')),
      missing: inDocument((xml) => xml.replace(' oc:label="4"', '')),
      misplaced: inDocument((xml) =>
        xml.replace(' oc:label="4"', '').replace('<w:p ', '<w:p oc:label="4" ')
      ),
      container: inDocument((xml) =>
        xml
          .replace('oc:label="0"', 'oc:label="x"')
          .replace('oc:label="3"', 'oc:label="0"')
          .replace('oc:label="x"', 'oc:label="3"')
      ),
      typer: inDocument((xml) => xml.replace('oc:typer="olive"', 'oc:typer="zed"')),
    }).toEqual({
      twice: 'checkpoint-manifest',
      missing: 'checkpoint-manifest',
      misplaced: 'checkpoint-manifest',
      container: 'checkpoint-manifest',
      typer: 'checkpoint-manifest',
    });
  });

  test('a manifest with a section and a field this engine does not know opens to the same print', async () => {
    const { checkpoint, blobs } = await roomFacts();
    const print = (bytes: Uint8Array) => {
      const opened = openOrderedCheckpoint(bytes, { file: 2 });
      return foldPrint(opened.document, opened.base, blobs);
    };
    const inManifest = (edit: (json: string) => string) =>
      crafted(checkpoint, 'ordered-checkpoint.json', edit);
    const added = inManifest((json) =>
      json
        .replace('"suggestions":[{', '"later":{"rows":[1,2]},"suggestions":[{"later":true,')
        .replace('"takeovers":[{', '"takeovers":[{"later":"x",')
        .replace('"untracked":[{', '"untracked":[{"later":[],')
    );
    expect(print(added)).toEqual(print(checkpoint));
    for (const key of ['__proto__', 'constructor', 'prototype']) {
      const forbidden = inManifest((json) =>
        json.replace('"suggestions":[', `"later":{"${key}":{}},"suggestions":[`)
      );
      expect({ key, reason: openReason(forbidden) }).toEqual({
        key,
        reason: 'checkpoint-manifest',
      });
    }
  });
});

describe('a run of two typers', () => {
  // No edit draws one today: a run's units are one change's (`i:<change>`) or one loaded run's.
  // The state is made by hand, as a later edit might make it.
  test('is cut in the checkpoint and in the Word file alike, so both open to the same addresses', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const place = insideAnas(doc);
    apply(doc, 'olive:1', { kind: 'insert', ...DIRECT, place, text: 'XY' }, 'Olive');
    const due = paragraph(doc, 'Due');
    apply(doc, 'pia:1', { kind: 'insert', ...DIRECT, place: at(doc, due, 8), text: 'Z' }, 'Pia');
    expect(textOf(doc, due)).toBe('Due newXZY wordsin 30 days.');

    // Pia's letter in the run of Olive's, with its face
    const story = doc.stories.get(BODY)!;
    const olive = story.spaces.get('i:olive:1.1')![0]!;
    const pia = story.spaces.get('i:pia:1.1')![0]!;
    replaceItems(
      doc,
      story,
      [pia],
      [{ ...pia, run: olive.run, face: olive.face, leaf: undefined }]
    );

    const blobs = mediaStore(base);
    const runs = (bytes: Uint8Array) =>
      [...strFromU8(unzipSync(bytes)['word/document.xml']!).matchAll(/<w:r[ >].*?<\/w:r>/g)]
        .map(([run]) => run.replace(/<[^>]*>/g, ''))
        .filter((text) => /[XYZ]/.test(text));
    expect(runs(orderedPackageBytes(doc, base, { blobs }))).toEqual(['X', 'Z', 'Y']);
    expect(runs(await checkpointBytes(doc, base, blobs))).toEqual(['X', 'Z', 'Y']);
    await checkFold(doc, base, blobs, 'two typers');
  });
});

const PNG_MEDIA = {
  sha256: orderedBlobHash(RED_PNG),
  contentType: 'image/png',
  size: RED_PNG.byteLength,
} as const;

/** A matrix item's file, and its item: in its second paragraph, after `Alpha `. */
const matrixFile = (key: string) => docx(bodyOf(MATRIX_ITEMS.find((i) => i.key === key)!));

/** The address of the first item drawn from an element of this local name. */
const itemOf = (doc: OrderedDocument, localName: string, story = BODY) =>
  addressOf(
    [...doc.stories.get(story)!.tree.items()].find(
      (item) => item.node?.kind !== 'textValue' && item.node?.localName === localName
    )!
  );

/** Ana's pending words in APPLY_FILE's `Due` paragraph, and a place inside them. */
function insideAnas(doc: OrderedDocument): OrderedChangePlace {
  const due = paragraph(doc, 'Due');
  apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, due, 4), text: 'new words' });
  return at(doc, due, 7);
}

/**
 * Units of each kind a room change types inside a pending insertion, and the element each is
 * drawn in: text and atoms in a run, a field's runs, a bare unit (a simple field, an equation:
 * an element kept whole is drawn as one), units hidden by their run, and units of a new cell
 * and a new text box (both of its copies). Olive types them. A prompt's units are never typed:
 * a control comes in by a direct edit only, accepted at once.
 */
const UNIT_KINDS: readonly (readonly [string, Uint8Array, (doc: OrderedDocument) => void])[] = [
  [
    'text',
    APPLY_FILE,
    (doc) =>
      apply(
        doc,
        'olive:1',
        { kind: 'insert', ...DIRECT, place: insideAnas(doc), text: 'XY' },
        'Olive'
      ),
  ],
  [
    'a tab',
    APPLY_FILE,
    (doc) => {
      apply(doc, 'olive:1', { kind: 'insert-tab', ...DIRECT, place: insideAnas(doc) }, 'Olive');
    },
  ],
  [
    'a line break',
    APPLY_FILE,
    (doc) => {
      const place = insideAnas(doc);
      apply(doc, 'olive:1', { kind: 'insert-break', ...DIRECT, place, break: 'line' }, 'Olive');
    },
  ],
  [
    'a picture',
    APPLY_FILE,
    (doc) => {
      const object = {
        type: 'picture',
        media: PNG_MEDIA,
        extent: { cx: 914_400, cy: 457_200 },
      } as const;
      apply(
        doc,
        'olive:1',
        { kind: 'insert-object', ...DIRECT, place: insideAnas(doc), object },
        'Olive'
      );
    },
  ],
  [
    'a note reference and its note',
    APPLY_FILE,
    (doc) => {
      const body = { kind: 'insert-note', note: 'footnote', text: ' N' } as const;
      apply(doc, 'olive:1', { ...body, ...EDIT, place: insideAnas(doc) }, 'Olive');
    },
  ],
  [
    'a complex field',
    APPLY_FILE,
    (doc) =>
      apply(
        doc,
        'olive:1',
        { kind: 'insert-field', ...DIRECT, place: insideAnas(doc), field: 'PAGE' },
        'Olive'
      ),
  ],
  [
    'a simple field',
    docx(p(t('Page ') + `<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple>`)),
    (doc) => {
      const field = itemOf(doc, 'fldSimple');
      apply(doc, 'olive:1', { kind: 'field-code', ...EDIT, field, code: 'NUMPAGES' }, 'Olive');
    },
  ],
  [
    'an equation',
    matrixFile('math'),
    (doc) => {
      const object = itemOf(doc, 'oMath');
      apply(
        doc,
        'olive:1',
        { kind: 'equation', ...EDIT, object, equation: 0, linear: 'x^2' },
        'Olive'
      );
    },
  ],
  [
    'hidden text',
    matrixFile('hidden'),
    (doc) =>
      apply(doc, 'olive:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 8), text: 'XY' }, 'Olive'),
  ],
  [
    'a new cell',
    APPLY_FILE,
    (doc) => {
      const cell = itemOf(doc, 'tc');
      apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, cell, where: 'right' });

      // The new cell, right of the first
      const place = at(doc, 5, 0);
      apply(doc, 'olive:1', { kind: 'insert', ...DIRECT, place, text: 'XY' }, 'Olive');
    },
  ],
  [
    'a new text box',
    APPLY_FILE,
    (doc) => {
      const object = { type: 'textbox', extent: { cx: 1_270_000 } } as const;
      apply(doc, 'ana:1', { kind: 'insert-object', ...EDIT, place: at(doc, 0, 2), object });
      const box = [...doc.stories.keys()].find(isTextBoxStoryName)!;
      const place = at(doc, 0, 0, box);
      apply(
        doc,
        'olive:1',
        { kind: 'insert', part: box, mode: 'edit', place, text: 'XY' },
        'Olive'
      );
    },
  ],
];

/** How many units of the document Olive typed, by `orderedUnitTyper`. */
function typedByOlive(doc: OrderedDocument): number {
  let count = 0;
  for (const story of doc.stories.values()) {
    for (const item of story.tree.items()) {
      for (let k = 0; k < item.length; k += 1) {
        if (orderedUnitTyper(doc, item, item.offset + k)?.namespace === 'olive') {
          count += 1;
        }
      }
    }
  }

  return count;
}

describe('who typed each unit kind', () => {
  for (const [label, file, edit] of UNIT_KINDS) {
    test(`${label}: the same after one fold and two`, async () => {
      const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
      edit(doc);
      const typed = typedByOlive(doc);
      expect(typed).toBeGreaterThan(0);
      const store = mediaStore(base);
      const blobs = (hash: string) => (hash === PNG_MEDIA.sha256 ? RED_PNG : store(hash));
      await checkFold(doc, base, blobs, label);
      const once = await folded(doc, base, 2, blobs);
      const twice = await folded(once.document, once.base, 3, blobs);
      expect([once.document, twice.document].map(typedByOlive)).toEqual([typed, typed]);
    });
  }
});
