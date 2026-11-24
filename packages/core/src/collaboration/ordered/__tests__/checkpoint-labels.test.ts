// A checkpoint names the places of its facts by label, and who typed a run on the run (design
// D14): a later engine that reads one more element than this one, before a labelled place, reads
// every fact at its place.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { CHECKPOINT_NAMESPACE } from '../model/checkpoint.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedUnitTyper } from '../model/typers.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DIRECT,
  EDIT,
  FOOTNOTES,
  opens,
} from './change-apply-fixture.ts';
import { roomBlobs, TYPED_BARE, TYPED_RUNS, withRoomFacts } from './checkpoint-fixture.ts';
import { checkFold, checkpointBytes, folded, foldPrint, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

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

describe('a part that holds a typer', () => {
  // Found by the fold audit (run-2, quick-short-1): an edit after compose (a list control's last
  // value following its text) refused a part whose typer's prefix was not declared yet
  test("takes the engine's edits after compose: a combo box's last value follows its text", async () => {
    const combo =
      '<w:sdt><w:sdtPr><w:id w:val="7"/><w:comboBox w:lastValue="Acme">' +
      '<w:listItem w:displayText="Acme" w:value="Acme"/></w:comboBox></w:sdtPr>' +
      `<w:sdtContent>${t('Acme')}</w:sdtContent></w:sdt>`;
    const { document: doc, base } = openOrderedWordFile(docx(p(t('Party: ') + combo)), { file: 1 });

    // Olive's value, a record the last value follows, and Ana's letter typed in it, pending
    const control = addressOf(opens(doc, 'sdt')[0]!);
    const value = { kind: 'text', text: 'Beta' } as const;
    apply(doc, 'olive:1', { kind: 'control-value', ...DIRECT, control, value }, 'Olive');
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 9), text: 'X' }, 'Ana');
    const blobs = mediaStore(base);
    const lastValue = (bytes: Uint8Array) =>
      /w:lastValue="([^"]*)"/.exec(strFromU8(unzipSync(bytes)['word/document.xml']!))?.[1];
    expect(lastValue(orderedPackageBytes(doc, base, { blobs }))).toBe('BeXta');
    expect(lastValue(await checkpointBytes(doc, base, blobs))).toBe('BeXta');
    await checkFold(doc, base, blobs, 'a combo box');
  });
});

/** Each unit Olive typed, by its story and what it shows. */
function typedByOlive(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (const story of doc.stories.values()) {
    const where = isTextBoxStoryName(story.name) ? 'text box' : story.name;
    for (const item of story.tree.items()) {
      const node = item.node?.kind === 'textValue' ? '' : (item.node?.localName ?? '');
      for (let k = 0; k < item.length; k += 1) {
        if (orderedUnitTyper(doc, item, item.offset + k)?.namespace === 'olive') {
          // A complex field's unit has no node of its own
          out.push(`${where} ${node || (item.pieces === undefined ? item.text[k] : 'field')}`);
        }
      }
    }
  }

  return out.sort();
}

describe('who typed each unit kind', () => {
  for (const { name, source, edit, typed } of [
    {
      ...TYPED_RUNS,
      typed: [
        ...[...'XXYY', 'br', 'drawing', 'field', 'footnoteReference', 'tab'].map(
          (u) => `${BODY} ${u}`
        ),
        `${FOOTNOTES}  `,
        `${FOOTNOTES} N`,
        `${FOOTNOTES} footnoteRef`,
        'text box X',
        'text box Y',
      ],
    },
    { ...TYPED_BARE, typed: ['X', 'Y', 'fldSimple', 'oMath'].map((u) => `${BODY} ${u}`) },
  ]) {
    test(`${name}: the same after one fold and two`, async () => {
      const { document: doc, base } = openOrderedWordFile(source, { file: 1 });
      edit(doc);
      expect(typedByOlive(doc)).toEqual(typed);
      const blobs = roomBlobs(base);
      await checkFold(doc, base, blobs, name);
      const once = await folded(doc, base, 2, blobs);
      const twice = await folded(once.document, once.base, 3, blobs);
      expect([once.document, twice.document].map(typedByOlive)).toEqual([typed, typed]);
    });
  }
});
