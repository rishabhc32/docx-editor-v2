// A checkpoint names the places of its facts by label, and who typed a run on the run (design
// D14): a later engine that reads one more element than this one, before a labelled place, reads
// every fact at its place.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import { writeOrderedCheckpoint } from '../compose/checkpoint.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  DIRECT,
  EDIT,
  markOf,
  opens,
  textOf,
} from './change-apply-fixture.ts';
import { foldPrint, mediaStore } from './fold-check.ts';
import { p, t } from './word-fixtures.ts';

/** The index of the shown body paragraph that starts with `text`. */
function paragraph(doc: OrderedDocument, text: string): number {
  for (let k = 0; textOf(doc, k) !== '<none>'; k += 1) {
    if (textOf(doc, k).startsWith(text)) {
      return k;
    }
  }

  throw new Error(`no paragraph ${text}`);
}

/**
 * Changes that leave a fact of every section past the fixture's first paragraph and first cell:
 * a takeover (Pia's press over Quin's), a row height in Ana's new table (an untracked record with
 * its table's cells), Ben's control value (an untracked record of `w:sdtPr`), and Olive's words
 * typed inside Ana's insertion (typers). Each is found by what it shows, after the added elements.
 */
function edited(doc: OrderedDocument): OrderedDocument {
  for (const [id, who, property] of [
    ['quin:1', 'Quin', { localName: 'jc', attributes: { val: 'center' } }],
    ['pia:1', 'Pia', { localName: 'spacing', attributes: { after: '240' } }],
  ] as const) {
    const marks = [markOf(doc, paragraph(doc, 'Signed'))];
    apply(doc, id, { kind: 'paragraph-property', ...EDIT, marks, properties: [property] }, who);
  }

  const due = () => paragraph(doc, 'Due');
  const table = { rows: 1, columns: 2, columnWidth: 3000 };
  apply(doc, 'ana:1', { kind: 'insert-table', ...EDIT, place: at(doc, due(), 0), ...table });
  const row = addressOf(opens(doc, 'tr').at(-1)!);
  const height = { op: 'setTableRowHeight', heightTwips: 480 } as const;
  apply(doc, 'ana:2', { kind: 'table-property', ...EDIT, on: [row], op: height });
  apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, due(), 4), text: 'new words' });
  apply(
    doc,
    'olive:1',
    { kind: 'insert', ...DIRECT, place: at(doc, due(), 7), text: 'XY' },
    'Olive'
  );
  const control = addressOf(opens(doc, 'sdt').at(-1)!);
  const value = { kind: 'text', text: 'Beta Ltd' } as const;
  apply(doc, 'ben:1', { kind: 'control-value', ...EDIT, control, value }, 'Ben');
  return doc;
}

/**
 * A package's main document with a control and a unit at the end of its first paragraph, and a
 * table at the end of its first cell: what a later preparation or unit rule might add before the
 * labelled places. No node before them moves.
 */
function withAdded(bytes: Uint8Array): Uint8Array {
  const entries = unzipSync(bytes);
  const control =
    '<w:sdt><w:sdtPr><w:id w:val="9"/><w:text/></w:sdtPr>' +
    `<w:sdtContent>${t('Z')}</w:sdtContent></w:sdt>${t('Z')}`;
  const cell = `<w:tc>${p(t('Z'))}</w:tc>`;
  const table = `<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr>${cell}</w:tr></w:tbl>`;
  const xml = strFromU8(entries['word/document.xml']!)
    .replace('</w:p>', `${control}</w:p>`)
    .replace('</w:tc>', `${table}${p('')}</w:tc>`);
  entries['word/document.xml'] = strToU8(xml);
  return zipSync(entries);
}

async function checkpointOf(doc: OrderedDocument, base: Parameters<typeof mediaStore>[0]) {
  const blobs = mediaStore(base);
  const chunks: Uint8Array[] = [];
  await writeOrderedCheckpoint(doc, base, {
    media: async (sha256) => blobs(sha256)!,
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return new Uint8Array(Buffer.concat(chunks));
}

describe('an element a later engine reads before a labelled place', () => {
  test('moves no fact: the reopened print is the room with that element', async () => {
    const room = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const checkpoint = withAdded(await checkpointOf(edited(room.document), room.base));
    const opened = openOrderedCheckpoint(checkpoint, { file: 2 });

    // The same room over the Word file with the same elements
    const added = openOrderedWordFile(withAdded(APPLY_FILE), { file: 1 });
    const blobs = mediaStore(added.base);
    expect(foldPrint(opened.document, opened.base, blobs)).toEqual(
      foldPrint(edited(added.document), added.base, blobs)
    );
  });
});
