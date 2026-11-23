// A checkpoint names the places of its facts by label, and who typed a run on the run (design
// D14): a later engine that reads one more element than this one, before a labelled place, reads
// every fact at its place.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { openOrderedCheckpoint, openOrderedWordFile } from '../client/open.ts';
import { APPLY_FILE } from './change-apply-fixture.ts';
import { withRoomFacts } from './checkpoint-fixture.ts';
import { checkpointBytes, foldPrint, mediaStore } from './fold-check.ts';
import { p, t } from './word-fixtures.ts';

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

describe('an element a later engine reads before a labelled place', () => {
  test('moves no fact: the reopened print is the room with that element', async () => {
    const room = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const blobs = mediaStore(room.base);
    const edited = withRoomFacts(room.document);
    const checkpoint = withAdded(await checkpointBytes(edited, room.base, blobs));
    const opened = openOrderedCheckpoint(checkpoint, { file: 2 });

    // The same room over the Word file with the same elements
    const added = openOrderedWordFile(withAdded(APPLY_FILE), { file: 1 });
    expect(foldPrint(opened.document, opened.base, blobs)).toEqual(
      foldPrint(withRoomFacts(added.document), added.base, blobs)
    );
  });
});
