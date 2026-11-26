// A checkpoint of version 6 saves the room's item list (design D14): an open gives back every
// stored field, the same Word file and the same bytes on the next save.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-serialize.ts';
import type { OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import { writeZip } from '../../../store/package/zip.ts';
import { normalizedRoomPackage, openOrderedRoomFile, openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { snapshotDifference } from '../model/snapshot-compare.ts';
import {
  PASTED_LIST,
  ROOM_FACTS,
  SPLIT_RUN,
  STYLE_RECORDS,
  TABLE_RECORDS,
  TYPED_BARE,
  TYPED_RUNS,
  type CheckpointRoom,
} from './checkpoint-fixture.ts';
import { addressOf, apply, EDIT, opens } from './change-apply-fixture.ts';
import { RESOURCE_ROOMS, resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkpointBytes, folded, wordMeaning } from './fold-check.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, sdt, t } from './word-fixtures.ts';

const ROOMS: readonly CheckpointRoom[] = [
  ROOM_FACTS,
  STYLE_RECORDS,
  TYPED_RUNS,
  TYPED_BARE,
  SPLIT_RUN,
  PASTED_LIST,
  TABLE_RECORDS,
  ...RESOURCE_ROOMS,
];

describe('a checkpoint of version 6', () => {
  for (const room of ROOMS) {
    test(`${room.name}: an open gives back every stored field and the same file`, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      const blobs = resourceBlobs(base);

      // Under the room's next file number: the document keeps the number of its addresses
      const next = await folded(doc, base, 2, blobs);
      const live = orderedPackageBytes(doc, base, { blobs });
      const opened = orderedPackageBytes(next.document, next.base, { blobs });
      expect(wordMeaning(opened)).toEqual(wordMeaning(live));

      // The first save writes the base as the package writer writes it; from then on, the same
      const again = openOrderedRoomFile(await checkpointBytes(next.document, next.base, blobs), {
        file: 3,
      });
      expect(orderedPackageBytes(again.document, again.base, { blobs })).toEqual(opened);
    });
  }
});

/** Each node id of a tree, in document order. */
const idsOf = (node: OoxmlNode): string[] =>
  node.kind === 'textValue' ? [node.id] : [node.id, ...node.children.flatMap(idsOf)];

/** Each XML part of a package by its fingerprint and its node ids. */
const treesOf = (pkg: OoxmlPackage) =>
  Object.fromEntries(
    [...pkg.parts].map(([name, part]) => [
      name,
      { fingerprint: canonicalOoxmlFingerprint(part), ids: idsOf(part.root) },
    ])
  );

describe('a room opened from a Word file', () => {
  // The open mints paragraph ids, puts a text box's fallback copy in step and gives an empty
  // control its prompt; a filled control names its prompt's glossary block
  test('keeps the bytes as uploaded, which every open normalizes to the same trees and node ids', async () => {
    const textBox = MATRIX_ITEMS.find(({ key }) => key === 'textbox')!;
    const control = p(t('Fee: ') + sdt('<w:id w:val="5"/><w:text/>', '') + t(' due.'));
    const file = docx(bodyOf(textBox) + control, MATRIX_EXTRAS);
    const { document: doc, base } = openOrderedWordFile(file, { file: 1 });
    const uploaded = readOoxmlPackage(file, {});
    if (!uploaded.ok) {
      throw new Error(uploaded.reason);
    }

    // The editor alone reads these bytes too
    expect(base.partBytes).toEqual(uploaded.package.partBytes);
    const read = readOoxmlPackage(writeZip(base.partBytes), {});
    if (!read.ok) {
      throw new Error(read.reason);
    }

    expect(treesOf(normalizedRoomPackage(read.package))).toEqual(treesOf(base));

    // Live and after a close: the same item list, node ids and glossary block names
    const sdtItem = opens(doc, 'sdt')[0]!;
    const value = { kind: 'text', text: 'Acme' } as const;
    apply(doc, 'ana:1', { kind: 'control-value', ...EDIT, control: addressOf(sdtItem), value });
    const blobs = resourceBlobs(base);
    const opened = openOrderedRoomFile(await checkpointBytes(doc, base, blobs), { file: 2 });
    expect(snapshotDifference(doc, opened.document)).toBeNull();
    expect(treesOf(opened.base)).toEqual(treesOf(base));
    const live = orderedPackageBytes(doc, base, { blobs });
    expect(wordMeaning(live)).toHaveProperty(['/word/glossary/document.xml']);
    expect(orderedPackageBytes(opened.document, opened.base, { blobs })).toEqual(live);
  });
});
