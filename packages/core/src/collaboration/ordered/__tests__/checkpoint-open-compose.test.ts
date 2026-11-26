// An open checks that a checkpoint's list composes (design D14) and gives those parts to the
// document's first compose: the package a compose of its own gives, and never a part of a state
// the document has left. A package and a publish draw a story's part without its view.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedRoomFile, openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPart, storyPart } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord, type OrderedUndoRecord } from '../model/undo-record.ts';
import { BODY } from './change-apply-fixture.ts';
import { ROOM_FACTS, TABLE_RECORDS } from './checkpoint-fixture.ts';
import { resourceBlobs } from './checkpoint-resources-fixture.ts';
import { checkpointBytes } from './fold-check.ts';
import { testPerson } from './test-persons.ts';

/** Typing `text` at the start of the body's first text, as Ana's suggestion. */
function typed(doc: OrderedDocument, text: string): OrderedUndoRecord {
  const first = [...doc.stories.get(BODY)!.tree.items()].find((item) => item.kind === 'text')!;
  const change = {
    v: 1,
    id: 'ana:90',
    kind: 'insert',
    part: BODY,
    mode: 'suggest',
    place: { space: first.space, offset: first.offset },
    text,
  } as OrderedChange;
  const person = testPerson('Ana');
  return withUndoRecord(doc, () =>
    applyOrderedChange(doc, change, {
      person,
      date: '2026-11-27T00:00:00Z',
      ids: orderedIdMinter('ana:90'),
    })
  ).record;
}

const bodyText = (bytes: Uint8Array): string => strFromU8(unzipSync(bytes)['word/document.xml']!);

describe("an open's compose check", () => {
  for (const room of [ROOM_FACTS, TABLE_RECORDS]) {
    test(`${room.name}: gives the first package the parts a compose of its own gives`, async () => {
      const { document: doc, base } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      const blobs = resourceBlobs(base);
      const bytes = await checkpointBytes(doc, base, blobs);
      const kept = openOrderedRoomFile(bytes, { file: 2 });
      const first = orderedPackageBytes(kept.document, kept.base, { blobs });

      // A change and its take-back leave the list as it was, in a state of its own
      const own = openOrderedRoomFile(bytes, { file: 2 });
      takeBack(own.document, typed(own.document, 'Zq'));
      expect(orderedPackageBytes(own.document, own.base, { blobs })).toEqual(first);
    });
  }

  for (const room of [ROOM_FACTS, TABLE_RECORDS]) {
    test(`${room.name}: draws each part without its view as with it`, () => {
      const { document: doc } = openOrderedWordFile(room.source, { file: 1 });
      room.edit(doc);
      for (const name of [...doc.stories.keys()].filter((story) => !isTextBoxStoryName(story))) {
        const next = doc.composeIds.next;
        const { part, composed } = composeOrderedPart(doc, name);
        const after = doc.composeIds.next;
        doc.composeIds.next = next;
        const drawn = storyPart(doc, name);
        expect(serializeOoxmlPart(drawn.part)).toBe(serializeOoxmlPart(part));
        expect(drawn.prompts).toEqual(composed.prompts);
        expect(doc.composeIds.next).toBe(after);
      }
    });
  }

  test('composes the document again after a change', async () => {
    const { document: doc, base } = openOrderedWordFile(ROOM_FACTS.source, { file: 1 });
    ROOM_FACTS.edit(doc);
    const blobs = resourceBlobs(base);
    const opened = openOrderedRoomFile(await checkpointBytes(doc, base, blobs), { file: 2 });
    typed(opened.document, 'Zq');
    expect(bodyText(orderedPackageBytes(opened.document, opened.base, { blobs }))).toContain('Zq');
  });
});
