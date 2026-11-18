// The room's Word file without comments, after real changes from two editors: a party's
// suggestion, and a comment on the organiser's selection.

import { afterEach, describe, expect, test } from 'bun:test';
import { unzipSync } from 'fflate';
import { orderedPackageBytes, writeOrderedPackage } from '../../compose/package.ts';
import { storyOf } from '../comments-fixture.ts';
import { expectSameFingerprint } from '../oracle.ts';
import { packageOf } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();
const text = (data: Uint8Array | undefined): string => new TextDecoder().decode(data);

describe("a room's Word file without comments", () => {
  test("the room's comment goes, the party's suggestion stays, and the room does not change", async () => {
    const room = localRoom();
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Now ');
    await settle();
    caretAt(olive, 0, 4, 7);
    expect(olive.handle.session.addSelectionComment('Why?')).not.toBeNull();
    await settle();
    expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert', 'comment']);

    const doc = room.room.document;
    const base = packageOf(FILE);
    const keep = orderedPackageBytes(doc, base);
    const kept = unzipSync(keep);
    expect(text(kept['word/comments.xml'])).toContain('Why?');
    expect(text(kept['word/document.xml'])).toContain('<w:commentRangeStart ');

    const chunks: Uint8Array[] = [];
    await writeOrderedPackage(doc, base, {
      media: async () => new Uint8Array(0),
      write: async (chunk) => {
        chunks.push(chunk);
      },
      comments: 'remove',
    });
    const remove = Buffer.concat(chunks);
    expect(Buffer.compare(remove, orderedPackageBytes(doc, base, { comments: 'remove' }))).toBe(0);

    const removed = unzipSync(remove);
    const named = Object.entries(removed).filter(([name, data]) =>
      /comment/i.test(`${name} ${text(data)}`)
    );
    expect(named.map(([name]) => name)).toEqual([]);
    const body = text(removed['word/document.xml']);
    expect(body).toContain('<w:ins ');
    expect(body).toContain('Now ');
    expectSameFingerprint(
      storyOf(remove, 'word/document.xml', false),
      storyOf(keep, 'word/document.xml', true)
    );

    // Both files written, the room and both editors are as they were
    expect(Buffer.compare(orderedPackageBytes(doc, base), keep)).toBe(0);
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
    }
  });
});
