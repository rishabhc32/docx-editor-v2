// Cards the editor alone decides one at a time, with real editors and one room: a format record
// whose runs reach into another person's suggested deletion, and a deletion over a paragraph mark
// whose paragraphs hold the author's own paragraph record. The organiser accepts or rejects one
// card; the room decides it as the editor alone does, the other card stays, every copy shows the
// same, and what the editors save reopens as the room's document.
//
// The room once refused both (`engine-refused`, `incomplete-group`): the engine planned the
// format record with the deletion, and the room planned a card of several suggestions as one
// batch, which would not drop the paragraph record with the mark.

import { afterEach, describe, expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint } from '../../../../store/package/ooxml-tree.ts';
import { orderedCardSeen } from '../../change/apply-decide.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();

const fingerprints = (doc: OrderedDocument): string[] =>
  [...doc.stories.keys()].map((name) =>
    canonicalOoxmlFingerprint(composeOrderedPart(doc, name).part)
  );

/** The kinds of the revision cards an editor shows. */
const kinds = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [item.revisionKind] : []));

async function seats() {
  const room = localRoom();
  const ana = await join(room, 'Ana', { review });
  const ben = await join(room, 'Ben', { review });
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
  return { room, ana, ben, olive };
}

/** The organiser decides the first card of `kind`; every copy then shows the room's document. */
async function decide(
  seated: Awaited<ReturnType<typeof seats>>,
  kind: string,
  decision: 'accept' | 'reject'
): Promise<void> {
  const card = seated.olive.editor
    .getReviewItems()
    .find((item) => item.kind === 'revision' && item.revisionKind === kind);
  const done =
    decision === 'accept'
      ? seated.olive.editor.acceptReviewItem(card?.key ?? '')
      : seated.olive.editor.rejectReviewItem(card?.key ?? '');
  expect(done.ok).toBe(true);
  await settle();
  const expected = fingerprints(seated.room.room.document);
  for (const person of [seated.ana, seated.ben, seated.olive]) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect(sameAsRoom(person, seated.room)).toBe(true);
    const saved = new Uint8Array(await person.editor.save());
    expect(fingerprints(openOrderedWordFile(saved, { file: 2 }).document)).toEqual(expected);
  }
}

describe('cards the editor alone decides one at a time', () => {
  for (const decision of ['accept', 'reject'] as const) {
    test(`a format record that reaches into another person's deletion, ${decision}ed`, async () => {
      const seated = await seats();
      caretAt(seated.ana, 0, 4, 11);
      expect(seated.ana.editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
      await settle();
      caretAt(seated.ben, 0, 8, 14);
      surface(seated.ben).deleteBackward();
      await settle();
      await decide(seated, 'format', decision);
      expect(kinds(seated.olive)).toEqual(['delete']);
      expect(editorXml(seated.olive)).not.toContain('rPrChange');
      expect(editorXml(seated.olive).includes('<w:b/>')).toBe(decision === 'accept');
    });

    test(`a deletion over a paragraph mark with the author's paragraph record, ${decision}ed`, async () => {
      const seated = await seats();
      caretAt(seated.ana, 0, 1, 2, 1);
      surface(seated.ana).setParagraphProperty('jc', { val: 'center' });
      await settle();
      caretAt(seated.ana, 0, 3, 2, 1);
      surface(seated.ana).deleteBackward();
      await settle();
      surface(seated.ana).insertPlainText('New');
      await settle();

      // Her deletion and her new word are one card of two suggestions
      const doc = seated.room.room.document;
      const typed = [...doc.suggestions.keys()].at(-1) ?? '';
      expect(orderedCardSeen(doc, typed)).toHaveLength(2);
      expect(kinds(seated.olive)).toContain('format');
      await decide(seated, kinds(seated.olive).find((kind) => kind !== 'format') ?? '', decision);
      expect(kinds(seated.olive)).toEqual(['format']);
    });
  }
});
