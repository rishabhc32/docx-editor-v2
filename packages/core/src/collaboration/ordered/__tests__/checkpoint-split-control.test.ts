// A control a pending Enter split, whose second part shows its prompt after a direct deletion of
// its words, across a close: a reject or a take-back of the Enter makes the parts one control
// again and takes the part's prompt away, on the room's next file as live.

import { describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import { apply, at, DIRECT, EDIT, textOf } from './change-apply-fixture.ts';
import { checkEveryCard, folded, mediaStore, wordMeaning } from './fold-check.ts';
import { anas } from './fold-shapes.ts';
import { docx, p, sdt, t } from './word-fixtures.ts';

const FILE = docx(
  p(t('Name: ') + sdt('<w:text w:multiLine="1"/>', t('Acme Corp')) + t(' end.')) + p(t('After.'))
);

/** The file opened, Ana's Enter in the control, and the organiser's deletion of its second part. */
function splitAndEmptied() {
  const opened = openOrderedWordFile(FILE, { file: 1 });
  const doc = opened.document;
  apply(doc, 'ana:1', { kind: 'split', ...EDIT, place: at(doc, 0, 10) });
  apply(doc, 'olive:1', { kind: 'delete', ...DIRECT, from: at(doc, 1, 0), to: at(doc, 1, 5) });
  expect([textOf(doc, 0), textOf(doc, 1)]).toEqual([
    'Name: Acme',
    'Click or tap here to enter text. end.',
  ]);
  return opened;
}

describe('a control a pending Enter split, its second part emptied', () => {
  test('a reject of the Enter after a close gives the file a reject gives live', async () => {
    const { document: doc, base } = splitAndEmptied();
    expect(await checkEveryCard(doc, base, mediaStore(base), 'a split control')).toBe(1);
  });

  test("Ana's take-back of the Enter after a close gives the file it gives live", async () => {
    const { document: doc, base } = splitAndEmptied();
    const blobs = mediaStore(base);
    const copy = await folded(doc, base, doc.file + 1, blobs);
    const takenBack = (d: OrderedDocument) => {
      apply(d, 'ana:2', { kind: 'retract', suggestions: anas(d) });
      return {
        texts: [textOf(d, 0), textOf(d, 1)],
        file: wordMeaning(orderedPackageBytes(d, base, { blobs })),
      };
    };

    const live = takenBack(doc);
    expect(live.texts).toEqual(['Name: Acme end.', 'After.']);
    expect(takenBack(copy.document)).toEqual(live);
  });
});
