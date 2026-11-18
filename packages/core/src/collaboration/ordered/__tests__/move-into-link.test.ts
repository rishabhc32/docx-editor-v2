// A move whose destination is inside a link (the 50-page acceptance run, seed 1006): a party
// moved a picture into a table of contents' link, and the room's saved file did not reopen as
// the same document. The save wrote the destination as a bare `w:moveTo`, without its range
// markers, and the read of a file joined a `w:moveTo` that no range holds to the last move range
// of its author, though that range had ended. So the reopened file had the halves as two cards,
// under another move's date and Word id.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { decideOrdered, orderedCards } from '../decide.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import type { OrderedDocument } from '../model/types.ts';
import { APPLY_FILE, apply, at, EDIT, markOf, textOf } from './change-apply-fixture.ts';
import { checkSaveAndLoad } from './sessions/session.ts';
import { docx, p, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

/** Every shown paragraph's text of the body. */
function texts(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (let k = 0; textOf(doc, k) !== '<none>'; k += 1) out.push(textOf(doc, k));
  return out;
}

/** The texts after every card is decided one way; the document is restored. */
function decidedTexts(doc: OrderedDocument, decision: 'accept' | 'reject'): string[] {
  const { record } = withUndoRecord(doc, () => {
    for (const card of orderedCards(doc)) decideOrdered(doc, card.suggestions[0] ?? '', decision);
  });
  const out = texts(doc);
  takeBack(doc, record);
  return out;
}

/** "fee" of the first paragraph moved by Ana into the link of the second ("the terms"). */
function movedIntoLink() {
  const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const doc = opened.document;
  apply(doc, 'ana-1:1', {
    kind: 'move',
    ...EDIT,
    from: at(doc, 0, 4),
    to: at(doc, 0, 7),
    place: at(doc, 1, 6),
  });
  return { doc, bytes: orderedPackageBytes(doc, opened.base) };
}

describe('a move into a link', () => {
  test("is saved with its range markers on both halves, the destination's inside the link", () => {
    const { bytes } = movedIntoLink();
    const xml = strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array());
    const link = /<w:hyperlink[^>]*>(?:(?!<\/w:hyperlink>).)*<\/w:hyperlink>/.exec(xml)?.[0] ?? '';
    expect(link).toContain('<w:moveTo ');
    expect(link).toContain('<w:moveToRangeStart ');
    expect(link).toContain('<w:moveToRangeEnd ');
    expect(xml).toContain('<w:moveFromRangeStart ');
  });

  test('reopens as one card, and decides to the same text as the live document', () => {
    const { doc, bytes } = movedIntoLink();
    const again = openOrderedWordFile(bytes, { file: 2 }).document;
    expect([orderedCards(again).length, texts(again)]).toEqual([
      orderedCards(doc).length,
      texts(doc),
    ]);
    for (const decision of ['accept', 'reject'] as const)
      expect(decidedTexts(again, decision)).toEqual(decidedTexts(doc, decision));
  });
});

/** One change by Ana at its own date, as the room applies it. */
function applyAt(doc: OrderedDocument, id: string, body: object, date: string): void {
  const change = { v: 1, id, ...body } as OrderedChange;
  withUndoRecord(doc, () =>
    applyOrderedChange(doc, change, { person: testPerson('Ana'), date, ids: orderedIdMinter(id) })
  );
}

/** "fee", then "100", of the first paragraph moved by Ana into the control ("Acme"). */
function movedIntoControl() {
  const opened = openOrderedWordFile(APPLY_FILE, { file: 1 });
  const doc = opened.document;
  const move = (from: number, to: number, place: number) => ({
    kind: 'move',
    ...EDIT,
    from: at(doc, 0, from),
    to: at(doc, 0, to),
    place: at(doc, 2, place),
  });
  applyAt(doc, 'ana-1:1', move(4, 7, 9), '2026-10-28T00:00:34Z');
  applyAt(doc, 'ana-1:2', move(11, 14, 8), '2026-10-28T00:00:54Z');
  return { doc, base: opened.base, bytes: orderedPackageBytes(doc, opened.base) };
}

describe("moves into a control's content (sessions-paste seed 71)", () => {
  test("are saved with their range markers inside the control's content", () => {
    const { bytes } = movedIntoControl();
    const xml = strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array());
    const content = /<w:sdtContent>(?:(?!<\/w:sdtContent>).)*<\/w:sdtContent>/.exec(xml)?.[0];
    expect(content?.match(/<w:moveTo /g)?.length).toBe(2);
    expect(content?.match(/<w:moveToRangeStart /g)?.length).toBe(2);
    expect(content?.match(/<w:moveToRangeEnd /g)?.length).toBe(2);
  });

  test('reopen as two moves, each with its own Word id and date, and decide the same', () => {
    const { doc, base } = movedIntoControl();
    expect(orderedCards(doc).length).toBe(2);
    checkSaveAndLoad(doc, 'two moves into a control', base);
  });
});

const BOB = 'w:author="Bob" w:date="2026-10-28T12:00:00Z"';

describe("a file's w:moveTo that no range holds", () => {
  test("is not read as part of its author's move whose range ended before it", () => {
    const file = docx(
      p(
        `<w:moveFromRangeStart w:id="1" ${BOB} w:name="m1"/>` +
          `<w:moveFrom w:id="2" ${BOB}>${t('moved')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` +
          t(' stays.')
      ) +
        p(
          t('Here: ') +
            `<w:moveToRangeStart w:id="3" ${BOB} w:name="m1"/>` +
            `<w:moveTo w:id="4" ${BOB}>${t('moved')}</w:moveTo><w:moveToRangeEnd w:id="3"/>`
        ) +
        p(
          t('Later ') +
            '<w:moveTo w:id="5" w:author="Bob" w:date="2026-10-28T12:05:00Z">' +
            `${t('bare')}</w:moveTo>`
        )
    );
    const doc = openOrderedWordFile(file, { file: 1 }).document;
    // The named move is one card; the bare destination after it is another
    expect(orderedCards(doc).length).toBe(2);
  });
});

const CARA = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';

/** The body a document saves to. */
const savedBody = (doc: OrderedDocument, base: Parameters<typeof orderedPackageBytes>[1]) =>
  strFromU8(unzipSync(orderedPackageBytes(doc, base))['word/document.xml'] ?? new Uint8Array());

// Found by the editor acceptance cell (s2, seed 1000): the organiser rejected Cara's move, so its
// source words read plain again, and a party then deleted two of them. The save wrote the deletion
// as a bare `w:moveFrom` of the party's, which a reopen read as a move of its own and wrote again
// with range markers. A move's labels stay on its words after its decision; only a pending move
// strikes them as moved
describe("a deletion of a rejected move's words", () => {
  test('is saved as a plain deletion, and a save and reopen keeps it', () => {
    const file = docx(
      p(
        `<w:moveFromRangeStart w:id="1" ${CARA} w:name="move1"/>` +
          `<w:moveFrom w:id="2" ${CARA}>${t('Moved sentence.')}</w:moveFrom>` +
          `<w:moveFromRangeEnd w:id="1"/>${t(' Stays here.')}`
      ) +
        p(
          t('Destination: ') +
            `<w:moveToRangeStart w:id="3" ${CARA} w:name="move1"/>` +
            `<w:moveTo w:id="4" ${CARA}>${t('Moved sentence.')}</w:moveTo><w:moveToRangeEnd w:id="3"/>`
        )
    );
    const opened = openOrderedWordFile(file, { file: 1 });
    const doc = opened.document;
    decideOrdered(doc, orderedCards(doc)[0]?.suggestions[0] ?? '', 'reject');
    apply(
      doc,
      'cy-1:1',
      { kind: 'delete', ...EDIT, from: at(doc, 0, 13), to: at(doc, 0, 15) },
      'Cy'
    );
    const xml = savedBody(doc, opened.base);
    expect(xml).not.toContain('<w:moveFrom');
    expect(xml).toMatch(/<w:del [^>]*w:author="Cy"[^>]*><w:r>(?:(?!<\/w:del>).)*<w:delText>e\.</);
    checkSaveAndLoad(doc, "a rejected move's words deleted", opened.base);
  });

  test("is saved as a plain deletion of a moved paragraph's mark too", () => {
    const moved = (id: number) => `<w:rPr><w:moveFrom w:id="${id}" ${CARA}/></w:rPr>`;
    const into = (id: number) => `<w:rPr><w:moveTo w:id="${id}" ${CARA}/></w:rPr>`;
    const file = docx(
      p(
        `<w:moveFromRangeStart w:id="1" ${CARA} w:name="move1"/>` +
          `<w:moveFrom w:id="2" ${CARA}>${t('First.')}</w:moveFrom>`,
        moved(3)
      ) +
        p(
          `<w:moveFrom w:id="4" ${CARA}>${t('Second.')}</w:moveFrom>` +
            `<w:moveFromRangeEnd w:id="1"/>${t(' Stays.')}`
        ) +
        p(
          `<w:moveToRangeStart w:id="5" ${CARA} w:name="move1"/>` +
            `<w:moveTo w:id="6" ${CARA}>${t('First.')}</w:moveTo>`,
          into(7)
        ) +
        p(`<w:moveTo w:id="8" ${CARA}>${t('Second.')}</w:moveTo><w:moveToRangeEnd w:id="5"/>`)
    );
    const opened = openOrderedWordFile(file, { file: 1 });
    const doc = opened.document;
    for (const card of orderedCards(doc)) decideOrdered(doc, card.suggestions[0] ?? '', 'reject');
    expect(orderedCards(doc)).toEqual([]);
    apply(doc, 'cy-1:1', { kind: 'join', ...EDIT, mark: markOf(doc, 0) }, 'Cy');
    const xml = savedBody(doc, opened.base);
    expect(xml).not.toContain('<w:moveFrom');
    expect(xml).toMatch(/<w:pPr><w:rPr><w:del [^>]*w:author="Cy"/);
    checkSaveAndLoad(doc, "a rejected move's paragraph mark deleted", opened.base);
  });
});
