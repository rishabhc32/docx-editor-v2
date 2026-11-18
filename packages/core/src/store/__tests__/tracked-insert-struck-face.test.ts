// Typing BESIDE another author's strike keeps the face of the text it continues.
//
// One author strikes words; another types inside, right after or right before them. The
// insertion goes beside the `w:del`, never into it, and the walk that places it holds no run
// of its own there — so the typed words came out in a bare `w:r`, without the bold, size or
// font of everything around them. Word gives typed text the formatting of the character
// before the caret, and at a paragraph's start of the character after it; struck words count,
// because they are still in the document.
//
// Each edit is its own transaction, as two people's edits are. Only the run properties are
// pinned new here: which side of the strike the words land on is pinned unchanged.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPart } from '../package/ooxml-tree.ts';
import { TreeDocumentStore } from '../store/tree-store.ts';
import type { TreeDocOp } from '../store/tree-op-validate.ts';
import { paragraphId, part, xml } from './tracked-edit-fixture.ts';

const ANA = { author: 'Ana', date: '2026-10-18T00:00:00Z' };
const BEN = { author: 'Ben', date: '2026-10-18T00:01:00Z' };
const FACE = '<w:rPr><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';
const SENTENCE = `<w:p><w:r>${FACE}<w:t xml:space="preserve">Alpha beta gamma</w:t></w:r></w:p>`;

/** Each op in its OWN transaction, in order, and the serialized result. */
function edited(before: OoxmlPart, ...ops: readonly TreeDocOp[]): string {
  const store = new TreeDocumentStore(before);
  for (const op of ops) {
    const result = store.transact((tx) => {
      tx.apply(op);
    });
    if (!result.ok) throw new Error(`op refused: ${result.reason} ${result.detail ?? ''}`);
  }
  return xml(store.part);
}

/** Ana strikes `[start, end)` of the sentence; Ben types "X " at `at`. */
function benTypesBesideAnasStrike(start: number, end: number, at: number): string {
  const before = part(SENTENCE);
  const id = paragraphId(before);
  return edited(
    before,
    { op: 'deleteText', paragraphId: id, start, end, revision: ANA },
    { op: 'insertText', paragraphId: id, offset: at, text: 'X ', revision: BEN }
  );
}

/** Ben's insertion, with the face it was given. */
const BENS_RUN = new RegExp(
  `<w:ins [^>]*w:author="Ben"[^>]*><w:r>${FACE}<w:t xml:space="preserve">X </w:t></w:r></w:ins>`
);

describe("typing beside another author's strike keeps the face", () => {
  test('inside the struck words: after the strike, in their face', () => {
    const out = benTypesBesideAnasStrike(6, 10, 8);
    expect(out).toMatch(BENS_RUN);
    // Placement is unchanged: struck words first, then the typing.
    expect(out).toMatch(/<w:delText>beta<\/w:delText><\/w:r><\/w:del><w:ins [^>]*w:author="Ben"/);
  });

  test('right after the struck words', () => {
    const out = benTypesBesideAnasStrike(6, 10, 10);
    expect(out).toMatch(BENS_RUN);
    expect(out).toMatch(/<w:delText>beta<\/w:delText><\/w:r><\/w:del><w:ins [^>]*w:author="Ben"/);
  });

  test('inside, and right after, a strike that ends the paragraph', () => {
    // No run follows the strike, so nothing at all lent the typed words a face.
    for (const at of [13, 16]) {
      const out = benTypesBesideAnasStrike(11, 16, at);
      expect(out).toMatch(BENS_RUN);
      expect(out).toMatch(/<\/w:del><w:ins [^>]*w:author="Ben"[^>]*>.*<\/w:ins><\/w:p>/);
    }
  });

  test("before a strike at the paragraph's start: the character after the caret", () => {
    const out = benTypesBesideAnasStrike(0, 5, 0);
    expect(out).toMatch(BENS_RUN);
    // Placement is unchanged: the typing stays in FRONT of the struck words.
    expect(out).toMatch(
      /<w:p><w:ins [^>]*w:author="Ben"[^>]*>.*<\/w:ins><w:del [^>]*w:author="Ana"/
    );
  });

  test('right after the strike, the struck words win over the run that follows', () => {
    // The character before the caret is the struck one. The toolbar reads its face there
    // in all-markup, and the untracked lane types it there: the typed words agree with both.
    const before = part(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r>' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>'
    );
    const id = paragraphId(before);
    const out = edited(
      before,
      { op: 'deleteText', paragraphId: id, start: 6, end: 10, revision: ANA },
      { op: 'insertText', paragraphId: id, offset: 10, text: 'X', revision: BEN }
    );
    expect(out).toMatch(
      /<\/w:del><w:ins [^>]*w:author="Ben"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>X<\/w:t><\/w:r><\/w:ins><w:r><w:rPr><w:i\/>/
    );
  });

  test("the face comes without the struck run's pending format record", () => {
    // `w:rPrChange` is somebody's undecided proposal about THOSE characters. The typed ones
    // were never in the state it describes, so they take the face and leave the record.
    const before = part(
      '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
        '<w:del w:id="5" w:author="Ana" w:date="2026-10-18T00:00:00Z"><w:r><w:rPr><w:b/>' +
        '<w:rPrChange w:id="6" w:author="Ana" w:date="2026-10-18T00:00:00Z"><w:rPr/></w:rPrChange>' +
        '</w:rPr><w:delText>beta</w:delText></w:r></w:del></w:p>'
    );
    const id = paragraphId(before);
    const out = edited(before, {
      op: 'insertText',
      paragraphId: id,
      offset: 10,
      text: 'X',
      revision: BEN,
    });
    expect(out).toMatch(
      /<\/w:del><w:ins [^>]*w:author="Ben"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>X<\/w:t>/
    );
  });

  test("a struck word inside a link does not lend the link's style past the link", () => {
    // The walk does not enter a link at its end, and typing there stays outside it; the
    // words take the run after the link, as before, not the link's character style.
    const before = part(
      '<w:p><w:hyperlink w:anchor="target"><w:r><w:t xml:space="preserve">see </w:t></w:r>' +
        '<w:del w:id="5" w:author="Ana" w:date="2026-10-18T00:00:00Z"><w:r><w:rPr>' +
        '<w:rStyle w:val="Hyperlink"/></w:rPr><w:delText>here</w:delText></w:r></w:del>' +
        `</w:hyperlink><w:r>${FACE}<w:t xml:space="preserve"> now</w:t></w:r></w:p>`
    );
    const id = paragraphId(before);
    const out = edited(before, {
      op: 'insertText',
      paragraphId: id,
      offset: 8,
      text: 'X ',
      revision: BEN,
    });
    expect(out).toMatch(/<\/w:hyperlink><w:ins [^>]*w:author="Ben"/);
    expect(out).toMatch(BENS_RUN);
  });
});
