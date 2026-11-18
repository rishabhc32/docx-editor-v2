// Where the review list places a section's change record (`w:sectPrChange`, a tracked page
// setup). Word keeps a section's properties with the paragraph mark that ends it, and the body's
// own section with the document's last paragraph mark (ECMA-376 Part 1, §17.6.17 and §17.6.18).
// The card sits there: before, the body's own section gave its card no range at all, so it
// sorted after every story and a reader that names a card's story by its range named none.

import { describe, expect, test } from 'bun:test';
import { applyTreeOp, collectReviewItems, readOoxmlPart, type OoxmlPart } from '../index.ts';
import type { ReviewRange } from '../store/review-reads.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RECORD =
  '<w:sectPrChange w:id="1" w:author="Cy" w:date="2026-10-30T09:00:00Z">' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/></w:sectPr></w:sectPrChange>';
const LANDSCAPE = `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>${RECORD}`;
const p = (text: string, properties = '') =>
  `<w:p>${properties}<w:r><w:t>${text}</w:t></w:r></w:p>`;

function load(body: string): OoxmlPart {
  const read = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}

/** The section record's card, and the paragraph ids of the body in order. */
function sectionCard(part: OoxmlPart) {
  const cards = collectReviewItems({ storyPart: part }).filter(
    (item) => item.kind === 'revision' && item.formattingKind === 'sectPrChange'
  );
  expect(cards).toHaveLength(1);
  const [card] = cards;
  if (card?.kind !== 'revision') throw new Error('no card');
  return card;
}

const at = (paragraphId: string, offset: number): ReviewRange => ({
  partName: '/word/document.xml',
  start: { paragraphId, offset },
  end: { paragraphId, offset },
});

/** The ids of the paragraphs that hold these texts, in order. */
function paragraphIds(part: OoxmlPart): string[] {
  const out: string[] = [];
  const visit = (node: OoxmlPart['root']): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') out.push(node.id);
    for (const child of node.children) visit(child as OoxmlPart['root']);
  };
  visit(part.root);
  return out;
}

describe("a section's change record in the review list", () => {
  test("the body's own section: a format card at the last paragraph mark", () => {
    const part = load(`${p('One.')}${p('Two.')}<w:sectPr>${LANDSCAPE}</w:sectPr>`);
    const card = sectionCard(part);
    expect(card.revisionKind).toBe('format');
    expect(card.author).toBe('Cy');
    expect(card.ranges).toEqual([at(paragraphIds(part)[1] ?? '', 4)]);
  });

  test('a section a paragraph ends: at that paragraph mark, as before', () => {
    const part = load(
      `${p('One.', `<w:pPr><w:sectPr>${LANDSCAPE}</w:sectPr></w:pPr>`)}${p('Two.')}` +
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    );
    expect(sectionCard(part).ranges).toEqual([at(paragraphIds(part)[0] ?? '', 4)]);
  });

  test('typing in the last paragraph moves the card with its mark', () => {
    const part = load(`${p('One.')}${p('Two.')}<w:sectPr>${LANDSCAPE}</w:sectPr>`);
    const last = paragraphIds(part)[1] ?? '';
    expect(sectionCard(part).ranges).toEqual([at(last, 4)]);
    const typed = applyTreeOp(part, { op: 'insertText', paragraphId: last, offset: 4, text: '!!' });
    if (!typed.ok) throw new Error(typed.reason);
    expect(sectionCard(typed.part).ranges).toEqual([at(last, 6)]);
  });

  test('the last paragraph in a block content control holds the mark', () => {
    const part = load(
      `${p('One.')}<w:sdt><w:sdtPr/><w:sdtContent>${p('Inside.')}</w:sdtContent></w:sdt>` +
        `<w:sectPr>${LANDSCAPE}</w:sectPr>`
    );
    expect(sectionCard(part).ranges).toEqual([at(paragraphIds(part)[1] ?? '', 7)]);
  });
});
