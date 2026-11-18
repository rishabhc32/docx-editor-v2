// The review queue across commits it was not read between. A room gives the editor its own
// version of a key after the editor's commit and before the key's layout, and a paint deferred
// under input pressure skips a revision too, so the queue is read only after the second
// commit. When every commit since the last read rewrote the same paragraph, that paragraph is
// read again and the rest of the queue is kept; anything else derives the whole queue again.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { ReviewModuleContribution } from '../../contracts/modules.ts';
import { findNode } from '../../store/package/ooxml-edit.ts';
import type { OoxmlPart } from '../../store/package/ooxml-tree.ts';
import { collectReviewItems, revisionItemsOf } from '../../store/store/review-reads.ts';
import { openTreeSession, type TreeDocxSession } from '../tree-session.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_DOC =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT_NS}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL_NS}">` +
        `<Relationship Id="rId1" Type="${OFFICE_DOC}" Target="word/document.xml"/>` +
        '</Relationships>'
    ),
    'word/document.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const inserted = (id: number, text: string) =>
  `<w:ins w:id="${id}" w:author="Zed" w:date="2026-10-01T00:00:00Z">${run(text)}</w:ins>`;
const FILE = docx(
  Array.from(
    { length: 8 },
    (_, k) =>
      `<w:p>${run(`Paragraph ${k} text`)}${k % 3 === 0 ? inserted(100 + k, ' added') : ''}</w:p>`
  ).join('')
);
const ANA = { author: 'Ana', date: '2026-11-02T10:00:00Z' };

/** A review module that counts whole-queue derivations and paragraph reads. */
function counting(): { model: ReviewModuleContribution; full: number; paragraph: number } {
  const counts = {
    full: 0,
    paragraph: 0,
    model: {
      displayModes: ['all-markup'],
      collectReviewItems: (input) => {
        counts.full += 1;
        return collectReviewItems(input);
      },
      revisionItemsOfParagraph: (part: OoxmlPart, paragraphId: string) => {
        counts.paragraph += 1;
        const paragraph = findNode(part, paragraphId);
        return paragraph && paragraph.kind === 'paragraph'
          ? revisionItemsOf({ ...part, root: paragraph })
          : [];
      },
    } satisfies ReviewModuleContribution,
  };
  return counts;
}

function open(): { session: TreeDocxSession; counts: ReturnType<typeof counting> } {
  const counts = counting();
  const opened = openTreeSession(FILE, { reviewModel: counts.model });
  if (!opened.ok) throw new Error(opened.reason);
  return { session: opened.session, counts };
}

function paragraphIds(session: TreeDocxSession): string[] {
  return session
    .part()
    .root.children.find((child) => child.kind !== 'textValue' && child.localName === 'body')!
    .children.filter((child) => child.kind === 'paragraph')
    .map((child) => child.id);
}

function insert(session: TreeDocxSession, paragraphId: string, offset: number, text: string) {
  const result = session.applyTreeOps([
    { op: 'insertText', paragraphId, offset, text, revision: ANA },
  ]);
  expect(result.committed).toBe(true);
}

/** The queue the session gives, against one derived from scratch. */
function sameAsDerived(session: TreeDocxSession): void {
  const items = session.reviewItems();
  expect(items).toEqual(collectReviewItems({ storyPart: session.part() }));
  expect(items.some((item) => item.kind === 'revision' && item.author === 'Ana')).toBe(true);
}

describe('the review queue across commits it was not read between', () => {
  test('two commits to one paragraph read that paragraph again, not the whole story', () => {
    const { session, counts } = open();
    const ids = paragraphIds(session);
    session.reviewItems();
    expect(counts.full).toBe(1);
    insert(session, ids[4]!, 2, 'x');
    insert(session, ids[4]!, 3, 'y');
    sameAsDerived(session);
    expect(counts.full).toBe(1);
    expect(counts.paragraph).toBeGreaterThan(0);
  });

  test('a paragraph that holds a suggestion already is read again the same way', () => {
    const { session, counts } = open();
    const ids = paragraphIds(session);
    session.reviewItems();
    insert(session, ids[3]!, 0, 'a');
    insert(session, ids[3]!, 1, 'b');
    insert(session, ids[3]!, 2, 'c');
    sameAsDerived(session);
    expect(counts.full).toBe(1);
  });

  test('commits to two paragraphs derive the whole queue again', () => {
    const { session, counts } = open();
    const ids = paragraphIds(session);
    session.reviewItems();
    insert(session, ids[1]!, 2, 'x');
    insert(session, ids[5]!, 2, 'y');
    sameAsDerived(session);
    expect(counts.full).toBe(2);
  });

  test('more commits than the queue follows derive the whole queue again', () => {
    const { session, counts } = open();
    const ids = paragraphIds(session);
    session.reviewItems();
    for (let k = 0; k < 20; k += 1) insert(session, ids[4]!, 2, 'x');
    sameAsDerived(session);
    expect(counts.full).toBe(2);
  });
});
