// A file's own `w:numberingChange` (a first-edition record, absent from the current schema).
// Word for Mac counts it as one revision, and Accept All and Reject All both remove it and keep
// the list (2026-10-29, numchange.docx). The engine does the same.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, revisionItemsOf, serializeOoxmlPart, type OoxmlPart } from '../index.ts';
import { applyTreeOp } from '../store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RECORD =
  '<w:numberingChange w:id="7" w:author="Cara" w:date="2026-01-01T00:00:00Z" w:original="%1)"/>';
const item = (numPr: string, text: string) =>
  `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>${numPr}</w:numPr></w:pPr>` +
  `<w:r><w:t>${text}</w:t></w:r></w:p>`;

function load(body: string): OoxmlPart {
  const read = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'application/xml',
  });
  if (!read.ok) throw new Error(read.reason);
  return read.part;
}

const NUMBERED = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';

describe("a file's own w:numberingChange", () => {
  const part = load(item(RECORD, 'One.') + item('', 'Two.'));

  test('is one revision', () => {
    const items = revisionItemsOf(part);
    expect(items.map((i) => [i.address.id, i.revisionKind])).toEqual([['7', 'structural']]);
  });

  for (const op of ['acceptAllRevisions', 'rejectAllRevisions'] as const)
    test(`${op} removes it and keeps the list`, () => {
      const result = applyTreeOp(part, { op });
      expect(result.ok).toBe(true);
      if (!result.ok) return;
      const xml = serializeOoxmlPart(result.part);
      expect(xml).not.toContain('numberingChange');
      expect(xml.split(NUMBERED)).toHaveLength(3);
    });

  test('one outside a numbering reference stays file content', () => {
    const odd = load(`<w:p><w:pPr>${RECORD}</w:pPr><w:r><w:t>Odd.</w:t></w:r></w:p>`);
    expect(revisionItemsOf(odd)).toEqual([]);
  });
});
