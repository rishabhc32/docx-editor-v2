// A revision decision that empties a content control which names no prompt.
//
// Word 16.114 (2026-10-30, samples `~$i6w-*` in compatibility mode and `~$i6n-*` in native
// mode): Accept All of a deletion of all the control's text, and Reject All of an insertion of
// all of it, give the same control in all 16 cases. Word writes `w:showingPlcHdr` and an empty
// `w:sdtEndPr`, drops `w:richText` (a plain text control keeps `w:text`), and fills the content
// with ONE run of five spaces in no style, inline and block. It adds no glossary block.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, serializeOoxmlPart, type OoxmlPart } from '../index.ts';
import { applyTreeOp } from '../store/tree-op-apply.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REV = 'w:id="91" w:author="Ana" w:date="2026-10-28T09:00:00Z"';
const BLANK = '<w:r><w:t xml:space="preserve">     </w:t></w:r>';

function parse(body: string): OoxmlPart {
  const result = readOoxmlPart(
    `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    {
      name: '/word/document.xml',
      contentType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    }
  );
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function controlXml(part: OoxmlPart): string {
  const found = /<w:sdt>.*<\/w:sdt>/s.exec(serializeOoxmlPart(part))?.[0];
  if (found === undefined) throw new Error('no control XML');
  return found;
}

const deleted = `<w:del ${REV}><w:r><w:delText>Acme</w:delText></w:r></w:del>`;
const inserted = `<w:ins ${REV}><w:r><w:t>Acme</w:t></w:r></w:ins>`;

const inline = (type: string, content: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="601"/>${type}` +
  `</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>` +
  `<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>`;
const block = (type: string, content: string) =>
  `<w:p><w:r><w:t>Before.</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="602"/>${type}</w:sdtPr>` +
  `<w:sdtContent><w:p>${content}</w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;

/** Word's control after the decision: sdtPr as Word writes it, then the blank prompt. */
const wordInline = (id: string, type: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/>${type}</w:sdtPr><w:sdtEndPr/>` +
  `<w:sdtContent>${BLANK}</w:sdtContent></w:sdt>`;
const wordBlock = (id: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="${id}"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/>` +
  `<w:sdtContent><w:p>${BLANK}</w:p></w:sdtContent></w:sdt>`;

describe("a decision that empties a control with no prompt gives Word's blank prompt", () => {
  const cases = [
    ['an inline plain text control', inline, '<w:text/>', wordInline('601', '<w:text/>')],
    ['an inline untyped control', inline, '', wordInline('601', '')],
    ['an inline rich text control', inline, '<w:richText/>', wordInline('601', '')],
    ['a block untyped control', block, '', wordBlock('602')],
    ['a block rich text control', block, '<w:richText/>', wordBlock('602')],
  ] as const;
  for (const [name, shape, type, expected] of cases) {
    test(`${name}: Accept All of a deletion`, () => {
      const decided = applyTreeOp(parse(shape(type, deleted)), { op: 'acceptAllRevisions' });
      if (!decided.ok) throw new Error(decided.reason);
      expect(controlXml(decided.part)).toBe(expected);
    });
    test(`${name}: Reject All of an insertion`, () => {
      const decided = applyTreeOp(parse(shape(type, inserted)), { op: 'rejectAllRevisions' });
      if (!decided.ok) throw new Error(decided.reason);
      expect(controlXml(decided.part)).toBe(expected);
    });
  }

  test('one decision on the only revision gives the same form as Accept All', () => {
    const decided = applyTreeOp(parse(inline('<w:text/>', deleted)), {
      op: 'acceptRevision',
      revision: { id: '91', author: 'Ana', date: '2026-10-28T09:00:00Z' },
    });
    if (!decided.ok) throw new Error(decided.reason);
    expect(controlXml(decided.part)).toBe(wordInline('601', '<w:text/>'));
  });

  test('a control that keeps its own end properties keeps them', () => {
    const body = inline('<w:text/>', inserted).replace(
      '</w:sdtPr>',
      '</w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr>'
    );
    const decided = applyTreeOp(parse(body), { op: 'rejectAllRevisions' });
    if (!decided.ok) throw new Error(decided.reason);
    expect(controlXml(decided.part)).toContain(
      '<w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent>'
    );
    expect(controlXml(decided.part)).toContain(BLANK);
  });

  test('a control that names a prompt shows that prompt, not the blank one', () => {
    const body = inline(
      '<w:placeholder><w:docPart w:val="MyPrompt"/></w:placeholder><w:text/>',
      inserted
    );
    const glossary = parse('<w:p><w:r><w:t>Type the name.</w:t></w:r></w:p>').root;
    const prompt = glossary.children.flatMap((body) =>
      body.kind === 'textValue' ? [] : body.children.filter((node) => node.kind === 'paragraph')
    );
    const decided = applyTreeOp(
      parse(body),
      { op: 'rejectAllRevisions' },
      { placeholderBlocks: () => prompt }
    );
    if (!decided.ok) throw new Error(decided.reason);
    expect(controlXml(decided.part)).toContain('<w:showingPlcHdr/>');
    expect(controlXml(decided.part)).toContain('Type the name.');
    expect(controlXml(decided.part)).not.toContain('sdtEndPr');
  });
});
