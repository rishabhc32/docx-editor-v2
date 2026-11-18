// A paste inside a link, an inline content control or another author's insertion, landed as
// Word lands it (Word 16.114 for Mac, paste evidence 2, `~$paste-out-e2-*.docx`). The source is
// Word's: bold "XX", then "z". `xx` is "XX" without its mark, `2p` is "XX¶z" without the last
// mark, `2pm` is both paragraphs with their marks. Word's Smart cut and paste spaces are not
// written here (design D18).

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../package/ooxml-tree.ts';
import { extractFragmentPackage } from '../store/clipboard-fragment-extract.ts';
import { revisionItemsOf } from '../store/review-reads.ts';
import type { TreePackageStore } from '../store/tree-package-store.ts';
import {
  bodyOf,
  buildPackage,
  fullBodyCoverage,
  openStore,
  paragraphIdsUnder,
  R,
} from './clipboard-fragment-fixtures.ts';

const LINK =
  `<w:p><w:r><w:t xml:space="preserve">Go </w:t></w:r><w:hyperlink xmlns:r="${R}" r:id="rId9">` +
  '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>LinkText</w:t></w:r></w:hyperlink></w:p>';
const CONTROL =
  '<w:p><w:sdt><w:sdtPr><w:alias w:val="Ctl"/><w:id w:val="424242"/></w:sdtPr><w:sdtContent>' +
  '<w:r><w:t>CtlText</w:t></w:r></w:sdtContent></w:sdt></w:p>';
const INSERTION =
  '<w:p><w:r><w:t xml:space="preserve">Base </w:t></w:r>' +
  '<w:ins w:id="1" w:author="Cara" w:date="2026-01-02T10:00:00Z"><w:r><w:t>CaraText</w:t></w:r>' +
  '</w:ins></w:p>';
const SOURCE =
  '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>XX</w:t></w:r></w:p><w:p><w:r><w:t>z</w:t></w:r></w:p>';
const ADA = { author: 'Ada', date: '2026-10-29T00:00:00Z' };

type Shape = 'xx' | '2p' | '2pm';

/** The editor's copy of the source in one of Word's three shapes. */
function copied(shape: Shape) {
  const source = buildPackage(SOURCE);
  const all = fullBodyCoverage(source);
  const ids = all.paragraphIds;
  const coverage =
    shape === '2pm'
      ? all
      : shape === '2p'
        ? { ...all, endOffset: 1, coveredParagraphIds: ids.slice(0, 1), lastMarkCovered: false }
        : {
            ...all,
            paragraphIds: ids.slice(0, 1),
            coveredParagraphIds: [],
            endOffset: 2,
            fullyCoveredBlockIds: [],
            lastMarkCovered: false,
          };
  const extracted = extractFragmentPackage(source, coverage);
  if (!extracted.ok) throw new Error(extracted.reason);
  return extracted;
}

function body(store: TreePackageStore): string {
  const pkg = store.currentPackage();
  const xml = serializeOoxmlPart(pkg.parts.get(pkg.mainDocumentPart)!);
  return xml
    .slice(xml.indexOf('<w:body>') + '<w:body>'.length, xml.indexOf('</w:body>'))
    .replaceAll(` xmlns:r="${R}"`, '')
    .replaceAll(/ w:date="[^"]*"/g, '')
    .replaceAll(/<w:sdtPr>.*?<\/w:sdtPr>/g, '');
}

/** Paste `shape` at `offset` of `host`'s paragraph: the body, or the refusal. */
function pasted(host: string, offset: number, shape: Shape, tracked: boolean) {
  const store = openStore(buildPackage(host));
  const paragraphId = paragraphIdsUnder(
    bodyOf(store.currentPackage().parts.get('/word/document.xml')!)
  )[0]!;
  const fragment = copied(shape);
  const result = store.applyFragmentPaste(
    { kind: 'body' },
    {
      paragraphId,
      offset,
      fragmentBytes: fragment.bytes,
      lastMarkCovered: fragment.lastMarkCovered,
      ...(tracked ? { revision: ADA } : {}),
    }
  );
  return { store, xml: result.ok ? body(store) : `refused: ${result.reason}` };
}

const B = (text: string) => `<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r>`;
const SPACE_B = '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>';
const SPACE = '<w:r><w:t xml:space="preserve"> </w:t></w:r>';
const LINKED = (text: string) =>
  `<w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${text}</w:t></w:r>`;
const INS_ADA = '<w:ins w:author="Ada" w:id="0">';
const CARA = (text: string) =>
  `<w:ins w:author="Cara" w:id="1"><w:r><w:t>${text}</w:t></w:r></w:ins>`;

describe('a paste inside a link', () => {
  test('without a mark: into the link, with its own formatting (`~$paste-out-e2-link-xx.docx`)', () => {
    const go = '<w:r><w:t xml:space="preserve">Go </w:t></w:r>';
    expect(pasted(LINK, 7, 'xx', false).xml).toBe(
      `<w:p>${go}<w:hyperlink r:id="rId9">${LINKED('Link')}${B('XX')}${LINKED('Text')}</w:hyperlink></w:p>`
    );
    const tracked = pasted(LINK, 7, 'xx', true);
    expect(tracked.xml).toBe(
      `<w:p>${go}<w:hyperlink r:id="rId9">${LINKED('Link')}${INS_ADA}${B('XX')}</w:ins>` +
        `${LINKED('Text')}</w:hyperlink></w:p>`
    );
    tracked.store.transact({ kind: 'body' }, (ctx) => {
      ctx.apply({ op: 'rejectAllRevisions' });
    });
    expect(
      body(tracked.store).replaceAll(
        '</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>',
        ''
      )
    ).toBe(`<w:p>${go}<w:hyperlink r:id="rId9">${LINKED('LinkText')}</w:hyperlink></w:p>`);
  });

  // Word turns the link into a HYPERLINK field across the pasted paragraphs
  // (`~$paste-out-e2-link-2p.docx`): not written yet, so refused by name
  test.each(['2p', '2pm'] as const)('%s: refused by name, not approximated', (shape) => {
    expect(pasted(LINK, 7, shape, false).xml).toBe('refused: fragment-paragraphs-in-link');
    expect(pasted(LINK, 7, shape, true).xml).toBe('refused: fragment-paragraphs-in-link');
  });
});

describe('a paste inside a rich text control', () => {
  const control = (inner: string) =>
    `<w:p><w:sdt><w:sdtContent><w:r><w:t>Ctl</w:t></w:r>${inner}<w:r><w:t>Text</w:t></w:r></w:sdtContent></w:sdt></w:p>`;

  test('without a mark: into the control (`~$paste-out-e2-sdt-xx.docx`)', () => {
    expect(pasted(CONTROL, 3, 'xx', false).xml).toBe(control(B('XX')));
    expect(pasted(CONTROL, 3, 'xx', true).xml).toBe(control(`${INS_ADA}${B('XX')}</w:ins>`));
  });

  test('paragraphs: each mark a space in the format before it, no paragraph made (`-2p`, `-2pm`)', () => {
    const z = '<w:r><w:t>z</w:t></w:r>';
    expect(pasted(CONTROL, 3, '2p', false).xml).toBe(control(`${B('XX')}${SPACE_B}${z}`));
    expect(pasted(CONTROL, 3, '2pm', false).xml).toBe(control(`${B('XX')}${SPACE_B}${z}${SPACE}`));
    expect(pasted(CONTROL, 3, '2p', true).xml).toBe(
      control(`${INS_ADA}${B('XX')}${SPACE_B}${z}</w:ins>`)
    );
  });
});

describe("a paste at an inline control's prompt", () => {
  const PROMPT =
    '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '<w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/>' +
    '</w:rPr><w:t>Click here to enter text.</w:t></w:r></w:sdtContent></w:sdt></w:p>';

  test('lines replace the prompt as one line inside the control, each mark a space', () => {
    const xml = pasted(PROMPT, 9, '2pm', false).xml;
    expect(xml).toContain(
      `<w:sdtContent>${B('XX')}${SPACE_B}<w:r><w:t>z</w:t></w:r>${SPACE}</w:sdtContent>`
    );
    expect(xml).not.toContain('showingPlcHdr');
    expect(xml.match(/<w:p>/g)).toHaveLength(1);
  });
});

describe("a paste inside another author's insertion: beside it, never nested", () => {
  const base = '<w:r><w:t xml:space="preserve">Base </w:t></w:r>';

  test('without a mark (`~$paste-out-e2-insoff-xx.docx`, `-inson-xx`)', () => {
    expect(pasted(INSERTION, 9, 'xx', false).xml).toBe(
      `<w:p>${base}${CARA('Cara')}${B('XX')}${CARA('Text')}</w:p>`
    );
    const tracked = pasted(INSERTION, 9, 'xx', true);
    expect(tracked.xml).toBe(
      `<w:p>${base}${CARA('Cara')}<w:ins w:author="Ada" w:id="2">${B('XX')}</w:ins>${CARA('Text')}</w:p>`
    );
    // Cara's halves are two items, as Word's split gives two insertions
    const part = tracked.store.currentPackage().parts.get('/word/document.xml')!;
    expect(
      revisionItemsOf(part)
        .map((item) => item.author)
        .sort()
    ).toEqual(['Ada', 'Cara', 'Cara']);
  });

  test('paragraphs: the insertion splits across them (`-insoff-2p`, `-inson-2p`)', () => {
    expect(pasted(INSERTION, 9, '2p', false).xml).toBe(
      `<w:p>${base}${CARA('Cara')}${B('XX')}</w:p><w:p><w:r><w:t>z</w:t></w:r>${CARA('Text')}</w:p>`
    );
    const mark = '<w:pPr><w:rPr><w:ins w:author="Ada" w:id="2"/></w:rPr></w:pPr>';
    expect(pasted(INSERTION, 9, '2p', true).xml).toBe(
      `<w:p>${mark}${base}${CARA('Cara')}<w:ins w:author="Ada" w:id="2">${B('XX')}</w:ins></w:p>` +
        `<w:p><w:ins w:author="Ada" w:id="2"><w:r><w:t>z</w:t></w:r></w:ins>${CARA('Text')}</w:p>`
    );
  });
});

describe('places a paste still cannot land inside', () => {
  test.each([
    [
      'a plain text control',
      '<w:p><w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent>' +
        '<w:r><w:t>Plain</w:t></w:r></w:sdtContent></w:sdt></w:p>',
      2,
    ],
    [
      'a deletion',
      '<w:p><w:del w:id="1" w:author="Cara"><w:r><w:delText>Struck</w:delText></w:r></w:del></w:p>',
      3,
    ],
    [
      'a link inside an insertion',
      `<w:p><w:ins w:id="1" w:author="Cara"><w:hyperlink xmlns:r="${R}" r:id="rId9">` +
        '<w:r><w:t>Linked</w:t></w:r></w:hyperlink></w:ins></w:p>',
      3,
    ],
  ])('%s: refused as before, so the editor falls back to plain text', (_name, host, offset) => {
    expect(pasted(host, offset, 'xx', false).xml).toBe('refused: indivisible-content');
  });
});
