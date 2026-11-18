// Suggesting mode records a different first page and the header and footer distances as a
// `w:sectPrChange`, as Word does with track changes on (Word for Mac, 2026-10-30: "Formatted:
// Different first page header", "Formatted: Header distance from edge"). Before, the editor
// wrote them untracked: no card, and nothing to reject. Odd and even pages is a setting, which
// has no tracked form: it is still written untracked.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const FILE = zipSync({
  '[Content_Types].xml': strToU8(
    `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
  ),
  '_rels/.rels': strToU8(
    `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
  ),
  'word/document.xml': strToU8(
    `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
      '</w:sectPr></w:body></w:document>'
  ),
});

const mounted: { editor: DocxEditorInstance; container: HTMLElement }[] = [];

afterEach(() => {
  for (const { editor, container } of mounted.splice(0)) {
    editor.destroy();
    container.remove();
  }
});

function open(mode: 'suggesting' | 'edit'): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Bea',
    mode,
    modules: [stubReviewModule()],
  });
  mounted.push({ editor, container });
  return editor;
}

async function partsOf(editor: DocxEditorInstance): Promise<Record<string, string>> {
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(unzipSync(new Uint8Array(await editor.save()))))
    out[name] = strFromU8(bytes);
  return out;
}

/** What the section's `w:sectPrChange` by Bea records, or undefined. */
const recordOf = (xml: string): string | undefined =>
  /<w:sectPrChange [^>]*w:author="Bea"[^>]*><w:sectPr>([\s\S]*?)<\/w:sectPr>/.exec(xml)?.[1];

describe('suggesting a different first page and the header and footer distances', () => {
  test("a different first page is Word's w:sectPrChange, and undo takes it back", async () => {
    const editor = open('suggesting');
    expect(
      editor.exec({ type: 'setHeaderFooterOptions', sectionIndex: 0, titlePage: true })
    ).toEqual({ ok: true, changed: true });
    const xml = (await partsOf(editor))['word/document.xml'] ?? '';
    expect(xml).toContain('<w:titlePg/>');
    expect(recordOf(xml)).toContain('<w:titlePg w:val="0"/>');
    editor.surface?.undo();
    const after = (await partsOf(editor))['word/document.xml'] ?? '';
    expect(after).not.toContain('titlePg');
    expect(after).not.toContain('sectPrChange');
  });

  test('a header or footer distance records the earlier margins', async () => {
    for (const [options, now] of [
      [{ headerDistanceTwips: 1080 }, 'w:header="1080"'],
      [{ footerDistanceTwips: 540 }, 'w:footer="540"'],
    ] as const) {
      const editor = open('suggesting');
      expect(editor.exec({ type: 'setHeaderFooterOptions', sectionIndex: 0, ...options }).ok).toBe(
        true
      );
      const xml = (await partsOf(editor))['word/document.xml'] ?? '';
      expect(xml.split('<w:sectPrChange')[0]).toContain(now);
      expect(recordOf(xml)).toContain('w:header="720" w:left="1440" w:right="1440" w:top="1440"');
      expect(recordOf(xml)).not.toContain(now);
    }
  });

  test('odd and even pages, a setting, is written untracked; editing mode writes no record', async () => {
    const suggesting = open('suggesting');
    expect(suggesting.exec({ type: 'setHeaderFooterOptions', evenAndOddHeaders: true }).ok).toBe(
      true
    );
    const parts = await partsOf(suggesting);
    expect(parts['word/settings.xml']).toContain('evenAndOddHeaders');
    expect(parts['word/document.xml']).not.toContain('sectPrChange');
    const editing = open('edit');
    expect(
      editing.exec({ type: 'setHeaderFooterOptions', sectionIndex: 0, titlePage: true }).ok
    ).toBe(true);
    const xml = (await partsOf(editing))['word/document.xml'] ?? '';
    expect(xml).toContain('<w:titlePg/>');
    expect(xml).not.toContain('sectPrChange');
  });
});
