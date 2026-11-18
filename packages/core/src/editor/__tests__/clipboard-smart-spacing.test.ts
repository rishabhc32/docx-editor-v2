// Word's Smart cut and paste (on by default): a paste between two words gets a space on each
// side (paste evidence 2, Word 16.114 for Mac, `~$paste-out-e2-*.docx`). Each space is a run of
// the paste: the one before in the pasted text's format, the one after in the format of the
// text after the place, both in a tracked paste's one insertion. Typing adds none.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { mountPaginatedSurface } from '../paginated-surface.ts';
import type { PaginatedSurface } from '../paginated-surface-contract.ts';
import { docx, mount, paragraph, putCaret } from './paginated-surface-fixtures.ts';

afterEach(() => {
  document.getSelection()?.removeAllRanges();
});

const run = (text: string, rPr = '') =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** The clipboard flavours of `[from, to)` in the first paragraph of a body. */
function copied(body: string, from: number, to: number): { text: string; html: string | null } {
  const source = mount(body);
  const paragraphId = source.surface.session.paragraphIds()[0]!;
  source.surface.setSelection({
    anchor: { paragraphId, offset: from },
    head: { paragraphId, offset: to },
  });
  return source.surface.copyFlavours();
}

/**
 * Bold "XX", copied without its paragraph mark: a selection to a paragraph's end takes the
 * mark too, as Word's smart paragraph selection does, so the line goes on after it.
 */
const BOLD_XX = () => copied(`<w:p>${run('XX', '<w:b/>')}${run(' more')}</w:p>`, 0, 2);

/** The runs of the first paragraph: their text, and whether each is bold. */
function runs(surface: PaginatedSurface): string[] {
  const xml = serializeOoxmlPart(surface.session.part());
  const first = /<w:p\b.*?<\/w:p>/s.exec(xml)?.[0] ?? '';
  return [...first.matchAll(/<w:r>(.*?)<\/w:r>/gs)].map((m) => {
    const text = /<w:t[^>]*>([^<]*)<\/w:t>/.exec(m[1]!)?.[1] ?? '';
    return m[1]!.includes('<w:b/>') ? `[${text}]` : text;
  });
}

const text = (surface: PaginatedSurface) => surface.session.bodyText();

describe('Smart cut and paste spacing', () => {
  test('a rich paste between two words gets a space on each side, as Word writes them', () => {
    const target = mount(`<w:p>${run('LinkText')}</w:p>`);
    putCaret(target.surface, 4);
    const flavours = BOLD_XX();
    expect(target.surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    expect(text(target.surface)).toBe('Link XX Text');
    // The space before takes the paste's format, the space after the following text's
    expect(runs(target.surface)).toEqual(['Link', '[ ]', '[XX]', ' ', 'Text']);
    target.surface.undo();
    expect(text(target.surface)).toBe('LinkText');
  });

  test('a plain paste gets the spaces too; typing gets none', () => {
    const target = mount(`<w:p>${run('LinkText')}</w:p>`);
    putCaret(target.surface, 4);
    expect(target.surface.pasteRich('XX', null)).toBe(true);
    expect(text(target.surface)).toBe('Link XX Text');
    const typed = mount(`<w:p>${run('LinkText')}</w:p>`);
    putCaret(typed.surface, 4);
    typed.surface.insertPlainText('XX');
    expect(text(typed.surface)).toBe('LinkXXText');
  });

  test('a space only where a word meets a word', () => {
    const cases: [string, number, string, string][] = [
      // Before a word, after a space: a space after only
      ['The fee is', 'The fee '.length, 'XX', 'The fee XX is'],
      // After a word, before a space: a space before only
      ['The fee is', 'The fee'.length, 'XX', 'The fee XX is'],
      // Before punctuation: none after
      ['The fee.', 'The fee'.length, 'XX', 'The fee XX.'],
      // At a paragraph's start: none before
      ['Text', 0, 'XX', 'XX Text'],
      // Pasted text with its own spaces: none added
      ['LinkText', 4, ' XX ', 'Link XX Text'],
      // Punctuation pasted: none
      ['LinkText', 4, '(XX)', 'Link(XX)Text'],
      // Scripts written without spaces: none
      ['日本語', 2, '中文', '日本中文語'],
    ];
    for (const [line, offset, pasted, expected] of cases) {
      const target = mount(`<w:p>${run(line)}</w:p>`);
      putCaret(target.surface, offset);
      expect(target.surface.pasteRich(pasted, null)).toBe(true);
      expect({ line, offset, pasted, result: text(target.surface) }).toEqual({
        line,
        offset,
        pasted,
        result: expected,
      });
    }
  });

  // Word 16.114 for Mac, W4 evidence (`~$paste-out-w4-<case>.docx`): "fee" or a Thai word
  // pasted beside a sentence end, a comma, Chinese and Japanese, and inside Thai
  test('beside a sentence end, punctuation, CJK and Thai, as Word spaces it', () => {
    const cases: [string, string, number, string, string][] = [
      ['a', 'Stop.Go', 'Stop.'.length, 'fee', 'Stop. fee Go'],
      ['b', 'Go, then', 'Go'.length, 'fee', 'Go fee, then'],
      ['c', '日本語', 2, 'fee', '日本fee語'],
      ['d', 'สวัสดีครับ', 'สวัสดี'.length, 'fee', 'สวัสดี fee ครับ'],
      ['e', 'paynow', 'pay'.length, '料金', 'pay料金now'],
      ['f', '日本語', 2, '料金', '日本料金語'],
      ['g', 'สวัสดีครับ', 'สวัสดี'.length, 'ราคา', 'สวัสดี ราคาครับ'],
    ];
    for (const [sample, line, offset, pasted, expected] of cases) {
      const source = `<w:p>${run(pasted, '<w:b/>')}${run(' more')}</w:p>`;
      const flavours = copied(source, 0, pasted.length);
      for (const html of [null, flavours.html]) {
        const target = mount(`<w:p>${run(line)}</w:p>`);
        putCaret(target.surface, offset);
        expect(target.surface.pasteRich(pasted, html)).toBe(true);
        expect({ sample, rich: html !== null, result: text(target.surface) }).toEqual({
          sample,
          rich: html !== null,
          result: expected,
        });
      }
    }
  });

  test('after a sentence end, the space before takes the paste format, as Word writes it', () => {
    const target = mount(`<w:p>${run('Stop.Go')}</w:p>`);
    putCaret(target.surface, 'Stop.'.length);
    const flavours = BOLD_XX();
    expect(target.surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    expect(runs(target.surface)).toEqual(['Stop.', '[ ]', '[XX]', ' ', 'Go']);
  });

  test('pasted paragraphs: a space before the first, and after the last without its mark', () => {
    const source = `<w:p>${run('Line A', '<w:b/>')}</w:p><w:p>${run('Line B more')}</w:p>`;
    const open = mount(source);
    const ids = open.surface.session.paragraphIds();
    open.surface.setSelection({
      anchor: { paragraphId: ids[0]!, offset: 0 },
      head: { paragraphId: ids[1]!, offset: 'Line B'.length },
    });
    const twoLines = open.surface.copyFlavours();
    const target = mount(`<w:p>${run('LinkText')}</w:p>`);
    putCaret(target.surface, 4);
    expect(target.surface.pasteRich(twoLines.text, twoLines.html)).toBe(true);
    expect(text(target.surface)).toBe('Link Line A\nLine B Text');
  });

  test('pasted paragraphs with their last mark: a space before the first only', () => {
    const source = `<w:p>${run('Line A', '<w:b/>')}</w:p><w:p>${run('Line B')}</w:p>`;
    const open = mount(source);
    open.surface.selectAll();
    const covered = open.surface.copyFlavours();
    const target = mount(`<w:p>${run('LinkText')}</w:p>`);
    putCaret(target.surface, 4);
    expect(target.surface.pasteRich(covered.text, covered.html)).toBe(true);
    expect(text(target.surface)).toBe('Link Line A\nLine B\nText');
  });

  test('inside a link, the space after takes the link text format', () => {
    const style = '<w:rStyle w:val="Hyperlink"/>';
    const target = mount(
      `<w:p><w:hyperlink w:anchor="here">${run('LinkText', style)}</w:hyperlink></w:p>`
    );
    putCaret(target.surface, 4);
    const flavours = BOLD_XX();
    expect(target.surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    expect(text(target.surface)).toBe('Link XX Text');
    const xml = serializeOoxmlPart(target.surface.session.part());
    const link = /<w:hyperlink\b[^>]*>(.*?)<\/w:hyperlink>/s.exec(xml)?.[1] ?? '';
    // Every run stays in the link; the space after the paste has the link's style
    expect(link).toContain('<w:b/></w:rPr><w:t>XX</w:t>');
    expect(link).toMatch(
      /<w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t xml:space="preserve"> <\/w:t>/
    );
  });

  test('a tracked paste holds both spaces in its one insertion', () => {
    const container = document.createElement('div');
    const mounted = mountPaginatedSurface(container, docx(`<w:p>${run('LinkText')}</w:p>`), {
      scale: 1,
      author: 'Reviewer',
      editingMode: 'suggest',
    });
    if (!mounted.ok) throw new Error(mounted.reason);
    putCaret(mounted.surface, 4);
    const flavours = BOLD_XX();
    expect(mounted.surface.pasteRich(flavours.text, flavours.html)).toBe(true);
    const xml = serializeOoxmlPart(mounted.surface.session.part());
    const insertions = [...xml.matchAll(/<w:ins\b[^>]*>(.*?)<\/w:ins>/gs)].map((m) => m[1]!);
    expect(insertions).toHaveLength(1);
    expect(insertions[0]!.match(/<w:t[^>]*>([^<]*)<\/w:t>/g)?.length).toBe(3);
    expect(text(mounted.surface)).toBe('Link XX Text');
  });

  test('no space at an empty paragraph', () => {
    const target = mount(paragraph(''));
    putCaret(target.surface, 0);
    expect(target.surface.pasteRich('XX', null)).toBe(true);
    expect(text(target.surface)).toBe('XX');
  });
});
