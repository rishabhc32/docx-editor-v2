/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `ReviewItemQuery.wholeParagraphFormat`: a paragraph formatting change (`w:pPrChange`) as its
// whole paragraph in `getReviewItemRects`, `getReviewItemsAt`, and `isActive`, so a host that
// draws its own review chrome does not read painted DOM for it. Without the option the change
// keeps its collapsed range at the paragraph end.
//
// Expected text geometry comes from text-highlight marks, as in `review-items-at.test.ts`;
// line and marker geometry comes from the published layout records.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import type {
  HighlightRect,
  ReviewItemPlacement,
  ReviewItemQuery,
} from '@docx-editor.dev/core/contracts/editor';
import {
  paragraphFragmentsOf,
  type PageRecord,
  type ParagraphFragmentRecord,
} from '@docx-editor.dev/core/layout';
import { reviewModule } from '../review/review-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

interface DocxParts {
  readonly body: string;
  /** Default-header content; wires the part, its relationship and the section reference. */
  readonly header?: string;
  /** `w:footnote` elements; wires footnotes.xml, its relationship and content type. */
  readonly footnotes?: string;
  /** A `w:numbering` part. */
  readonly numbering?: string;
  /** Page size and margins for the section. */
  readonly page?: string;
}

function docx({ body, header, footnotes, numbering, page = '' }: DocxParts): Uint8Array {
  /** File name, relationship and content type kind, XML. */
  const parts: (readonly [string, string, string])[] = [];
  if (header) parts.push(['header1.xml', 'header', `<w:hdr xmlns:w="${W}">${header}</w:hdr>`]);
  if (footnotes) {
    const xml = `<w:footnotes xmlns:w="${W}">${footnotes}</w:footnotes>`;
    parts.push(['footnotes.xml', 'footnotes', xml]);
  }
  if (numbering) parts.push(['numbering.xml', 'numbering', numbering]);
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
        parts
          .map(
            ([name, kind]) =>
              `<Override PartName="/word/${name}" ContentType="${WML}.${kind}+xml"/>`
          )
          .join('') +
        `</Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr>` +
        (header ? `<w:headerReference w:type="default" r:id="rId-header"/>` : '') +
        `${page}</w:sectPr></w:body></w:document>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        parts
          .map(
            ([name, kind]) =>
              `<Relationship Id="rId-${kind}" Type="${R}/${kind}" Target="${name}"/>`
          )
          .join('') +
        `</Relationships>`
    ),
  };
  for (const [name, , xml] of parts) files[`word/${name}`] = strToU8(xml);
  return zipSync(files);
}

const AUTHOR = 'w:author="Grace Hopper" w:date="2026-01-02T03:04:05Z"';
/** A suggestion to center the paragraph: `w:jc` now, no alignment before. */
const CENTERED = `<w:jc w:val="center"/><w:pPrChange w:id="5" ${AUTHOR}><w:pPr/></w:pPrChange>`;
const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
/** A centered heading suggestion; `content` is its runs. */
const heading = (content: string) => `<w:p><w:pPr>${CENTERED}</w:pPr>${content}</w:p>`;
const INSERTED_WORD = `<w:ins w:id="6" ${AUTHOR}>${run('schedule')}</w:ins>`;
const PLAIN = `<w:p>${run('Body text after the heading.')}</w:p>`;

const ORIGIN = { left: 40, top: 30 };
const WHOLE: ReviewItemQuery = { wholeParagraphFormat: true };
const cleanups: (() => void)[] = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0).reverse()) cleanup();
});

function mount(parts: DocxParts): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(parts),
    modules: [reviewModule()],
  });
  cleanups.push(() => {
    editor.destroy();
    container.remove();
  });
  const layer = container.querySelector<HTMLElement>('.docx-text-highlight-overlay')!;
  layer.getBoundingClientRect = () =>
    ({ ...ORIGIN, x: ORIGIN.left, y: ORIGIN.top, right: 0, bottom: 0 }) as DOMRect;
  return editor;
}

/** The paragraph formatting change, with its paragraph and that paragraph's length. */
function paragraphChange(
  editor: DocxEditorInstance,
  query?: ReviewItemQuery
): { placement: ReviewItemPlacement; paragraphId: string; length: number } {
  const placement = editor
    .getReviewItems(query)
    .find((item) => item.kind === 'revision' && item.item.formattingKind === 'pPrChange');
  if (!placement || placement.kind !== 'revision') throw new Error('no paragraph change');
  const range = placement.item.ranges[0]!;
  return { placement, paragraphId: range.start.paragraphId, length: range.start.offset };
}

/** Client rectangles of a probe highlight over one model range. */
function probe(
  editor: DocxEditorInstance,
  blockId: string,
  start: number,
  length: number
): HighlightRect[] {
  expect(editor.setHighlights('probe', [{ blockId, start, length }]).applied).toBe(1);
  const marks = [
    ...document.querySelectorAll<HTMLElement>('[data-highlight-set="probe"] .docx-text-highlight'),
  ].map((mark) => {
    const left = ORIGIN.left + Number.parseFloat(mark.style.left);
    const top = ORIGIN.top + Number.parseFloat(mark.style.top);
    const width = Number.parseFloat(mark.style.width);
    const height = Number.parseFloat(mark.style.height);
    return { x: left, y: top, left, top, width, height, right: left + width, bottom: top + height };
  });
  editor.clearHighlights('probe');
  expect(marks.length).toBeGreaterThan(0);
  return marks;
}

/** Body and table fragments of one paragraph, with their pages. */
function fragmentsOf(
  editor: DocxEditorInstance,
  paragraphId: string
): { page: PageRecord; fragment: ParagraphFragmentRecord }[] {
  return editor.surface!.publishedLayout().pages.flatMap((page) =>
    paragraphFragmentsOf(page)
      .filter((fragment) => fragment.paragraphId === paragraphId)
      .map((fragment) => ({ page, fragment }))
  );
}

/** A page-content x in client coordinates. */
function clientX(editor: DocxEditorInstance, page: PageRecord, x: number): number {
  return ORIGIN.left + (page.contentBox.x + x) * editor.getRenderScale();
}

const keys = (hits: readonly { readonly placement: ReviewItemPlacement }[]) =>
  hits.map((hit) => hit.placement.key);

function caretAt(editor: DocxEditorInstance, paragraphId: string, offset: number): void {
  editor.surface!.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

const isActive = (editor: DocxEditorInstance, key: string, query?: ReviewItemQuery): boolean =>
  editor.getReviewItems(query).find((item) => item.key === key)!.isActive;

/** Open the change's story, then put the caret at the paragraph start. */
function expectActiveInItsStory(
  editor: DocxEditorInstance,
  key: string,
  paragraphId: string
): void {
  expect(editor.setActiveReviewItem(key).ok).toBe(true);
  caretAt(editor, paragraphId, 0);
  expect(editor.surface!.state().selection.head).toEqual({ paragraphId, offset: 0 });
  expect(isActive(editor, key, WHOLE)).toBe(true);
}

/** Each text band has a band on its line that covers it, at the text's height. */
function expectBandsOver(bands: readonly HighlightRect[], text: readonly HighlightRect[]): void {
  for (const mark of text) {
    const band = bands.find((candidate) => Math.abs(candidate.top - mark.top) < 0.01);
    expect(band).toBeDefined();
    expect(band!.left).toBeLessThanOrEqual(mark.left + 0.01);
    expect(band!.right).toBeGreaterThanOrEqual(mark.right - 0.01);
    expect(band!.height).toBeCloseTo(mark.height, 3);
  }
}

describe('a centered heading suggestion', () => {
  const parts = { body: heading(run('Payment terms')) + PLAIN };

  test('rects are one band per line, across the text column', () => {
    const editor = mount(parts);
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    const [{ page, fragment }] = fragmentsOf(editor, paragraphId);
    expect(rects).toHaveLength(fragment!.lines.length);
    expectBandsOver(rects, probe(editor, paragraphId, 0, length));
    const line = fragment!.lines[0]!;
    expect(rects[0]!.left).toBeCloseTo(clientX(editor, page!, line.box.x), 3);
    expect(rects[0]!.width).toBeCloseTo(line.box.width * editor.getRenderScale(), 3);
  });

  test('a hit on the heading text, or beside it on the line, returns the change', () => {
    const editor = mount(parts);
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const [text] = probe(editor, paragraphId, 0, length);
    const [band] = editor.getReviewItemRects(placement.key, WHOLE);
    const y = text!.top + text!.height / 2;
    const onText = editor.getReviewItemsAt(text!.left + text!.width / 2, y, WHOLE);
    expect(keys(onText)).toEqual([placement.key]);
    expect(onText[0]!.rect).toEqual(band!);
    // Centered, so the line's start is empty column.
    expect(keys(editor.getReviewItemsAt(band!.left + 1, y, WHOLE))).toEqual([placement.key]);
  });

  test('the caret at the start, middle, and end marks it active', () => {
    const editor = mount(parts);
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    for (const offset of [0, Math.floor(length / 2), length]) {
      caretAt(editor, paragraphId, offset);
      expect(isActive(editor, placement.key, WHOLE)).toBe(true);
    }
  });

  test('the caret in the next paragraph does not mark it active', () => {
    const editor = mount(parts);
    const { placement, paragraphId } = paragraphChange(editor, WHOLE);
    caretAt(editor, paragraphId, 0);
    expect(isActive(editor, placement.key, WHOLE)).toBe(true);
    const next = editor.surface!.publishedLayout().pages[0]!.fragments[1]!;
    if (next.kind !== 'paragraph') throw new Error('expected a paragraph');
    caretAt(editor, next.paragraphId, 0);
    expect(isActive(editor, placement.key, WHOLE)).toBe(false);
  });

  test('without the option the change keeps its point at the paragraph end', () => {
    const editor = mount(parts);
    const { placement, paragraphId, length } = paragraphChange(editor);
    const [text] = probe(editor, paragraphId, 0, length);
    const y = text!.top + text!.height / 2;
    for (const query of [undefined, { wholeParagraphFormat: false }]) {
      expect(editor.getReviewItemRects(placement.key, query)).toEqual([]);
      expect(editor.getReviewItemsAt(text!.left + text!.width / 2, y, query)).toEqual([]);
      caretAt(editor, paragraphId, 0);
      expect(isActive(editor, placement.key, query)).toBe(false);
      caretAt(editor, paragraphId, length);
      expect(isActive(editor, placement.key, query)).toBe(true);
    }
  });

  test('an activation exclusion keeps the change inactive', () => {
    const editor = mount(parts);
    const { placement, paragraphId } = paragraphChange(editor, WHOLE);
    caretAt(editor, paragraphId, 0);
    expect(isActive(editor, placement.key, WHOLE)).toBe(true);
    editor.setReviewActivationExclusions(['format']);
    expect(isActive(editor, placement.key, WHOLE)).toBe(false);
  });
});

describe('a word change inside the heading', () => {
  const parts = { body: heading(run('Payment ') + INSERTED_WORD) + PLAIN };

  test('the word change is returned first, then the paragraph change', () => {
    const editor = mount(parts);
    const { placement, paragraphId } = paragraphChange(editor, WHOLE);
    const insertion = editor
      .getReviewItems(WHOLE)
      .find((item) => item.kind === 'revision' && item.item.revisionKind === 'insert')!;
    const [word] = probe(editor, paragraphId, 8, 8);
    const hits = editor.getReviewItemsAt(
      word!.left + word!.width / 2,
      word!.top + word!.height / 2,
      WHOLE
    );
    expect(keys(hits)).toEqual([insertion.key, placement.key]);
    // Off the word, only the paragraph change.
    const [plain] = probe(editor, paragraphId, 0, 7);
    expect(
      keys(editor.getReviewItemsAt(plain!.left + 1, plain!.top + plain!.height / 2, WHOLE))
    ).toEqual([placement.key]);
  });

  test('the caret in the word marks both active', () => {
    const editor = mount(parts);
    const { placement, paragraphId } = paragraphChange(editor, WHOLE);
    caretAt(editor, paragraphId, 10);
    const active = editor
      .getReviewItems(WHOLE)
      .filter((item) => item.isActive)
      .map((item) => item.kind === 'revision' && item.item.revisionKind);
    expect(active).toEqual(['insert', 'format']);
    expect(isActive(editor, placement.key)).toBe(false);
  });
});

describe('every story the review queue lists', () => {
  test('a paragraph split across two pages bands on both pages', () => {
    const words = Array.from({ length: 120 }, (_, at) => `word${at}`).join(' ');
    const editor = mount({
      body: heading(run(words)),
      page:
        '<w:pgSz w:w="12240" w:h="3600"/>' +
        '<w:pgMar w:top="720" w:right="1440" w:bottom="720" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/>',
    });
    const { placement, paragraphId } = paragraphChange(editor, WHOLE);
    const fragments = fragmentsOf(editor, paragraphId);
    expect(new Set(fragments.map(({ page }) => page.index)).size).toBeGreaterThanOrEqual(2);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    const lines = fragments.reduce((sum, { fragment }) => sum + fragment.lines.length, 0);
    expect(rects).toHaveLength(lines);
    const scale = editor.getRenderScale();
    const onPage = (page: PageRecord) =>
      rects.filter(
        (rect) =>
          rect.top >= ORIGIN.top + page.box.y * scale &&
          rect.bottom <= ORIGIN.top + (page.box.y + page.box.height) * scale
      ).length;
    for (const { page, fragment } of fragments) expect(onPage(page)).toBe(fragment.lines.length);
  });

  test('a table cell', () => {
    const cell = (content: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
    const editor = mount({
      body:
        `<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
        `<w:tr>${cell(heading(run('Fees')))}${cell(`<w:p>${run('Other')}</w:p>`)}</w:tr></w:tbl>` +
        PLAIN,
    });
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    expect(rects).toHaveLength(1);
    const [text] = probe(editor, paragraphId, 0, length);
    expectBandsOver(rects, [text!]);
    const hits = editor.getReviewItemsAt(rects[0]!.left + 1, text!.top + 1, WHOLE);
    expect(keys(hits)).toEqual([placement.key]);
    caretAt(editor, paragraphId, 0);
    expect(isActive(editor, placement.key, WHOLE)).toBe(true);
  });

  test('a header', () => {
    const editor = mount({ body: PLAIN, header: heading(run('Letterhead')) });
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    const text = probe(editor, paragraphId, 0, length);
    // The header repeats on every painted page, so each page bands it.
    expect(rects).toHaveLength(text.length);
    expectBandsOver(rects, text);
    const hits = editor.getReviewItemsAt(rects[0]!.left + 1, text[0]!.top + 1, WHOLE);
    expect(keys(hits)).toEqual([placement.key]);
    expectActiveInItsStory(editor, placement.key, paragraphId);
  });

  test('a footnote', () => {
    const editor = mount({
      body: `<w:p>${run('alpha')}<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`,
      footnotes:
        `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:type="continuationSeparator" w:id="0">` +
        `<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
        `<w:footnote w:id="1"><w:p><w:pPr>${CENTERED}</w:pPr><w:r><w:footnoteRef/></w:r>` +
        `${run(' see appendix')}</w:p></w:footnote>`,
    });
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    expect(rects).toHaveLength(1);
    const [text] = probe(editor, paragraphId, 0, length);
    expectBandsOver(rects, [text!]);
    const hits = editor.getReviewItemsAt(rects[0]!.left + 1, text!.top + 1, WHOLE);
    expect(keys(hits)).toEqual([placement.key]);
    expectActiveInItsStory(editor, placement.key, paragraphId);
  });

  test('viewing mode', () => {
    const editor = mount({ body: heading(run('Payment terms')) + PLAIN });
    expect(editor.setEditingMode('viewing').ok).toBe(true);
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const rects = editor.getReviewItemRects(placement.key, WHOLE);
    expect(rects).toHaveLength(1);
    const [text] = probe(editor, paragraphId, 0, length);
    const y = text!.top + text!.height / 2;
    expect(keys(editor.getReviewItemsAt(text!.left + 1, y, WHOLE))).toEqual([placement.key]);
    caretAt(editor, paragraphId, 1);
    expect(isActive(editor, placement.key, WHOLE)).toBe(true);
  });
});

describe('a numbered paragraph', () => {
  const NUMBERING =
    `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">` +
    `<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>` +
    `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
  const NUMBERED = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`;
  /** A spacing suggestion: the text stays left-aligned, so the number stays in the indent. */
  const SPACED = `<w:spacing w:after="240"/><w:pPrChange w:id="5" ${AUTHOR}><w:pPr/></w:pPrChange>`;

  test('the first band covers the list number, and a hit on the number returns the change', () => {
    const editor = mount({
      body: `<w:p><w:pPr>${NUMBERED}${SPACED}</w:pPr>${run('First clause')}</w:p>` + PLAIN,
      numbering: NUMBERING,
    });
    const { placement, paragraphId, length } = paragraphChange(editor, WHOLE);
    const [{ page, fragment }] = fragmentsOf(editor, paragraphId);
    const marker = fragment!.marker!.box;
    // The number sits in the hanging indent, outside the line's text column.
    expect(marker.x + marker.width).toBeLessThan(fragment!.lines[0]!.box.x);
    const [text] = probe(editor, paragraphId, 0, length);
    // The client mapping agrees with the painted text before it is used for the marker.
    expect(clientX(editor, page!, fragment!.lines[0]!.contentX)).toBeCloseTo(text!.left, 3);
    const markerLeft = clientX(editor, page!, marker.x);
    const markerRight = clientX(editor, page!, marker.x + marker.width);
    expect(markerRight).toBeLessThan(text!.left);
    const [band] = editor.getReviewItemRects(placement.key, WHOLE);
    expect(band!.left).toBeLessThanOrEqual(markerLeft + 0.01);
    expect(band!.right).toBeGreaterThanOrEqual(text!.right - 0.01);
    const y = text!.top + text!.height / 2;
    const hits = editor.getReviewItemsAt((markerLeft + markerRight) / 2, y, WHOLE);
    expect(keys(hits)).toEqual([placement.key]);
    expect(editor.getReviewItemsAt((markerLeft + markerRight) / 2, y)).toEqual([]);
  });
});
