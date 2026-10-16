// A toggle press that takes a format suggestion back leaves none (`inherited-toggle-writes.ts`).
//
// Pressing Underline twice over plain text in suggesting mode left `w:u w:val="none"` on the
// run and a `w:rPrChange` whose original had no underline, so the review pane kept a
// "Changed text formatting" card for a change nobody could see.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { collectReviewItems, serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { mountPaginatedSurface, type PaginatedSurface } from '../paginated-surface.ts';
import { runToolbarCommand } from '../toolbar-commands.ts';
import { FORMAT_DOC_DEFAULTS, selectCellRectangle } from './paginated-surface-fixtures.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAIN_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const STYLES =
  '<w:style w:type="character" w:styleId="Underlined"><w:name w:val="Underlined"/>' +
  '<w:rPr><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Strong"><w:name w:val="Strong"/>' +
  '<w:rPr><w:b/></w:rPr></w:style>';

function styledDocx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN_CT}.document.main+xml"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${MAIN_CT}.styles+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
    'word/styles.xml': strToU8(
      `<w:styles xmlns:w="${W}">${FORMAT_DOC_DEFAULTS}${STYLES}</w:styles>`
    ),
  });
}

function withSurface(
  body: string,
  suggesting: boolean,
  run: (surface: PaginatedSurface, select: (start: number, end: number) => void) => void
): void {
  const container = document.createElement('div');
  document.body.append(container);
  const opened = mountPaginatedSurface(container, styledDocx(body), { author: 'Ada' });
  if (!opened.ok) throw new Error(opened.reason);
  const surface = opened.surface;
  if (suggesting) surface.setEditingMode('suggest');
  const paragraphId = surface.session.paragraphIds()[0]!;
  const select = (start: number, end: number) =>
    surface.setSelection({
      anchor: { paragraphId, offset: start },
      head: { paragraphId, offset: end },
    });
  try {
    run(surface, select);
  } finally {
    surface.destroy();
    container.remove();
  }
}

const xmlOf = (surface: PaginatedSurface) => serializeOoxmlPart(surface.session.part());
const cardsOf = (surface: PaginatedSurface) =>
  collectReviewItems({ storyPart: surface.session.part() });

describe('a toggle press writes nothing the styles already give', () => {
  test('Underline pressed twice over plain text leaves no suggestion', () => {
    withSurface('<w:p><w:r><w:t>one two</w:t></w:r></w:p>', true, (surface, select) => {
      select(0, 3);
      surface.toggleRunProperty('u', { val: 'single' });
      expect(cardsOf(surface)).toHaveLength(1);
      surface.toggleRunProperty('u', { val: 'single' });
      expect(xmlOf(surface)).not.toContain('w:u');
      expect(xmlOf(surface)).not.toContain('rPrChange');
      expect(cardsOf(surface)).toEqual([]);
    });
  });

  test('Bold pressed twice over a whole paragraph restores its runs and its mark', () => {
    withSurface('<w:p><w:r><w:t>one two</w:t></w:r></w:p>', true, (surface, select) => {
      select(0, 7);
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).toContain('<w:pPr><w:rPr><w:b/>');
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).not.toMatch(/<w:b[ />]/);
      expect(xmlOf(surface)).not.toContain('rPrChange');
      expect(cardsOf(surface)).toEqual([]);
    });
  });

  test('a right-to-left run drops the complex-script twin too', () => {
    const body = '<w:p><w:r><w:rPr><w:rtl/></w:rPr><w:t>שלום</w:t></w:r></w:p>';
    withSurface(body, true, (surface, select) => {
      select(0, 4);
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).toContain('w:bCs');
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).not.toMatch(/<w:bC?s?[ />]/);
      expect(cardsOf(surface)).toEqual([]);
    });
  });

  test('Underline off over a style that underlines still writes the off value', () => {
    const body =
      '<w:p><w:r><w:rPr><w:rStyle w:val="Underlined"/></w:rPr><w:t>one</w:t></w:r></w:p>';
    withSurface(body, true, (surface, select) => {
      select(0, 3);
      surface.toggleRunProperty('u', { val: 'single' });
      expect(xmlOf(surface)).toContain('<w:u w:val="none"/>');
      expect(cardsOf(surface)).toHaveLength(1);
    });
  });

  test('Bold on again over a bold style removes the off value it wrote', () => {
    const body =
      '<w:p><w:pPr><w:pStyle w:val="Strong"/></w:pPr><w:r><w:t>one two</w:t></w:r></w:p>';
    withSurface(body, true, (surface, select) => {
      select(0, 3);
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).toContain('<w:b w:val="0"/>');
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).not.toMatch(/<w:b[ />]/);
      expect(cardsOf(surface)).toEqual([]);
    });
  });

  test('an off value the original records is kept, so the record goes', () => {
    const body = '<w:p><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:t>one</w:t></w:r></w:p>';
    withSurface(body, true, (surface, select) => {
      select(0, 3);
      surface.toggleRunProperty('b');
      expect(cardsOf(surface)).toHaveLength(1);
      surface.toggleRunProperty('b');
      expect(xmlOf(surface)).toContain('<w:b w:val="0"/>');
      expect(xmlOf(surface)).not.toContain('rPrChange');
      expect(cardsOf(surface)).toEqual([]);
    });
  });

  test('without a pending format suggestion the off value stays', () => {
    withSurface('<w:p><w:r><w:t>one two</w:t></w:r></w:p>', false, (surface, select) => {
      select(0, 3);
      surface.toggleRunProperty('i');
      surface.toggleRunProperty('i');
      expect(xmlOf(surface)).toContain('<w:i w:val="0"/>');
    });
  });
});

/** A suggesting editor over `body`, its first paragraph's `[start, end)` selected. */
function suggestingEditor(body: string, start = 0, end = 3): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: styledDocx(body),
    author: 'Ada',
    mode: 'suggesting',
    modules: [stubReviewModule()],
  });
  editors.push(editor);
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
  return editor;
}

const editors: DocxEditorInstance[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
});

const editorCards = (editor: DocxEditorInstance) =>
  collectReviewItems({ storyPart: editor.surface!.session.part() });

/** The saved file's `word/document.xml`. */
async function savedXml(editor: DocxEditorInstance): Promise<string> {
  return strFromU8(unzipSync(new Uint8Array(await editor.save()))['word/document.xml']!);
}

/** An explicit off value on any toggle, or a format record, anywhere in the saved body. */
const LEFTOVER = /w:rPrChange|<w:(?:b|i|strike|bCs|iCs) w:val="0"|w:val="none"|w:val="baseline"/;

const PLAIN = '<w:p><w:r><w:t>one two</w:t></w:r></w:p>';

describe('every toggle, from every entry point, pressed twice as a suggester', () => {
  const MARKS = ['bold', 'italic', 'underline', 'strike', 'superscript', 'subscript'] as const;

  for (const mark of MARKS) {
    test(`${mark}: one press is a card, two presses leave nothing in the saved file`, async () => {
      const editor = suggestingEditor(PLAIN);
      expect(editor.exec({ type: 'toggleMark', mark }).ok).toBe(true);
      expect(editorCards(editor)).toHaveLength(1);
      expect(editor.exec({ type: 'toggleMark', mark }).ok).toBe(true);
      expect(editorCards(editor)).toEqual([]);
      expect(await savedXml(editor)).not.toMatch(LEFTOVER);
    });
  }

  test('the toolbar buttons take the suggestion back', async () => {
    const editor = suggestingEditor(PLAIN);
    for (const slot of ['text.bold', 'text.underline'] as const) {
      runToolbarCommand(editor, slot);
      runToolbarCommand(editor, slot);
    }
    expect(editorCards(editor)).toEqual([]);
    expect(await savedXml(editor)).not.toMatch(LEFTOVER);
  });

  test('the keyboard shortcuts take the suggestion back', async () => {
    const editor = suggestingEditor(PLAIN);
    const pages = containers.at(-1)!.querySelector<HTMLElement>('.docx-pages')!;
    const press = (key: string, shiftKey = false) =>
      pages.dispatchEvent(
        new KeyboardEvent('keydown', { key, ctrlKey: true, shiftKey, bubbles: true })
      );
    for (const [key, shift] of [['b'], ['i'], ['u'], ['+', true], ['=']] as const) {
      press(key, shift);
      expect(editorCards(editor)).toHaveLength(1);
      press(key, shift);
      expect(editorCards(editor)).toEqual([]);
    }
    expect(await savedXml(editor)).not.toMatch(LEFTOVER);
  });

  test('selected table cells take the suggestion back', async () => {
    const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    const table =
      '<w:tbl><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('left')}${cell('right')}</w:tr></w:tbl>`;
    const editor = suggestingEditor(`${table}${PLAIN}`);
    selectCellRectangle(editor.surface!, { row: 0, column: 0 }, { row: 0, column: 1 });
    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    expect(editorCards(editor).length).toBeGreaterThan(0);
    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
    expect(editorCards(editor)).toEqual([]);
    expect(await savedXml(editor)).not.toMatch(LEFTOVER);
  });

  test("another author's pending record stays when the suggester turns the format off", () => {
    const record =
      '<w:rPrChange w:id="7" w:author="Bob" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>';
    const editor = suggestingEditor(
      `<w:p><w:r><w:rPr><w:u w:val="single"/>${record}</w:rPr><w:t>one</w:t></w:r></w:p>`
    );
    expect(editor.exec({ type: 'toggleMark', mark: 'underline' }).ok).toBe(true);
    expect(serializeOoxmlPart(editor.surface!.session.part())).toContain('w:author="Bob"');
  });
});
