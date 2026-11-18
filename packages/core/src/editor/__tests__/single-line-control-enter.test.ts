// Enter inside a content control that holds one line (task 9.13 of
// `ordered-suggestion-collaboration`). Word refuses it in a plain text control without
// `w:multiLine` and in a list or date control; the editor split the control in two. It now
// refuses by name (`single-line-control`, diagnostic `edit.single-line-control`) and changes
// nothing, in editing and in suggesting mode. Word also ignores Enter inside an inline rich
// text control, typed or untyped (Word 16.114, 2026-10-30, samples `~$i2-split.docx` and
// `~$i2n-split.docx`: one paragraph, no revision, and typing at the same caret works), so the
// editor refuses it there too. A multi-line plain text control still splits, and Enter at a
// control's edge goes beside it.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { paragraphTextOf } from '../../store/store/tree-ops.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const control = (properties: string, text: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${run(text)}</w:sdtContent></w:sdt>`;

/** One control of each kind, each in its own paragraph after a label. */
const BODY =
  `<w:p>${run('Plain: ')}${control('<w:id w:val="1"/><w:text/>', 'Ana Smith')}</w:p>` +
  `<w:p>${run('Lines: ')}${control('<w:id w:val="2"/><w:text w:multiLine="1"/>', 'Ana Smith')}</w:p>` +
  `<w:p>${run('Rich: ')}${control('<w:id w:val="3"/>', 'Ana Smith')}</w:p>` +
  `<w:p>${run('Typed: ')}${control('<w:id w:val="5"/><w:richText/>', 'Ana Smith')}</w:p>` +
  `<w:p>${run('List: ')}${control('<w:id w:val="4"/><w:comboBox><w:listItem w:displayText="Ana Smith" w:value="Ana Smith"/></w:comboBox>', 'Ana Smith')}</w:p>`;

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const mounted: DocxEditorInstance[] = [];
afterEach(() => {
  for (const editor of mounted.splice(0)) editor.destroy();
});

function open(mode: 'edit' | 'suggesting' = 'edit'): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(BODY),
    author: 'Ana',
    mode,
    modules: [stubReviewModule()],
  });
  mounted.push(editor);
  return editor;
}

/** Press Enter at `offset` of the paragraph that starts with `label`. */
function enterAt(editor: DocxEditorInstance, label: string, offset: number): void {
  const surface = editor.surface;
  if (!surface) throw new Error('no surface');
  const id = surface.session
    .paragraphIds()
    .find((candidate) => paragraphTextOf(surface.session.part(), candidate)?.startsWith(label));
  if (!id) throw new Error(`no paragraph ${label}`);
  surface.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
  surface.splitParagraph();
}

const xml = (editor: DocxEditorInstance) => serializeOoxmlPart(editor.surface!.session.part());
const paragraphs = (editor: DocxEditorInstance) => editor.surface!.session.paragraphIds().length;

describe('Enter inside a content control that holds one line', () => {
  for (const mode of ['edit', 'suggesting'] as const)
    test(`a plain text control refuses it by name (${mode})`, () => {
      const editor = open(mode);
      const before = xml(editor);
      // "Plain: Ana| Smith": inside the control
      enterAt(editor, 'Plain', 10);
      expect(editor.surface!.state().lastRejection).toBe('single-line-control');
      expect(xml(editor)).toBe(before);
    });

  test('a list control refuses it too', () => {
    const editor = open();
    const before = xml(editor);
    enterAt(editor, 'List', 9);
    expect(editor.surface!.state().lastRejection).toBe('single-line-control');
    expect(xml(editor)).toBe(before);
  });

  for (const mode of ['edit', 'suggesting'] as const)
    test(`an inline rich text control refuses it, as Word ignores it (${mode})`, () => {
      const editor = open(mode);
      const before = xml(editor);
      // "Rich: Ana| Smith" (untyped) and "Typed: Ana| Smith" (w:richText)
      enterAt(editor, 'Rich', 9);
      expect(editor.surface!.state().lastRejection).toBe('single-line-control');
      enterAt(editor, 'Typed', 10);
      expect(editor.surface!.state().lastRejection).toBe('single-line-control');
      expect(xml(editor)).toBe(before);
      // Typing at the same caret works
      editor.surface!.insertPlainText('X');
      expect(paragraphs(editor)).toBe(5);
      const typed = editor
        .surface!.session.paragraphIds()
        .map((id) => paragraphTextOf(editor.surface!.session.part(), id));
      expect(typed).toContain('Typed: AnaX Smith');
      expect(xml(editor)).toMatch(/<w:richText\/><\/w:sdtPr><w:sdtContent>.*X.*<\/w:sdtContent>/);
    });

  test('a multi-line plain text control splits', () => {
    const editor = open();
    enterAt(editor, 'Lines', 10);
    expect(editor.surface!.state().lastRejection).toBeNull();
    expect(paragraphs(editor)).toBe(6);
  });

  test("Enter at a rich text control's end goes beside it", () => {
    const editor = open();
    enterAt(editor, 'Rich', 15);
    expect(editor.surface!.state().lastRejection).toBeNull();
    expect(paragraphs(editor)).toBe(6);
  });

  test("Enter at a one-line control's end goes beside it", () => {
    const editor = open();
    enterAt(editor, 'Plain', 16);
    expect(editor.surface!.state().lastRejection).toBeNull();
    expect(paragraphs(editor)).toBe(6);
  });
});
