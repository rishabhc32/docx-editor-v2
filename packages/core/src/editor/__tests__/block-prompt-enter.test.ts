// Enter in a block content control's prompt.
//
// Word 16.114 (2026-10-30, `~$p4-enter.docx` -> `~$p4-enter-tracked.docx`, Track Changes on)
// cannot hold a collapsed caret in a prompt: it selects the whole prompt, and Return REPLACES
// it. `w:showingPlcHdr` goes, the prompt's text goes with no `w:del`, and the control holds two
// empty paragraphs, the first mark a tracked `w:ins`. Reject All (`~$p4-enter-rejected.docx`)
// leaves Word's blank prompt: the prompt is gone for good. The editor split the prompt and
// kept it ("Fir" / "st line." / "Second line."); it now does what Word does, wherever the
// caret is in the prompt.

import { GlobalRegistrator } from '@happy-dom/global-registrator';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { serializeOoxmlPart } from '../../store/package/ooxml-tree.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Word's sample: a block control whose prompt has two paragraphs, between two lines. */
const BODY =
  `<w:p>${run('Before.')}</w:p>` +
  `<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/><w:sdtContent>` +
  `<w:p>${run('First line.')}</w:p><w:p>${run('Second line.')}</w:p></w:sdtContent></w:sdt>` +
  `<w:p>${run('After.')}</w:p>`;

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

function open(mode: 'edit' | 'suggesting', body = BODY): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: docx(body),
    author: 'Ana',
    mode,
    modules: [stubReviewModule()],
  });
  mounted.push(editor);
  return editor;
}

/** The control as a reader of the saved file sees it, without paragraph ids and dates. */
function control(editor: DocxEditorInstance): string {
  const xml = serializeOoxmlPart(editor.surface!.session.part());
  return (/<w:sdt>.*<\/w:sdt>/s.exec(xml)?.[0] ?? '<none>')
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:date="[^"]*"/g, '');
}

function enterAt(editor: DocxEditorInstance, paragraph: number, offset: number): void {
  const surface = editor.surface!;
  const id = surface.session.paragraphIds()[paragraph];
  if (!id) throw new Error(`no paragraph ${paragraph}`);
  surface.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
  surface.splitParagraph();
}

const INSERTED_MARK = '<w:pPr><w:rPr><w:ins w:author="Ana" w:id="0"/></w:rPr></w:pPr>';
const WORD_TRACKED =
  `<w:sdt><w:sdtPr><w:id w:val="602"/></w:sdtPr><w:sdtEndPr/><w:sdtContent>` +
  `<w:p>${INSERTED_MARK}</w:p><w:p/></w:sdtContent></w:sdt>`;
const WORD_DIRECT =
  `<w:sdt><w:sdtPr><w:id w:val="602"/></w:sdtPr><w:sdtEndPr/><w:sdtContent>` +
  `<w:p/><w:p/></w:sdtContent></w:sdt>`;
const WORD_REJECTED =
  `<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/><w:sdtContent>` +
  `<w:p><w:r><w:t xml:space="preserve">     </w:t></w:r></w:p></w:sdtContent></w:sdt>`;

describe("Enter in a block control's prompt replaces the prompt, as Word does", () => {
  const places = [
    ['the start of its first line', 1, 0],
    ['inside its first line', 1, 3],
    ['the end of its first line', 1, 11],
    ['inside its second line', 2, 6],
  ] as const;
  for (const [where, paragraph, offset] of places) {
    test(`tracked, at ${where}: two empty paragraphs, the first mark inserted`, () => {
      const editor = open('suggesting');
      enterAt(editor, paragraph, offset);
      expect(editor.surface!.state().lastRejection).toBeNull();
      expect(control(editor)).toBe(WORD_TRACKED);
    });
  }

  test('untracked: two empty paragraphs, no revision', () => {
    const editor = open('edit');
    enterAt(editor, 1, 3);
    expect(control(editor)).toBe(WORD_DIRECT);
  });

  test("Reject All leaves Word's blank prompt: the prompt is gone for good", () => {
    const editor = open('suggesting');
    enterAt(editor, 1, 3);
    expect(editor.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(
      true
    );
    expect(control(editor)).toBe(WORD_REJECTED);
  });

  test('undo brings the prompt back whole', () => {
    const editor = open('suggesting');
    const before = control(editor);
    enterAt(editor, 1, 3);
    editor.surface!.undo();
    expect(control(editor)).toBe(before);
  });

  test('a comment on the prompt keeps its reference run', () => {
    // The emptied prompt keeps its comment anchors; only the empty run made for typing goes
    const reference = '<w:r><w:commentReference w:id="4"/></w:r>';
    const commented =
      `<w:p>${run('Before.')}</w:p>` +
      `<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent>` +
      `<w:p><w:commentRangeStart w:id="4"/>${run('First line.')}<w:commentRangeEnd w:id="4"/>` +
      `${reference}</w:p></w:sdtContent></w:sdt>` +
      `<w:p>${run('After.')}</w:p>`;
    const editor = open('edit', commented);
    enterAt(editor, 1, 3);
    expect(editor.surface!.state().lastRejection).toBeNull();
    expect(control(editor)).not.toContain('showingPlcHdr');
    expect(control(editor)).toContain(reference);
  });
});
