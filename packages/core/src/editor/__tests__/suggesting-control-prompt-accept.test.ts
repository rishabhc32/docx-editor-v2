// ACCEPTING A SUGGESTION OVER A CONTENT CONTROL'S PROMPT, THROUGH THE EDITOR.
//
// A press on a prompt selects it whole, and in suggesting mode the typed words replace it as
// Word does: the prompt goes without a revision, `w:showingPlcHdr` goes with it, and only the
// words are proposed. Accept leaves the control holding real content: no `w:showingPlcHdr`, and
// no `PlaceholderText` style on the words, which keep the prompt's other run properties. The
// next keystroke then edits the words instead of replacing them whole, and a save and reopen
// keeps them. Reject shows the control's own prompt again.
//
// Older versions struck the prompt in a `w:del` and kept `w:showingPlcHdr` while the decision
// was open. A file saved that way still opens with the struck prompt and the pending words, and
// accepting them must end the prompt too, or the next keystroke replaced the accepted words.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  collectReviewItems as engineCollectReviewItems,
  contentControlTextOf,
  contentControlsIn,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const PROMPT = 'Click here to enter text.';
const STYLED = '<w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr>';
const PROMPT_RUN = `<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`;

function docxOf(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr/></w:body></w:document>`
    ),
  });
}

/** The engine's own review derivation, wired as a module (core may not import pro). */
function reviewModule(): EditorModule {
  return {
    id: 'review',
    review: {
      displayModes: ['all-markup', 'proposed', 'original'],
      collectReviewItems: engineCollectReviewItems,
      revisionItemsOfParagraph: (part, paragraphId) => {
        const paragraph = findNode(part, paragraphId);
        if (!paragraph || paragraph.kind !== 'paragraph') return [];
        return revisionItemsOf({
          id: part.id,
          name: part.name,
          contentType: part.contentType,
          root: paragraph,
        });
      },
    },
  };
}

const editors: DocxEditorInstance[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
});

function mount(document_: Uint8Array, author = 'Ana') {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: document_,
    author,
    modules: [reviewModule()],
  });
  editors.push(editor);
  if (!editor.surface) throw new Error('surface failed to mount');
  return { editor, container };
}

/** A press on the control's prompt, by layout geometry, which selects the prompt whole. */
function pressPrompt(surface: PaginatedSurface, container: HTMLElement): void {
  const control = surface.layout().contentControls![0]!;
  const fragment = control.fragments[0]!;
  const page = surface.layout().pages[fragment.pageIndex]!;
  const x = page.contentBox.x + fragment.box.x + fragment.box.width / 2;
  const y = page.contentBox.y + fragment.box.y + fragment.box.height / 2;
  const pages = container.querySelector<HTMLElement>('.docx-pages')!;
  for (const type of ['pointerdown', 'pointerup'] as const) {
    pages.dispatchEvent(
      new PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        button: 0,
        pointerId: 1,
        pointerType: 'mouse',
        clientX: x,
        clientY: y,
      })
    );
  }
}

function controlOf(editor: DocxEditorInstance) {
  const found = contentControlsIn(editor.surface!.session.part().root)[0];
  if (!found) throw new Error('no control');
  return found.node;
}

/** The serialized `w:sdt`, as a reader of the saved file sees it. */
function controlXml(editor: DocxEditorInstance): string {
  const xml = serializeOoxmlPart(editor.surface!.session.part());
  const found = /<w:sdt>.*<\/w:sdt>/s.exec(xml)?.[0];
  if (found === undefined) throw new Error('no control XML');
  return found;
}

/** What the control's `w:sdtContent` holds, paragraph ids removed. */
function contentXml(editor: DocxEditorInstance): string {
  const xml = controlXml(editor).replace(/ w14:(paraId|textId)="[^"]*"/g, '');
  return /<w:sdtContent>(.*)<\/w:sdtContent>/s.exec(xml)?.[1] ?? '';
}

/** Suggest `text` over the prompt the way a reader does: press the prompt, then type. */
function suggestOverPrompt(body: string, text: string) {
  const { editor, container } = mount(docxOf(body));
  const surface = editor.surface!;
  const original = contentXml(editor);
  expect(editor.setEditingMode('suggesting').ok).toBe(true);
  pressPrompt(surface, container);
  const { anchor, head } = surface.state().selection;
  expect(Math.abs(head.offset - anchor.offset)).toBe(PROMPT.length);
  surface.type(text);
  expect(surface.state().lastRejection ?? null).toBeNull();
  // Pending: the prompt is gone without a revision, and only the words are proposed.
  const xml = controlXml(editor);
  expect(xml).not.toContain(PROMPT);
  expect(xml).not.toContain('<w:showingPlcHdr/>');
  expect(xml).not.toContain('PlaceholderText');
  expect(xml).not.toMatch(/<w:del /);
  expect(xml).toMatch(/<w:ins /);
  return { editor, original };
}

function expectWordsNotPrompt(editor: DocxEditorInstance, text: string): void {
  const xml = controlXml(editor);
  expect(contentControlTextOf(controlOf(editor))).toBe(text);
  expect(xml).not.toContain('showingPlcHdr');
  expect(xml).not.toContain('PlaceholderText');
  expect(xml).not.toMatch(/<w:(ins|del) /);
  // The prompt's other run properties stay with the words.
  expect(xml).toContain(`<w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r>`);
  expect(editor.getReviewItems()).toHaveLength(0);
}

function expectPromptBack(editor: DocxEditorInstance, original: string): void {
  expect(editor.getReviewItems()).toHaveLength(0);
  expect(contentXml(editor)).toBe(original);
  expect(controlXml(editor)).toContain('<w:showingPlcHdr/>');
}

const inline = (properties: string, content = PROMPT_RUN) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>` +
  `<w:showingPlcHdr/>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>` +
  `<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>`;
const block = (properties: string, content = PROMPT_RUN) =>
  `<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/>${properties}</w:sdtPr>` +
  `<w:sdtContent><w:p>${content}</w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>after</w:t></w:r></w:p>`;

const DROP_DOWN =
  '<w:dropDownList><w:listItem w:displayText="One" w:value="1"/>' +
  '<w:listItem w:displayText="Two" w:value="2"/></w:dropDownList>';
const DATE_FORMAT = '<w:date><w:dateFormat w:val="yyyy-MM-dd"/><w:lid w:val="en-US"/></w:date>';

const CASES = [
  ['an inline plain text control', inline, '<w:text/>', 'Bob'],
  ['an inline rich text control', inline, '<w:richText/>', 'Bob'],
  ['an inline untyped control', inline, '', 'Bob'],
  ['a block rich text control', block, '<w:richText/>', 'Bob'],
  ['a block untyped control', block, '', 'Bob'],
  ['a drop-down list', inline, DROP_DOWN, 'Two'],
  ['a date control', inline, DATE_FORMAT, '2026-03-04'],
] as const;

describe('accepting words suggested over a content control prompt', () => {
  for (const [name, shape, properties, text] of CASES) {
    describe(name, () => {
      test('Accept on the one review item leaves the words, not a prompt', () => {
        const { editor } = suggestOverPrompt(shape(properties), text);
        const items = editor.getReviewItems();
        expect(items).toHaveLength(1);
        expect(editor.acceptReviewItem(items[0]!.key)).toMatchObject({ ok: true, changed: true });
        expectWordsNotPrompt(editor, text);
      });

      test('Accept All leaves the words, and a save and reopen keeps them', async () => {
        const { editor } = suggestOverPrompt(shape(properties), text);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
        expectWordsNotPrompt(editor, text);
        const { editor: reopened } = mount(new Uint8Array(await editor.save()));
        expectWordsNotPrompt(reopened, text);
      });

      test("Reject All shows the control's own prompt again", () => {
        const { editor, original } = suggestOverPrompt(shape(properties), text);
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
        expectPromptBack(editor, original);
      });
    });
  }

  test('after Accept, the next keystroke edits the words instead of replacing them', () => {
    const { editor } = suggestOverPrompt(inline('<w:text/>'), 'Bob');
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(editor.setEditingMode('editing').ok).toBe(true);
    const surface = editor.surface!;
    const paragraphId = surface.session.paragraphIds()[0]!;
    // "Name: Bob end": the caret after "Bob", inside the control.
    const position = { paragraphId, offset: 9 };
    surface.setSelection({ anchor: position, head: position });
    surface.type('s');
    expect(contentControlTextOf(controlOf(editor))).toBe('Bobs');
    expect(surface.session.bodyText()).toBe('Name: Bobs end');
  });

  for (const [name, properties, value, shown] of [
    ['a drop-down list', DROP_DOWN, '2', 'Two'],
    ['a date control', DATE_FORMAT, '2026-03-04', '2026-03-04'],
  ] as const) {
    test(`a value chosen for ${name} in suggesting mode is plain content once accepted`, () => {
      const { editor } = mount(docxOf(inline(properties)));
      const surface = editor.surface!;
      expect(editor.setEditingMode('suggesting').ok).toBe(true);
      const id = surface.layout().contentControls![0]!.id;
      expect(surface.contentControls.setValue(id, value)).toBe(true);
      for (const item of editor.getReviewItems()) editor.acceptReviewItem(item.key);
      expectWordsNotPrompt(editor, shown);
    });
  }
});

describe('accepting a struck prompt that an older version saved', () => {
  const DATE = '2026-10-01T10:00:00Z';
  /** The prompt struck and the words proposed, as older versions saved them. */
  const STRUCK_AND_TYPED = (words: string) =>
    `<w:del w:id="1" w:author="Ana" w:date="${DATE}"><w:r>${STYLED}<w:delText>${PROMPT}</w:delText></w:r></w:del>` +
    `<w:ins w:id="2" w:author="Ana" w:date="${DATE}"><w:r>${STYLED}<w:t>${words}</w:t></w:r></w:ins>`;

  const LEGACY = [
    ['an inline plain text control', inline, '<w:text/>', 'Bob'],
    ['an inline untyped control', inline, '', 'Bob'],
    ['a block untyped control', block, '', 'Bob'],
    ['a drop-down list', inline, DROP_DOWN, 'Two'],
  ] as const;

  for (const [name, shape, properties, words] of LEGACY) {
    describe(name, () => {
      const open = () => {
        const { editor } = mount(docxOf(shape(properties, STRUCK_AND_TYPED(words))), 'Cy');
        // The file opens as it was saved: the prompt struck, the words pending, the flag on.
        expect(controlXml(editor)).toContain(`<w:delText>${PROMPT}</w:delText>`);
        expect(controlXml(editor)).toContain('<w:showingPlcHdr/>');
        return editor;
      };

      test('Accept on the one review item leaves the words, not a prompt', () => {
        const editor = open();
        const items = editor.getReviewItems();
        expect(items).toHaveLength(1);
        expect(editor.acceptReviewItem(items[0]!.key)).toMatchObject({ ok: true, changed: true });
        expectWordsNotPrompt(editor, words);
      });

      test('Accept All leaves the words, and a save and reopen keeps them', async () => {
        const editor = open();
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
        expectWordsNotPrompt(editor, words);
        const { editor: reopened } = mount(new Uint8Array(await editor.save()));
        expectWordsNotPrompt(reopened, words);
      });

      test('Reject All gives the prompt back exactly', () => {
        const editor = open();
        const { editor: prompted } = mount(docxOf(shape(properties)));
        expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
        expect(controlXml(editor)).toBe(controlXml(prompted));
      });
    });
  }

  test('a drop-down list records the item it shows once accepted', () => {
    const { editor } = mount(docxOf(inline(DROP_DOWN, STRUCK_AND_TYPED('Two'))), 'Cy');
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' }).ok).toBe(true);
    expect(controlXml(editor)).toContain('<w:dropDownList w:lastValue="2">');
  });
});
