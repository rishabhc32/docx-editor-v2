// AN UNTYPED CONTENT CONTROL'S PROMPT, THROUGH THE EDITOR.
//
// Word writes no type element for the rich text controls it inserts, so a control without one
// is rich text. A press on its prompt selects the prompt whole, typing replaces it, deleting
// everything typed shows the prompt again, and the next keystroke replaces that prompt too,
// exactly as for a `w:richText` control.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  contentControlTextOf,
  contentControlsIn,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { PaginatedSurface } from '../paginated-surface.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const PROMPT = 'Click here to enter text.';
const PROMPT_RUN = `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${PROMPT}</w:t></w:r>`;

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

const editors: DocxEditorInstance[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
});

function mount(body: string): { surface: PaginatedSurface; container: HTMLElement } {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({ container, document: docxOf(body) });
  editors.push(editor);
  if (!editor.surface) throw new Error('surface failed to mount');
  return { surface: editor.surface, container };
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

function controlState(surface: PaginatedSurface) {
  const control = contentControlsIn(surface.session.part().root)[0]!.node;
  const xml = /<w:sdt>.*<\/w:sdt>/s.exec(serializeOoxmlPart(surface.session.part()))![0];
  return {
    text: contentControlTextOf(control),
    prompt: xml.includes('<w:showingPlcHdr/>'),
    // The one type element these tests vary is taken out, so the two shapes compare.
    xml: xml.replace('<w:richText/>', ''),
  };
}

/** Press the prompt, type over it, delete what was typed, and type again. */
function typeEmptyRetype(body: string) {
  const { surface, container } = mount(body);
  pressPrompt(surface, container);
  surface.type('Bob');
  const typed = controlState(surface);
  for (let index = 0; index < 'Bob'.length; index += 1) surface.deleteBackward();
  const emptied = controlState(surface);
  surface.type('Z');
  const retyped = controlState(surface);
  expect(surface.state().lastRejection ?? null).toBeNull();
  return { typed, emptied, retyped };
}

const inline = (type: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>` +
  `<w:showingPlcHdr/>${type}</w:sdtPr><w:sdtContent>${PROMPT_RUN}</w:sdtContent></w:sdt>` +
  `<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>`;
const block = (type: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="7"/><w:showingPlcHdr/>${type}</w:sdtPr>` +
  `<w:sdtContent><w:p>${PROMPT_RUN}</w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>after</w:t></w:r></w:p>`;

describe('an untyped control with a prompt behaves like a rich text control', () => {
  for (const [name, shape] of [
    ['an inline', inline],
    ['a block', block],
  ] as const) {
    test(`${name} untyped control: type over the prompt, empty it, and type again`, () => {
      const untyped = typeEmptyRetype(shape(''));
      expect(untyped.typed).toMatchObject({ text: 'Bob', prompt: false });
      expect(untyped.emptied).toMatchObject({ text: PROMPT, prompt: true });
      expect(untyped.retyped).toMatchObject({ text: 'Z', prompt: false });
      expect(untyped).toEqual(typeEmptyRetype(shape('<w:richText/>')));
    });
  }
});
