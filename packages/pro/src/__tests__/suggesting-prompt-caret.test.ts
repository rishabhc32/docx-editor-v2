/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The caret after typing over a content control's prompt in suggesting mode.
//
// A press on a prompt selects all of it, and typing replaces it. The caret must then stay right
// after the typed words, so the next key continues them ("Bob", then "by", gives "Bobby"), as it
// does in editing mode and in Word. A caret counted from where the prompt began, while the store
// kept the prompt struck in front of the words, put the next key inside the struck prompt, and
// "by" came out before "Bob".

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { reviewModule } from '../review/review-module.ts';
import { pressBackspace, pressContentControl, typeKeys } from './surface-gestures-support.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const PROMPT = 'Click here to enter your full name';
const PROMPT_RUN = `<w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr><w:t>${PROMPT}</w:t></w:r>`;
const LABEL = 'Name: ';
/** "Name: ", an inline control showing its prompt, then " end.". */
const inline = (type: string) =>
  `<w:p><w:r><w:t xml:space="preserve">${LABEL}</w:t></w:r><w:sdt><w:sdtPr><w:showingPlcHdr/>` +
  `${type}</w:sdtPr><w:sdtContent>${PROMPT_RUN}</w:sdtContent></w:sdt>` +
  '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>';
/** A block control, as Word writes a rich text one, showing its prompt between two paragraphs. */
const BLOCK =
  '<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr>' +
  `<w:sdtContent><w:p>${PROMPT_RUN}</w:p></w:sdtContent></w:sdt>` +
  '<w:p><w:r><w:t>After</w:t></w:r></w:p>';

interface Control {
  readonly name: string;
  readonly body: string;
  /** Index of the paragraph that holds the prompt. */
  readonly paragraph: number;
  /** Offset of the prompt's first character in that paragraph. */
  readonly start: number;
  /** The text of every paragraph once `value` replaces the prompt. */
  readonly shown: (value: string) => string;
}

const CONTROLS: readonly Control[] = [
  ['plain text', '<w:text/>'],
  ['inline rich text', '<w:richText/>'],
  ['untyped', ''],
].map(([name, type]) => ({
  name: name!,
  body: inline(type!),
  paragraph: 0,
  start: LABEL.length,
  shown: (value: string) => `${LABEL}${value} end.`,
}));
const ALL_CONTROLS: readonly Control[] = [
  ...CONTROLS,
  {
    name: 'block rich text',
    body: BLOCK,
    paragraph: 1,
    start: 0,
    shown: (value: string) => `Before\n${value}\nAfter`,
  },
];

const mounted: { editor: DocxEditorInstance; container: HTMLElement }[] = [];
afterEach(() => {
  for (const { editor, container } of mounted.splice(0)) {
    editor.destroy();
    container.remove();
  }
});

function mount(
  document: Uint8Array,
  mode: 'suggesting' | 'editing',
  author = 'Bea'
): DocxEditorInstance {
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const editor = createDocxEditor({ container, document, author, modules: [reviewModule()] });
  mounted.push({ editor, container });
  if (!editor.surface) throw new Error('surface failed to mount');
  if (mode === 'suggesting') expect(editor.setEditingMode('suggesting').ok).toBe(true);
  return editor;
}

const containerOf = (editor: DocxEditorInstance) =>
  mounted.find((candidate) => candidate.editor === editor)!.container;
const pressControl = (editor: DocxEditorInstance) =>
  pressContentControl(editor, containerOf(editor));
const keys = (editor: DocxEditorInstance, text: string) =>
  typeKeys(editor, containerOf(editor), text);
const backspace = (editor: DocxEditorInstance) => pressBackspace(editor, containerOf(editor));

/** Each paragraph's text, and the collapsed caret as paragraph index and offset. */
function state(editor: DocxEditorInstance): { text: string; caret: [number, number] | null } {
  const surface = editor.surface!;
  const { anchor, head } = surface.state().selection;
  const ids = surface.session.paragraphIds();
  const collapsed = anchor.paragraphId === head.paragraphId && anchor.offset === head.offset;
  return {
    text: surface.session.bodyText(),
    caret: collapsed ? [ids.indexOf(head.paragraphId), head.offset] : null,
  };
}

function selected(editor: DocxEditorInstance): [number, number, number] {
  const surface = editor.surface!;
  const { anchor, head } = surface.state().selection;
  expect(anchor.paragraphId).toBe(head.paragraphId);
  const index = surface.session.paragraphIds().indexOf(head.paragraphId);
  return [index, Math.min(anchor.offset, head.offset), Math.max(anchor.offset, head.offset)];
}

/** The control's content, ids, dates and paragraph tags removed. */
function controlContent(editor: DocxEditorInstance): string {
  const xml = serializeOoxmlPart(editor.surface!.session.part())
    .replace(/ w:date="[^"]*"/g, '')
    .replace(/ w:id="[^"]*"/g, '');
  const content = /<w:sdtContent>(.*?)<\/w:sdtContent>/s.exec(xml)?.[1] ?? '';
  return content.replace(/<\/?w:p\b[^>]*>/g, '');
}

const revisions = (editor: DocxEditorInstance) =>
  editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [`${item.revisionKind}:${item.author}`] : []));

const typeWord = (editor: DocxEditorInstance, text: string) => editor.surface!.type(text);
const WAYS = [
  ['as a word', typeWord],
  ['one key at a time', keys],
] as const;

describe('the caret after typing over a prompt in suggesting mode', () => {
  for (const control of ALL_CONTROLS) {
    for (const [way, type] of WAYS) {
      test(`${control.name}, ${way}: typing continues after the words, as in editing mode`, () => {
        const results: Record<string, unknown> = {};
        for (const mode of ['suggesting', 'editing'] as const) {
          const editor = mount(docx(control.body), mode);
          pressControl(editor);
          expect(selected(editor)).toEqual([
            control.paragraph,
            control.start,
            control.start + PROMPT.length,
          ]);
          const at = (offset: number): [number, number] => [
            control.paragraph,
            control.start + offset,
          ];

          type(editor, 'Bob');
          expect(state(editor)).toEqual({ text: control.shown('Bob'), caret: at(3) });
          type(editor, 'by');
          expect(state(editor)).toEqual({ text: control.shown('Bobby'), caret: at(5) });
          backspace(editor);
          expect(state(editor)).toEqual({ text: control.shown('Bobb'), caret: at(4) });
          type(editor, 'y');
          expect(state(editor)).toEqual({ text: control.shown('Bobby'), caret: at(5) });
          expect(editor.surface!.state().lastRejection).toBeFalsy();
          results[mode] = state(editor);

          if (mode === 'suggesting') {
            // One proposal, inside the control, with the prompt gone rather than struck.
            expect(controlContent(editor)).toBe(
              '<w:ins w:author="Bea"><w:r><w:t>Bobby</w:t></w:r></w:ins>'
            );
            expect(serializeOoxmlPart(editor.surface!.session.part())).not.toContain(
              'showingPlcHdr'
            );
            expect(revisions(editor)).toEqual(['insert:Bea']);
          }
        }
        expect(results['suggesting']).toEqual(results['editing']);
      });
    }
  }
});

describe('a second author typing in the control later', () => {
  /** Bea's "Bob", typed over the prompt and still a suggestion, reopened for Cy. */
  async function reopenedForCy(control: Control): Promise<DocxEditorInstance> {
    const bea = mount(docx(control.body), 'suggesting');
    pressControl(bea);
    keys(bea, 'Bob');
    return mount(new Uint8Array(await bea.save()), 'suggesting', 'Cy');
  }

  for (const control of ALL_CONTROLS) {
    test(`${control.name}: Cy's words follow Bea's, and the caret follows Cy's`, async () => {
      const cy = await reopenedForCy(control);
      const paragraphId = cy.surface!.session.paragraphIds()[control.paragraph]!;
      const end = { paragraphId, offset: control.start + 3 };
      cy.surface!.setSelection({ anchor: end, head: end });
      keys(cy, 'by');
      expect(state(cy)).toEqual({
        text: control.shown('Bobby'),
        caret: [control.paragraph, control.start + 5],
      });
      expect(controlContent(cy)).toBe(
        '<w:ins w:author="Bea"><w:r><w:t>Bob</w:t></w:r></w:ins>' +
          '<w:ins w:author="Cy"><w:r><w:t>by</w:t></w:r></w:ins>'
      );
      expect(revisions(cy)).toEqual(['insert:Bea', 'insert:Cy']);
    });

    test(`${control.name}: Cy typing over Bea's value continues after Cy's words`, async () => {
      const cy = await reopenedForCy(control);
      const paragraphId = cy.surface!.session.paragraphIds()[control.paragraph]!;
      cy.surface!.setSelection({
        anchor: { paragraphId, offset: control.start },
        head: { paragraphId, offset: control.start + 3 },
      });
      keys(cy, 'Ann');
      keys(cy, 'e');
      // Bea's words are struck, not a prompt: they stay, and Cy's follow them.
      expect(state(cy)).toEqual({
        text: control.shown('BobAnne'),
        caret: [control.paragraph, control.start + 7],
      });
      expect(controlContent(cy)).toContain(
        '<w:ins w:author="Cy"><w:r><w:t>Anne</w:t></w:r></w:ins>'
      );
    });
  }
});
