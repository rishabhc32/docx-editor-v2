// Rejecting a suggestion that replaced a content control's prompt shows THAT prompt again.
//
// Suggesting mode removes a prompt without a revision, as Word does, and proposes only what
// replaces it, so a rejection leaves the control empty. Word then shows the glossary block the
// control's `w:placeholder` names (§17.5.2.25). A control whose prompt lived only in its own
// content had nothing to show: Reject brought back the type's default prompt in one run ("Click
// here to enter text."), or no prompt at all in a block-level control, and the 24 prompts of a
// real contract were lost that way. The prompt now goes where Word keeps prompts, the glossary,
// before the suggestion removes it, so Reject, Reject All and a save and reopen show it as it was.
// A list or a date control can store a value beside its prompt. A value proposed over that prompt
// once wrote its own record over the stored one, so a reject showed the prompt with no value: the
// stored value now stays while the value is pending, and only an accept writes the new one.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  collectReviewItems as engineCollectReviewItems,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { EditorModule } from '../../contracts/modules.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const GREY = '<w:color w:val="808080"/>';
/** A run as the engine writes it: `xml:space` only where an edge space needs it. */
const run = (properties: string, text: string) =>
  `<w:r><w:rPr>${properties}</w:rPr>` +
  `<w:t${text.trim() === text ? '' : ' xml:space="preserve"'}>${text}</w:t></w:r>`;
const control = (properties: string, content: string) =>
  `<w:sdt><w:sdtPr>${properties}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const line = (inner: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r>${inner}` +
  '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>';

interface Case {
  readonly name: string;
  readonly body: string;
  /** The prompt exactly as the file has it, in `w:sdtContent`. */
  readonly prompt: string;
  /** Replace the prompt in suggesting mode, the way a reader does. */
  readonly replace: (editor: DocxEditorInstance) => void;
}

/** A press selects a whole prompt, and the first keystroke replaces it. */
function typeOverPrompt(paragraph: number, start: number, end: number) {
  return (editor: DocxEditorInstance) => {
    const surface = editor.surface!;
    const paragraphId = surface.session.paragraphIds()[paragraph]!;
    surface.setSelection({
      anchor: { paragraphId, offset: start },
      head: { paragraphId, offset: end },
    });
    for (const character of 'Jo') surface.type(character);
  };
}

/** A pick in the control's list or date widget. */
function pickValue(value: string) {
  return (editor: DocxEditorInstance) => {
    const surface = editor.surface!;
    const id = surface.layout().contentControls?.[0]?.id;
    expect(surface.contentControls.setValue(id!, value)).toBe(true);
  };
}

const TEXT_PROMPT = run(
  `<w:rStyle w:val="PlaceholderText"/>${GREY}`,
  'Click here to enter your full name'
);
const RICH_PROMPT = run(`<w:i/>${GREY}`, 'Enter the ') + run(`<w:b/><w:i/>${GREY}`, 'party name');
const BLOCK_PROMPT =
  '<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>' +
  `${run(`<w:i/>${GREY}`, 'Describe the scope of the services')}</w:p>`;
const TWO_PARAGRAPH_PROMPT =
  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${run(GREY, 'Describe the scope')}</w:p>` +
  `<w:p>${run(`<w:i/>${GREY}`, 'and the term')}</w:p>`;
const LIST_PROMPT = run(`<w:rStyle w:val="PlaceholderText"/>${GREY}`, 'Choose a payment term');
const DATE_PROMPT = run(`<w:rStyle w:val="PlaceholderText"/>${GREY}`, 'Pick the start date');

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/><w:listItem w:displayText="Net 60" w:value="60"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';

/** A control's properties beside a stored value, its prompt, the value picked, and the records. */
const STORED = [
  [
    'drop-down list',
    `<w:showingPlcHdr/><w:dropDownList w:lastValue="30">${ITEMS}</w:dropDownList>`,
    LIST_PROMPT,
    '60',
    ['w:lastValue="30"', 'w:lastValue="60"'],
  ],
  [
    'combo box',
    `<w:showingPlcHdr/><w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`,
    LIST_PROMPT,
    '60',
    ['w:lastValue="30"', 'w:lastValue="60"'],
  ],
  [
    'date',
    `<w:showingPlcHdr/><w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`,
    DATE_PROMPT,
    '2026-03-04',
    ['w:fullDate="2026-06-30T00:00:00Z"', 'w:fullDate="2026-03-04T00:00:00Z"'],
  ],
] as const;

const CASES: readonly Case[] = [
  {
    name: 'plain text control',
    body: line(control('<w:id w:val="11"/><w:showingPlcHdr/><w:text/>', TEXT_PROMPT)),
    prompt: TEXT_PROMPT,
    replace: typeOverPrompt(0, 6, 6 + 'Click here to enter your full name'.length),
  },
  {
    name: 'inline rich text control with a two-run prompt',
    body: line(control('<w:alias w:val="Party"/><w:showingPlcHdr/>', RICH_PROMPT)),
    prompt: RICH_PROMPT,
    replace: typeOverPrompt(0, 6, 6 + 'Enter the party name'.length),
  },
  {
    name: 'block rich text control',
    body:
      control('<w:alias w:val="Clause"/><w:showingPlcHdr/>', BLOCK_PROMPT) +
      '<w:p><w:r><w:t>After.</w:t></w:r></w:p>',
    prompt: BLOCK_PROMPT,
    replace: typeOverPrompt(0, 0, 'Describe the scope of the services'.length),
  },
  {
    name: 'block rich text control with a two-paragraph prompt',
    body:
      control('<w:alias w:val="Clause"/><w:showingPlcHdr/>', TWO_PARAGRAPH_PROMPT) +
      '<w:p><w:r><w:t>After.</w:t></w:r></w:p>',
    prompt: TWO_PARAGRAPH_PROMPT,
    // A caret at the prompt's edge types into the control and replaces the whole prompt.
    replace: typeOverPrompt(0, 0, 0),
  },
  {
    name: 'drop-down list control',
    body: line(
      control(
        '<w:showingPlcHdr/><w:dropDownList><w:listItem w:displayText="Net 30" w:value="30"/>' +
          '<w:listItem w:displayText="Net 60" w:value="60"/></w:dropDownList>',
        LIST_PROMPT
      )
    ),
    prompt: LIST_PROMPT,
    replace: pickValue('60'),
  },
  {
    name: 'date control',
    body: line(
      control(
        '<w:showingPlcHdr/><w:date><w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/></w:date>',
        DATE_PROMPT
      )
    ),
    prompt: DATE_PROMPT,
    replace: pickValue('2026-03-04'),
  },
];

function docx(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
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
        return revisionItemsOf({ ...part, root: paragraph });
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

function open(bytes: Uint8Array, author = 'Bea'): DocxEditorInstance {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    author,
    modules: [reviewModule()],
  });
  if (!editor.surface) throw new Error('surface failed to mount');
  editors.push(editor);
  return editor;
}

function xml(editor: DocxEditorInstance): string {
  return serializeOoxmlPart(editor.surface!.session.part());
}

/** What the control's `w:sdtContent` holds, ids and paragraph ids removed. */
function content(editor: DocxEditorInstance): string {
  const source = xml(editor).replace(/ w14:(paraId|textId)="[^"]*"/g, '');
  return /<w:sdtContent>(.*?)<\/w:sdtContent>/s.exec(source)?.[1] ?? '';
}

function properties(editor: DocxEditorInstance): string {
  return /<w:sdtPr>(.*?)<\/w:sdtPr>/s.exec(xml(editor))?.[1] ?? '';
}

function suggest(editor: DocxEditorInstance): void {
  expect(editor.setEditingMode('suggesting').ok).toBe(true);
}

function rejectOne(editor: DocxEditorInstance): void {
  const items = editor.getReviewItems();
  expect(items).toHaveLength(1);
  expect(editor.rejectReviewItem(items[0]!.key).ok).toBe(true);
}

function acceptOne(editor: DocxEditorInstance): void {
  const items = editor.getReviewItems();
  expect(items).toHaveLength(1);
  expect(editor.acceptReviewItem(items[0]!.key).ok).toBe(true);
}

function rejectAll(editor: DocxEditorInstance): void {
  const result = editor.exec({
    type: 'resolveAllReviewChanges',
    action: 'reject',
    scope: 'document',
  });
  expect(result.ok).toBe(true);
}

function expectPromptBack(editor: DocxEditorInstance, prompt: string): void {
  expect(editor.getReviewItems()).toHaveLength(0);
  expect(content(editor)).toBe(prompt);
  expect(properties(editor)).toContain('<w:showingPlcHdr/>');
  expect(properties(editor)).not.toContain('w:lastValue');
  expect(properties(editor)).not.toContain('w:fullDate');
}

describe('rejecting a suggestion over a content control prompt', () => {
  for (const sample of CASES) {
    for (const [label, decide] of [
      ['Reject', rejectOne],
      ['Reject All', rejectAll],
    ] as const) {
      test(`${label} shows the ${sample.name}'s own prompt again`, () => {
        const editor = open(docx(sample.body));
        expect(content(editor)).toBe(sample.prompt);
        suggest(editor);
        sample.replace(editor);
        expect(editor.surface!.state().lastRejection).toBeFalsy();
        expect(properties(editor)).not.toContain('<w:showingPlcHdr/>');
        expect(content(editor)).toContain('<w:ins ');
        decide(editor);
        expectPromptBack(editor, sample.prompt);
      });
    }
  }

  test('the prompt survives a save and reopen while the suggestion is pending', async () => {
    for (const sample of CASES) {
      const editor = open(docx(sample.body));
      suggest(editor);
      sample.replace(editor);
      const saved = new Uint8Array(await editor.save());

      // Where Word keeps a prompt: a placeholder block in the glossary, named by the control.
      const files = unzipSync(saved);
      const glossary = strFromU8(files['word/glossary/document.xml']!);
      const name = /<w:placeholder><w:docPart w:val="([^"]+)"\/><\/w:placeholder>/.exec(
        strFromU8(files['word/document.xml']!)
      )?.[1];
      expect(name).toBeDefined();
      expect(glossary).toContain(`<w:name w:val="${name!}"/>`);
      expect(glossary).toContain('<w:gallery w:val="placeholder"/>');
      expect(strFromU8(files['word/_rels/document.xml.rels']!)).toContain(
        'relationships/glossaryDocument'
      );
      expect(strFromU8(files['[Content_Types].xml']!)).toContain('document.glossary+xml');

      const reopened = open(saved, 'Cy');
      rejectAll(reopened);
      expectPromptBack(reopened, sample.prompt);
    }
  });

  test('undo takes the glossary block back with the suggestion', async () => {
    const [sample] = CASES;
    const editor = open(docx(sample!.body));
    const before = xml(editor);
    suggest(editor);
    sample!.replace(editor);
    while (xml(editor) !== before && editor.surface!.session.canUndo()) editor.surface!.undo();
    expect(xml(editor)).toBe(before);
    const files = unzipSync(new Uint8Array(await editor.save()));
    expect(files['word/glossary/document.xml']).toBeUndefined();
  });

  test('two controls with one prompt share one glossary block', async () => {
    const twice =
      line(control('<w:showingPlcHdr/><w:text/>', TEXT_PROMPT)) +
      line(control('<w:showingPlcHdr/><w:text/>', TEXT_PROMPT));
    const editor = open(docx(twice));
    suggest(editor);
    const end = 6 + 'Click here to enter your full name'.length;
    typeOverPrompt(0, 6, end)(editor);
    typeOverPrompt(1, 6, end)(editor);
    const files = unzipSync(new Uint8Array(await editor.save()));
    const glossary = strFromU8(files['word/glossary/document.xml']!);
    expect(glossary.match(/<w:docPart>/g)).toHaveLength(1);
    rejectAll(editor);
    const source = xml(editor).replace(/ w14:(paraId|textId)="[^"]*"/g, '');
    expect(source.match(/Click here to enter your full name/g)).toHaveLength(2);
  });
});

describe('a value proposed over a prompt the control shows beside a stored value', () => {
  for (const [name, props, prompt, value, [stored, written]] of STORED) {
    test(`the ${name} keeps its stored value on a reject and takes the new one on an accept`, async () => {
      const file = docx(line(control(props, prompt)));
      for (const [decide, record] of [
        [rejectOne, stored],
        [acceptOne, written],
      ] as const) {
        // Live, and in a copy saved while the value is pending
        const editor = open(file);
        suggest(editor);
        pickValue(value)(editor);
        const saved = new Uint8Array(await editor.save());
        for (const copy of [editor, open(saved, 'Cy')]) {
          decide(copy);
          expect(properties(copy)).toContain(record);
          expect(properties(copy).includes('<w:showingPlcHdr/>')).toBe(decide === rejectOne);
        }
      }
    });
  }
});

describe('a date picked over a date in a format that shows no day', () => {
  test('a reject gives back the date the file stores, live and after a save', async () => {
    const props = `<w:date w:fullDate="2026-06-30T00:00:00Z"><w:dateFormat w:val="MMMM yyyy"/></w:date>`;
    const editor = open(docx(line(control(props, run('', 'June 2026')))));
    suggest(editor);
    pickValue('2026-09-15')(editor);
    const saved = new Uint8Array(await editor.save());
    for (const copy of [editor, open(saved, 'Cy')]) {
      rejectOne(copy);
      expect(properties(copy)).toContain('w:fullDate="2026-06-30T00:00:00Z"');
      expect(content(copy)).toContain('June 2026');
    }

    // An accept reads the date back from the words, at the day the format cannot show
    const accepted = open(saved, 'Cy');
    acceptOne(accepted);
    expect(properties(accepted)).toContain('w:fullDate="2026-09-01T00:00:00Z"');
  });
});
