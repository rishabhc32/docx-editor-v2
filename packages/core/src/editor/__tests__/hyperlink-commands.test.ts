// The link commands' tooltip, and the `Hyperlink` style a new link brings with it.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { describe, expect, test } from 'bun:test';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import { stylesPartOf } from '../../store/package/ooxml-indexes.ts';
import { stubReviewModule } from './review-test-module.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

function docx(body: string, stylesXml = ''): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        (stylesXml
          ? '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          : '') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        (stylesXml ? `<Relationship Id="rIdStyles" Type="${R}/styles" Target="styles.xml"/>` : '') +
        '</Relationships>'
    ),
    ...(stylesXml ? { 'word/styles.xml': strToU8(stylesXml) } : {}),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const PID = (index: number) => `/word/document.xml#0.0.${index}`;

interface Mounted {
  readonly editor: DocxEditorInstance;
  readonly container: HTMLElement;
}

function mount(body: string, stylesXml = ''): Mounted {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: docx(body, stylesXml) });
  if (!editor.surface) throw new Error('surface failed to mount');
  return { editor, container };
}

function select(mounted: Mounted, index: number, start: number, end: number): void {
  mounted.editor.surface!.setSelection({
    anchor: { paragraphId: PID(index), offset: start },
    head: { paragraphId: PID(index), offset: end },
  });
}

function caret(mounted: Mounted, index: number, offset: number): void {
  select(mounted, index, offset, offset);
}

describe('a link’s tooltip through the command', () => {
  test('insertHyperlink writes the tooltip that hyperlinkAt reads back', () => {
    const mounted = mount(p('Visit example today'));
    select(mounted, 0, 6, 13);
    expect(
      mounted.editor.exec({
        type: 'insertHyperlink',
        href: 'https://example.com',
        tooltip: 'Our home page',
      }).ok
    ).toBe(true);
    caret(mounted, 0, 8);
    expect(mounted.editor.query({ type: 'hyperlinkAt' })?.tooltip).toBe('Our home page');
    // On the link the caret is in, a new tooltip replaces the old one, and leaving it out
    // keeps the one the link has.
    const retarget = (href: string, tooltip?: string) =>
      mounted.editor.exec({
        type: 'insertHyperlink',
        href,
        ...(tooltip !== undefined ? { tooltip } : {}),
      }).ok;
    expect(retarget('https://example.org', 'Moved')).toBe(true);
    expect(mounted.editor.query({ type: 'hyperlinkAt' })?.tooltip).toBe('Moved');
    expect(retarget('https://example.net')).toBe(true);
    expect(mounted.editor.query({ type: 'hyperlinkAt' })).toMatchObject({
      href: 'https://example.net',
      tooltip: 'Moved',
    });
  });

  test('an empty or unwritable tooltip is refused before the target is written', () => {
    const mounted = mount(p('Visit example today'));
    select(mounted, 0, 6, 13);
    for (const tooltip of ['', 'bell\u0007']) {
      const result = mounted.editor.exec({
        type: 'insertHyperlink',
        href: 'https://example.com',
        tooltip,
      });
      expect(result.ok === false && result.code).toBe('invalidArgs');
    }
    expect(mounted.editor.surface!.hyperlinks.linksInCaretParagraph()).toEqual([]);
    expect(mounted.editor.surface!.session.currentPackage().externalTargets).toEqual([]);
  });
});

describe('removeHyperlink through the command', () => {
  test('reports a refused removal as refused, not as a missing link', () => {
    const container = document.createElement('div');
    document.body.append(container);
    const twoFormats =
      '<w:hyperlink w:anchor="old"><w:r><w:t>Do</w:t></w:r>' +
      '<w:r><w:rPr><w:b/></w:rPr><w:t>cs</w:t></w:r></w:hyperlink>';
    const editor = createDocxEditor({
      container,
      document: docx(`<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>${twoFormats}</w:p>`),
      author: 'Ada Lovelace',
      modules: [stubReviewModule()],
    });
    const mounted = { editor, container };
    expect(editor.setEditingMode('suggesting').ok).toBe(true);
    caret(mounted, 0, 'See D'.length);
    expect(editor.exec({ type: 'removeHyperlink' })).toMatchObject({
      ok: false,
      code: 'invalidArgs',
      reason: expect.stringContaining('plain text in one format'),
    });
    // The refusal above must not make a caret outside any link read as refused too.
    caret(mounted, 0, 1);
    expect(editor.exec({ type: 'removeHyperlink' })).toMatchObject({
      ok: false,
      code: 'notFound',
    });
  });
});

/** Word's default character style, as a styles part from Word declares it. */
const STYLES_WITH_DEFAULT_FONT =
  `<w:styles xmlns:w="${W}">` +
  '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">' +
  '<w:name w:val="Default Paragraph Font"/></w:style></w:styles>';

/** Link `[start, end)` of the first paragraph. */
function link(mounted: Mounted, start: number, end: number, href: string): void {
  select(mounted, 0, start, end);
  expect(mounted.editor.exec({ type: 'insertHyperlink', href }).ok).toBe(true);
}

/** The `Hyperlink` style definitions in the styles part, and the run styles in the body. */
function hyperlinkStyles(mounted: Mounted): { definitions: string[]; runStyles: string[] } {
  const session = mounted.editor.surface!.session;
  const styles = stylesPartOf(session.currentPackage());
  return {
    definitions:
      (styles ? serializeOoxmlPart(styles) : '').match(
        /<w:style [^>]*w:styleId="Hyperlink".*?<\/w:style>/g
      ) ?? [],
    runStyles: [...serializeOoxmlPart(session.part()).matchAll(/<w:rStyle w:val="([^"]+)"/g)].map(
      (match) => match[1]!
    ),
  };
}

// Word writes a built-in style into `styles.xml` only once something uses it, so a document
// commonly reaches its first link with no `Hyperlink` style. Word adds the definition then.
describe('a new link in a document without the Hyperlink style', () => {
  test('adds the definition Word writes, once, in the same undo step', () => {
    const mounted = mount(p('Visit example today'));
    link(mounted, 6, 13, 'https://example.com');
    const [definition] = hyperlinkStyles(mounted).definitions;
    expect(definition).toContain('w:type="character"');
    expect(definition).toContain('<w:name w:val="Hyperlink"/>');
    expect(definition).toContain(
      '<w:rPr><w:color w:themeColor="hyperlink" w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
    );
    link(mounted, 0, 5, 'https://example.org');
    expect(hyperlinkStyles(mounted).definitions).toHaveLength(1);
    expect(hyperlinkStyles(mounted).runStyles).toEqual(['Hyperlink', 'Hyperlink']);
    // Undo takes the first link and the definition it brought back together.
    mounted.editor.exec({ type: 'undo' });
    mounted.editor.exec({ type: 'undo' });
    expect(hyperlinkStyles(mounted)).toEqual({ definitions: [], runStyles: [] });
  });

  test('bases it on the default character style and saves it', () => {
    const mounted = mount(p('Visit example today'), STYLES_WITH_DEFAULT_FONT);
    link(mounted, 6, 13, 'https://example.com');
    const saved = unzipSync(mounted.editor.surface!.session.save());
    const styles = new TextDecoder().decode(saved['word/styles.xml']!);
    expect(styles).toMatch(/w:styleId="Hyperlink".*<w:basedOn w:val="DefaultParagraphFont"\/>/);
    expect(new TextDecoder().decode(saved['word/document.xml']!)).toContain(
      '<w:rStyle w:val="Hyperlink"/>'
    );
  });

  test('finds a localized one by name, and writes under a free id when its id is taken', () => {
    // Word localizes a built-in's id but not its name: a French document calls it
    // `Lienhypertexte`. An id another style already holds gets a suffix, as List Paragraph does.
    const localized = mount(
      p('Visit example today'),
      `<w:styles xmlns:w="${W}"><w:style w:type="character" w:styleId="Lienhypertexte">` +
        '<w:name w:val="Hyperlink"/></w:style></w:styles>'
    );
    link(localized, 6, 13, 'https://example.com');
    expect(hyperlinkStyles(localized)).toEqual({ definitions: [], runStyles: ['Lienhypertexte'] });

    const taken = mount(
      p('Visit example today'),
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Hyperlink">` +
        '<w:name w:val="Links"/></w:style></w:styles>'
    );
    link(taken, 6, 13, 'https://example.com');
    expect(hyperlinkStyles(taken).runStyles).toEqual(['Hyperlink1']);
    expect(
      taken.editor.surface!.session.documentStyles().find((style) => style.styleId === 'Hyperlink1')
    ).toMatchObject({ type: 'character', name: 'Hyperlink' });
  });
});
