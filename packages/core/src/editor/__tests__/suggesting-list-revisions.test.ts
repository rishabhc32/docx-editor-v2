// Suggesting mode records the toolbar's list buttons as a `w:pPrChange`.
//
// Bullet List, Numbered List and the indent buttons on a list item emit `setListNumbering` and
// `setListLevel`. The store already records both as a paragraph property change, but the
// surface never attributed them, so a suggester's list press applied untracked: no card,
// nothing to reject, and the list stayed. The #458 failure, for lists.
//
// It hid behind List Paragraph. Bullet List on a plain paragraph also writes `w:pStyle`, and
// that write WAS attributed, so its record happened to hold the paragraph before the list too.
// Every press that writes no style had nothing: a document with no List Paragraph style, a
// conversion between bullets and numbers, a list turned off, and an indent.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import type { OoxmlNode } from '@docx-editor.dev/core/store';
import { revisionItemsOf } from '../../store/store/review-reads.ts';
import { planRevisionBatch } from '../../store/store/revision-batch.ts';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import type { PaginatedSurface } from '../paginated-surface.ts';
import { runToolbarCommand } from '../toolbar-commands.ts';
import { stubReviewModule } from './review-test-module.ts';

const AUTHOR = 'Bea';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

const NORMAL =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
// List Paragraph already states `w:contextualSpacing`, so the list press writes no style repair.
const LIST_PARAGRAPH =
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>';
const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="0">` +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';

const PLAIN = '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>';
const LIST_ITEM =
  '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>' +
  '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
  '<w:r><w:t>alpha</w:t></w:r></w:p>';

function docx(
  body: string,
  { styles = NORMAL + LIST_PARAGRAPH, settings = '' }: { styles?: string; settings?: string } = {}
): Uint8Array {
  const type = (name: string) =>
    `application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${type('document.main')}"/>` +
        `<Override PartName="/word/styles.xml" ContentType="${type('styles')}"/>` +
        `<Override PartName="/word/numbering.xml" ContentType="${type('numbering')}"/>` +
        `<Override PartName="/word/settings.xml" ContentType="${type('settings')}"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdS" Type="${R}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rIdN" Type="${R}/numbering" Target="numbering.xml"/>` +
        `<Relationship Id="rIdT" Type="${R}/settings" Target="settings.xml"/></Relationships>`
    ),
    'word/styles.xml': strToU8(`<w:styles xmlns:w="${W}">${styles}</w:styles>`),
    'word/numbering.xml': strToU8(NUMBERING),
    'word/settings.xml': strToU8(`<w:settings xmlns:w="${W}">${settings}</w:settings>`),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

/** Torn down after each test: a container left on `document` leaks into the serial run. */
const mounted: { editor: DocxEditorInstance; container: HTMLElement }[] = [];

afterEach(() => {
  for (const entry of mounted.splice(0)) {
    entry.editor.destroy();
    entry.container.remove();
  }
});

function suggesting(bytes: Uint8Array): { editor: DocxEditorInstance; surface: PaginatedSurface } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    author: AUTHOR,
    mode: 'suggesting',
    modules: [stubReviewModule()],
  });
  mounted.push({ editor, container });
  if (!editor.surface) throw new Error('surface failed to mount');
  const paragraphId = editor.surface.session.paragraphIds()[0]!;
  editor.surface.setSelection({
    anchor: { paragraphId, offset: 2 },
    head: { paragraphId, offset: 2 },
  });
  return { editor, surface: editor.surface };
}

function findAll(node: OoxmlNode, localName: string): OoxmlNode[] {
  const found: OoxmlNode[] = [];
  const walk = (current: OoxmlNode): void => {
    if (current.kind === 'textValue') return;
    if (current.localName === localName) found.push(current);
    for (const child of current.children) walk(child);
  };
  walk(node);
  return found;
}

function attributeOf(node: OoxmlNode | undefined, localName: string): string | undefined {
  if (!node || node.kind === 'textValue') return undefined;
  return node.attributes.find((attribute) => attribute.localName === localName)?.value;
}

function childOf(node: OoxmlNode | undefined, localName: string): OoxmlNode | undefined {
  if (!node || node.kind === 'textValue') return undefined;
  return node.children.find((child) => child.kind !== 'textValue' && child.localName === localName);
}

function childNames(node: OoxmlNode | undefined): string[] {
  if (!node || node.kind === 'textValue') return [];
  return node.children.flatMap((child) => (child.kind === 'textValue' ? [] : [child.localName]));
}

/** The paragraph's live `w:pPr` and every `w:pPrChange` in the body. */
function properties(surface: PaginatedSurface) {
  const root = surface.session.part().root;
  const paragraph = findAll(root, 'p')[0];
  const pPr = childOf(paragraph, 'pPr');
  const changes = findAll(root, 'pPrChange');
  return { pPr, changes, recorded: childOf(changes[0], 'pPr') };
}

/** Reject every decision the way the review pane's Reject All plans it. */
function rejectAll(surface: PaginatedSurface): number {
  const plan = planRevisionBatch(surface.session.part(), 'reject');
  const result = surface.applyAutomationOps(() => [...plan.ops]);
  if (result.rejected) throw new Error(String(result.reason));
  return plan.result.resolved.length;
}

describe('list changes in suggesting mode', () => {
  test('Bullets on a plain paragraph is one suggestion that rejecting takes back', () => {
    const { editor, surface } = suggesting(docx(PLAIN));
    expect(runToolbarCommand(editor, 'list.bullet')).toMatchObject({ ok: true, changed: true });

    const after = properties(surface);
    expect(childNames(after.pPr)).toEqual(['pStyle', 'numPr', 'pPrChange']);
    expect(attributeOf(childOf(after.pPr, 'pStyle'), 'val')).toBe('ListParagraph');
    // ONE record for the style write and the numbering write: the press is one proposal.
    // The transaction shares one `w:pPrChange` id between them.
    expect(after.changes).toHaveLength(1);
    expect(attributeOf(after.changes[0], 'author')).toBe(AUTHOR);
    // The paragraph had no `w:pPr`, so the record is empty: no list, no style.
    expect(childNames(after.recorded)).toEqual([]);
    expect(revisionItemsOf(surface.session.part())).toMatchObject([
      { revisionKind: 'format', author: AUTHOR },
    ]);

    expect(rejectAll(surface)).toBe(1);
    const rejected = properties(surface);
    expect(rejected.changes).toHaveLength(0);
    expect(childNames(rejected.pPr)).not.toContain('numPr');
    expect(childNames(rejected.pPr)).not.toContain('pStyle');
    expect(revisionItemsOf(surface.session.part())).toEqual([]);
    expect(surface.session.bodyText()).toBe('alpha');
  });

  test('Bullets with no List Paragraph style to apply is still one suggestion', () => {
    // The press writes `w:numPr` alone, so nothing else records the paragraph before it.
    const { editor, surface } = suggesting(docx(PLAIN, { styles: NORMAL }));
    expect(runToolbarCommand(editor, 'list.bullet')).toMatchObject({ ok: true, changed: true });

    const after = properties(surface);
    expect(childNames(after.pPr)).toEqual(['numPr', 'pPrChange']);
    expect(after.changes).toHaveLength(1);
    expect(attributeOf(after.changes[0], 'author')).toBe(AUTHOR);
    expect(childNames(after.recorded)).toEqual([]);
    expect(revisionItemsOf(surface.session.part())).toHaveLength(1);

    expect(rejectAll(surface)).toBe(1);
    const rejected = properties(surface);
    expect(rejected.changes).toHaveLength(0);
    expect(childNames(rejected.pPr)).not.toContain('numPr');
  });

  for (const [slot, gesture] of [
    ['list.numbered', 'converts a bulleted item to numbers'],
    ['list.bullet', 'turns a bulleted item off'],
  ] as const) {
    // The item is already in List Paragraph, so the press writes `w:numPr` alone.
    test(`a press that ${gesture} records the list it had`, () => {
      const { editor, surface } = suggesting(docx(LIST_ITEM));
      expect(runToolbarCommand(editor, slot)).toMatchObject({ ok: true, changed: true });

      const after = properties(surface);
      expect(after.changes).toHaveLength(1);
      expect(attributeOf(after.changes[0], 'author')).toBe(AUTHOR);
      expect(childNames(after.recorded)).toEqual(['pStyle', 'numPr']);
      expect(attributeOf(childOf(childOf(after.recorded, 'numPr'), 'numId'), 'val')).toBe('1');

      expect(rejectAll(surface)).toBe(1);
      const rejected = properties(surface);
      expect(rejected.changes).toHaveLength(0);
      expect(attributeOf(childOf(childOf(rejected.pPr, 'numPr'), 'numId'), 'val')).toBe('1');
    });
  }

  test('Increase Indent on a list item records the level it had', () => {
    const { editor, surface } = suggesting(docx(LIST_ITEM));
    expect(runToolbarCommand(editor, 'list.indent')).toMatchObject({ ok: true, changed: true });

    const after = properties(surface);
    expect(attributeOf(childOf(childOf(after.pPr, 'numPr'), 'ilvl'), 'val')).toBe('1');
    expect(after.changes).toHaveLength(1);
    expect(attributeOf(after.changes[0], 'author')).toBe(AUTHOR);
    expect(childNames(after.recorded)).toEqual(['pStyle', 'numPr']);
    expect(attributeOf(childOf(childOf(after.recorded, 'numPr'), 'ilvl'), 'val')).toBe('0');
    expect(attributeOf(childOf(childOf(after.recorded, 'numPr'), 'numId'), 'val')).toBe('1');

    expect(rejectAll(surface)).toBe(1);
    const rejected = properties(surface);
    expect(rejected.changes).toHaveLength(0);
    expect(attributeOf(childOf(childOf(rejected.pPr, 'numPr'), 'ilvl'), 'val')).toBe('0');
  });

  test('a document that turns off formatting tracking takes the list untracked', () => {
    // `w:doNotTrackFormatting` silences every `w:pPrChange` writer, and a list press writes
    // one, so it follows the same switch as the other paragraph property writes.
    const { editor, surface } = suggesting(docx(PLAIN, { settings: '<w:doNotTrackFormatting/>' }));
    expect(runToolbarCommand(editor, 'list.bullet')).toMatchObject({ ok: true, changed: true });

    const after = properties(surface);
    expect(childNames(after.pPr)).toEqual(['pStyle', 'numPr']);
    expect(after.changes).toHaveLength(0);
    expect(revisionItemsOf(surface.session.part())).toEqual([]);
  });
});
