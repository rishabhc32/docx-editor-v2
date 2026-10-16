// REPEAT HEADER ROWS AND TABLE ALIGNMENT.
//
// Both were refused in every mode as "not supported by the tree editor". Repeat Header Rows
// now follows Word's rule that header rows are the table's leading rows, and the header rows
// repeat atop each page the table continues onto. Table alignment moves the painted table, and
// takes left and right as the reader sees them in a right-to-left table too; a preferred width
// is written in its own unit. In suggesting mode each is one card that Accept and Reject answer
// whole.

import { GlobalRegistrator } from '@happy-dom/global-registrator';
if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

import { afterEach, describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { createDocxEditor, type DocxEditorInstance } from '../docx-editor.ts';
import {
  collectReviewItems as engineCollectReviewItems,
  findNode,
  revisionItemsOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import type { EditorCommand } from '../../contracts/editor.ts';
import type { EditorModule } from '../../contracts/modules.ts';
import { resolveRevisionMarkup } from '../../contracts/revision-markup.ts';
import type { TableFragmentRecord } from '../../layout/semantic-records.ts';
import {
  FLOATING_TABLE_ALIGNMENT_REASON,
  HEADER_ROW_FROM_STYLE_REASON,
} from '../table-property-commands.ts';
import { selectCellRectangle } from './paginated-surface-fixtures.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
const row = (t: string, properties = '') =>
  `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ''}${cell(`${t}1`)}${cell(`${t}2`)}</w:tr>`;
const HEADER = '<w:tblHeader/>';
/** A two-column table two inches wide, so alignment has room to move it. */
const table = (rows: readonly string[], properties = '') =>
  `<w:tbl><w:tblPr>${properties}<w:tblW w:w="2880" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
  '<w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>' +
  `${rows.join('')}</w:tbl>`;
const rowsNamed = (count: number, header = 0): string[] =>
  Array.from({ length: count }, (_, index) =>
    row(String.fromCharCode(97 + (index % 26)), index < header ? HEADER : '')
  );
const BODY = p('before') + table(rowsNamed(4)) + p('after');
/** Paragraph ordinals: `before` is 0, then two per row. */
const firstCellOfRow = (index: number) => 1 + index * 2;

const STYLE_HEADER =
  '<w:style w:type="table" w:styleId="Repeating"><w:name w:val="Repeating"/>' +
  `<w:trPr>${HEADER}</w:trPr></w:style>`;

function docxOf(
  body: string,
  options: { doNotTrackFormatting?: boolean; styles?: string } = {}
): Uint8Array {
  const parts: Record<string, string> = {};
  const overrides: string[] = [];
  const relationships: string[] = [];
  if (options.doNotTrackFormatting === true) {
    parts['word/settings.xml'] =
      `<w:settings xmlns:w="${W}"><w:doNotTrackFormatting/></w:settings>`;
    overrides.push('settings');
    relationships.push(`<Relationship Id="rId2" Type="${OD}/settings" Target="settings.xml"/>`);
  }
  if (options.styles !== undefined) {
    parts['word/styles.xml'] = `<w:styles xmlns:w="${W}">${options.styles}</w:styles>`;
    overrides.push('styles');
    relationships.push(`<Relationship Id="rId3" Type="${OD}/styles" Target="styles.xml"/>`);
  }
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        overrides
          .map(
            (name) =>
              `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`
          )
          .join('') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">${relationships.join('')}</Relationships>`
    ),
    ...Object.fromEntries(Object.entries(parts).map(([name, xml]) => [name, strToU8(xml)])),
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
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
});

function mount(document: Uint8Array, mode: 'editing' | 'suggesting' = 'editing') {
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document,
    author: 'Grace Hopper',
    modules: [reviewModule()],
  });
  if (!editor.surface) throw new Error('surface failed to mount');
  editors.push(editor);
  expect(editor.exec({ type: 'setEditingMode', mode }).ok).toBe(true);
  return editor;
}

function caret(editor: DocxEditorInstance, paragraph: number): void {
  const position = { paragraphId: editor.surface!.session.paragraphIds()[paragraph]!, offset: 0 };
  editor.surface!.setSelection({ anchor: position, head: position });
}

const xmlOf = (editor: DocxEditorInstance): string =>
  serializeOoxmlPart(editor.surface!.session.part());
const cards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind));
/** Which rows carry `w:tblHeader`, in table order. */
const headerFlags = (editor: DocxEditorInstance): boolean[] =>
  (xmlOf(editor).match(/<w:tr>.*?<\/w:tr>/g) ?? []).map((xml) =>
    /<w:trPr>(?:(?!<w:trPrChange).)*<w:tblHeader w:val="1"\/>/.test(xml)
  );
const tblPr = (editor: DocxEditorInstance): string =>
  xmlOf(editor).match(/<w:tblPr>.*?<\/w:tblPr>(?=<w:tblGrid>)/)![0];

function tables(editor: DocxEditorInstance): TableFragmentRecord[] {
  return editor
    .surface!.layout()
    .pages.flatMap((page) => page.fragments)
    .filter((fragment): fragment is TableFragmentRecord => fragment.kind === 'table');
}
const tableX = (editor: DocxEditorInstance): number => tables(editor)[0]!.box.x;

const exec = (editor: DocxEditorInstance, command: EditorCommand) =>
  expect(editor.exec(command)).toEqual({ ok: true, changed: true });
const align = (justification: 'left' | 'center' | 'right' | null): EditorCommand => ({
  type: 'setTableProperties',
  justification,
});

describe('toggleHeaderRow', () => {
  test('turns on every row from the first through the selected one', () => {
    const editor = mount(docxOf(BODY));
    caret(editor, firstCellOfRow(2));
    expect(editor.can({ type: 'toggleHeaderRow' })).toEqual({ ok: true });
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(false);
    exec(editor, { type: 'toggleHeaderRow' });
    expect(headerFlags(editor)).toEqual([true, true, true, false]);
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(true);
  });

  test('turning it off keeps the header rows above the selection', () => {
    const editor = mount(docxOf(p('before') + table(rowsNamed(4, 3)) + p('after')));
    caret(editor, firstCellOfRow(1));
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(true);
    exec(editor, { type: 'toggleHeaderRow' });
    expect(headerFlags(editor)).toEqual([true, false, false, false]);
    caret(editor, firstCellOfRow(0));
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(true);
    exec(editor, { type: 'toggleHeaderRow' });
    expect(headerFlags(editor)).toEqual([false, false, false, false]);
  });

  test('a flag below a plain row repeats nothing, so the row reads off and turns on', () => {
    const rows = rowsNamed(4);
    rows[2] = row('c', HEADER);
    const editor = mount(docxOf(p('before') + table(rows) + p('after')));
    caret(editor, firstCellOfRow(2));
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(false);
    exec(editor, { type: 'toggleHeaderRow' });
    expect(headerFlags(editor)).toEqual([true, true, true, false]);
  });

  test('a cell selection over several rows reads and toggles all of them', () => {
    const editor = mount(docxOf(p('before') + table(rowsNamed(4, 1)) + p('after')));
    selectCellRectangle(editor.surface!, { row: 0, column: 0 }, { row: 1, column: 1 });
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(false);
    exec(editor, { type: 'toggleHeaderRow' });
    expect(headerFlags(editor)).toEqual([true, true, false, false]);
    selectCellRectangle(editor.surface!, { row: 0, column: 0 }, { row: 1, column: 1 });
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(true);
  });

  test('the header rows repeat atop each page the table continues onto', () => {
    const editor = mount(docxOf(p('before') + table(rowsNamed(90)) + p('after')));
    expect(tables(editor).length).toBeGreaterThan(2);
    expect(tables(editor)[1]!.rows[0]!.isHeaderRepeat).toBe(false);
    caret(editor, firstCellOfRow(1));
    exec(editor, { type: 'toggleHeaderRow' });
    for (const { rows } of tables(editor).slice(1)) {
      expect(rows.slice(0, 2).map((entry) => [entry.rowIndex, entry.isHeaderRepeat])).toEqual([
        [0, true],
        [1, true],
      ]);
    }
  });

  test('a header row the table style makes is refused rather than left on', () => {
    const style = '<w:tblStyle w:val="Repeating"/>';
    const editor = mount(
      docxOf(p('before') + table(rowsNamed(4), style) + p('after'), { styles: STYLE_HEADER })
    );
    caret(editor, firstCellOfRow(0));
    expect(editor.isActive({ type: 'toggleHeaderRow' })).toBe(true);
    expect(editor.can({ type: 'toggleHeaderRow' })).toEqual({
      ok: false,
      code: 'unsupported',
      reason: HEADER_ROW_FROM_STYLE_REASON,
    });
  });

  test('is refused outside a table and with arguments', () => {
    const editor = mount(docxOf(BODY));
    caret(editor, 0);
    expect(editor.can({ type: 'toggleHeaderRow' }).ok).toBe(false);
    caret(editor, firstCellOfRow(0));
    const extra = { type: 'toggleHeaderRow', rows: 2 } as unknown as EditorCommand;
    expect(editor.can(extra)).toMatchObject({ ok: false, code: 'invalidArgs' });
  });
});

describe('setTableProperties', () => {
  test('alignment writes w:jc in schema order and moves the painted table', () => {
    const editor = mount(docxOf(BODY));
    caret(editor, firstCellOfRow(0));
    const left = tableX(editor);
    exec(editor, align('center'));
    expect(tblPr(editor)).toBe(
      '<w:tblPr><w:tblW w:type="dxa" w:w="2880"/><w:jc w:val="center"/>' +
        '<w:tblLayout w:type="fixed"/></w:tblPr>'
    );
    const center = tableX(editor);
    exec(editor, align('right'));
    const right = tableX(editor);
    // Centred halfway between the margins, within the half point an outer rule moves it.
    expect(center).toBeGreaterThan(left + 100);
    expect(Math.abs(right - center - (center - left))).toBeLessThanOrEqual(1);
    exec(editor, align(null));
    expect(tblPr(editor)).not.toContain('<w:jc');
    expect(tableX(editor)).toBe(left);
  });

  test('a right-to-left table takes left and right as the reader sees them', () => {
    const ltr = mount(docxOf(BODY));
    caret(ltr, firstCellOfRow(0));
    const pageLeft = tableX(ltr);
    exec(ltr, align('right'));
    const pageRight = tableX(ltr);
    const rtl = mount(docxOf(p('before') + table(rowsNamed(4), '<w:bidiVisual/>') + p('after')));
    caret(rtl, firstCellOfRow(0));
    exec(rtl, align('left'));
    expect(tblPr(rtl)).toContain('<w:jc w:val="right"/>');
    expect(Math.abs(tableX(rtl) - pageLeft)).toBeLessThanOrEqual(1);
    exec(rtl, align('right'));
    expect(tblPr(rtl)).toContain('<w:jc w:val="left"/>');
    expect(Math.abs(tableX(rtl) - pageRight)).toBeLessThanOrEqual(1);
  });

  test('a table that text wraps around refuses alignment, which would not move it', () => {
    const floating = '<w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="200"/>';
    const editor = mount(docxOf(p('before') + table(rowsNamed(4), floating) + p('after')));
    caret(editor, firstCellOfRow(0));
    expect(editor.can(align('center'))).toEqual({
      ok: false,
      code: 'unsupported',
      reason: FLOATING_TABLE_ALIGNMENT_REASON,
    });
    exec(editor, { type: 'setTableProperties', width: 2500, widthType: 'pct' });
    expect(tblPr(editor)).toContain('<w:tblW w:type="pct" w:w="2500"/>');
  });

  test('a preferred width writes w:tblW in its unit, and null removes it', () => {
    const editor = mount(docxOf(BODY));
    caret(editor, firstCellOfRow(0));
    const widths: readonly (readonly [EditorCommand, string | null])[] = [
      [{ type: 'setTableProperties', width: 4320, widthType: 'dxa' }, 'w:type="dxa" w:w="4320"'],
      [{ type: 'setTableProperties', width: 2500, widthType: 'pct' }, 'w:type="pct" w:w="2500"'],
      [{ type: 'setTableProperties', width: null, widthType: 'auto' }, 'w:type="auto" w:w="0"'],
      [{ type: 'setTableProperties', widthType: null }, null],
    ];
    for (const [command, attributes] of widths) {
      exec(editor, command);
      expect(tblPr(editor)).toBe(
        `<w:tblPr>${attributes ? `<w:tblW ${attributes}/>` : ''}<w:tblLayout w:type="fixed"/></w:tblPr>`
      );
    }
  });

  test('malformed or out-of-range properties are refused', () => {
    const editor = mount(docxOf(BODY));
    caret(editor, firstCellOfRow(0));
    const refused = [
      {},
      { justification: 'justify' },
      { width: 1440 },
      { width: 1440, widthType: 'auto' },
      { widthType: 'dxa' },
      { width: 99_999, widthType: 'dxa' },
      { width: 6000, widthType: 'pct' },
      { width: 14.5, widthType: 'dxa' },
      { justification: 'left', styleId: 'x' },
    ];
    for (const fields of refused) {
      const command = { type: 'setTableProperties', ...fields } as unknown as EditorCommand;
      expect(editor.can(command)).toMatchObject({ ok: false, code: 'invalidArgs' });
    }
    caret(editor, 0);
    expect(editor.can(align('center')).ok).toBe(false);
  });
});

interface Gesture {
  readonly name: string;
  readonly body: string;
  readonly run: (editor: DocxEditorInstance) => void;
  /** Whether the document shows the gesture's effect, accepted or applied outright. */
  readonly applied: (editor: DocxEditorInstance) => boolean;
}

const GESTURES: readonly Gesture[] = [
  {
    name: 'repeat header rows',
    body: BODY,
    run: (editor) => {
      caret(editor, firstCellOfRow(1));
      exec(editor, { type: 'toggleHeaderRow' });
    },
    applied: (editor) => headerFlags(editor).join() === 'true,true,false,false',
  },
  {
    name: 'repeat header rows off',
    body: p('before') + table(rowsNamed(4, 2)) + p('after'),
    run: (editor) => {
      caret(editor, firstCellOfRow(0));
      exec(editor, { type: 'toggleHeaderRow' });
    },
    applied: (editor) => !headerFlags(editor).includes(true),
  },
  {
    name: 'table alignment',
    body: BODY,
    run: (editor) => {
      caret(editor, firstCellOfRow(0));
      exec(editor, align('center'));
    },
    applied: (editor) => tblPr(editor).includes('<w:jc w:val="center"/>'),
  },
  {
    name: 'table width',
    body: BODY,
    run: (editor) => {
      caret(editor, firstCellOfRow(0));
      exec(editor, { type: 'setTableProperties', width: 2500, widthType: 'pct' });
    },
    applied: (editor) => tblPr(editor).includes('<w:tblW w:type="pct" w:w="2500"/>'),
  },
];

const RECORDS = /<w:(trPrChange|tcPrChange|tblPrChange|tblGridChange)\b/;

describe.each(GESTURES)('$name', (gesture) => {
  test('suggesting leaves one card, and Accept applies it as editing mode does', () => {
    const editor = mount(docxOf(gesture.body), 'suggesting');
    gesture.run(editor);
    expect(cards(editor)).toEqual(['format']);
    expect(gesture.applied(editor)).toBe(true);
    // The proposal paints as the change would: the record holding the old properties does not.
    const editing = mount(docxOf(gesture.body));
    gesture.run(editing);
    expect(tables(editor).map((fragment) => fragment.box)).toEqual(
      tables(editing).map((fragment) => fragment.box)
    );
    expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(xmlOf(editor)).toBe(xmlOf(editing));
  });

  test('Reject restores the document, every other row and cell included', () => {
    const editor = mount(docxOf(gesture.body), 'suggesting');
    const original = xmlOf(editor);
    gesture.run(editor);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    // A row with no `w:trPr` keeps the empty one its record needed; nothing else differs.
    expect(xmlOf(editor).replaceAll('<w:tr><w:trPr/>', '<w:tr>')).toBe(original);
  });

  test('one undo step takes the proposal back, and redo brings it again', () => {
    const editor = mount(docxOf(gesture.body), 'suggesting');
    const original = xmlOf(editor);
    gesture.run(editor);
    const proposed = xmlOf(editor);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(original);
    expect(cards(editor)).toEqual([]);
    expect(editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(proposed);
  });

  test('a save keeps the proposal, and reopening reads the same card', async () => {
    const editor = mount(docxOf(gesture.body), 'suggesting');
    gesture.run(editor);
    const reopened = mount(new Uint8Array(await editor.save()));
    expect(cards(reopened)).toEqual(['format']);
    expect(xmlOf(reopened)).toBe(xmlOf(editor));
  });

  test('editing mode applies it outright, with no record', () => {
    const editing = mount(docxOf(gesture.body));
    gesture.run(editing);
    expect(cards(editing)).toEqual([]);
    expect(gesture.applied(editing)).toBe(true);
    expect(xmlOf(editing)).not.toMatch(RECORDS);
  });

  test('with formatting changes untracked, it applies bare where it is otherwise proposed', () => {
    const muted = mount(docxOf(gesture.body), 'suggesting');
    muted.surface!.setRevisionMarkup(resolveRevisionMarkup({ trackFormatting: false }));
    gesture.run(muted);
    const document = mount(docxOf(gesture.body, { doNotTrackFormatting: true }), 'suggesting');
    gesture.run(document);
    for (const editor of [muted, document]) {
      expect(cards(editor)).toEqual([]);
      expect(gesture.applied(editor)).toBe(true);
      expect(xmlOf(editor)).not.toMatch(RECORDS);
    }
  });
});
