// SUGGESTING MODE PROPOSES ROW HEIGHT AND COLUMN WIDTH.
//
// Dragging a row divider, a column divider or the table's right edge changed the document
// outright in suggesting mode: no review card, nothing to reject. Each now leaves one card that
// Accept and Reject answer, one undo step, and markup a save keeps.

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
import type { EditorModule } from '../../contracts/modules.ts';
import { resolveRevisionMarkup } from '../../contracts/revision-markup.ts';
import type { ExecResult } from '../../contracts/types.ts';
import { HELD_ROW_REASON } from '../surface-tracked-table-refusal.ts';
import { planTableRowHeightResize } from '../table-command-plan.ts';
import {
  findTableOccurrence,
  tableColumnDividerResizeTargetFrom,
  tableRightEdgeResizeTargetFrom,
  tableRowOccurrenceTargetFrom,
  type TableOccurrenceRef,
} from '../../layout/table-interaction-targets.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}</w:tr><w:tr>${cell('b1')}${cell('b2')}</w:tr></w:tbl>` +
  p('after');

/** The document, with a `settings.xml` asking for formatting changes NOT to be recorded. */
function docxOf(body: string, options: { doNotTrackFormatting?: boolean } = {}): Uint8Array {
  const settings = options.doNotTrackFormatting === true;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        (settings
          ? '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
          : '') +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${OD}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
    ...(settings
      ? {
          'word/_rels/document.xml.rels': strToU8(
            `<Relationships xmlns="${REL}"><Relationship Id="rId2" Type="${OD}/settings" Target="settings.xml"/></Relationships>`
          ),
          'word/settings.xml': strToU8(
            `<w:settings xmlns:w="${W}"><w:doNotTrackFormatting/></w:settings>`
          ),
        }
      : {}),
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

function mount(document: Uint8Array, mode: 'editing' | 'suggesting'): DocxEditorInstance {
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

const xmlOf = (editor: DocxEditorInstance): string =>
  serializeOoxmlPart(editor.surface!.session.part());
const cards = (editor: DocxEditorInstance) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind));

/** One row's occurrence of the table, as a pointer gesture targets it. */
function rowAt(editor: DocxEditorInstance, index: number): TableOccurrenceRef {
  const layout = editor.surface!.layout();
  const table = layout.pages[0]!.fragments.find((fragment) => fragment.kind === 'table');
  if (table?.kind !== 'table') throw new Error('no table on the first page');
  const ref = findTableOccurrence(layout, table.tableId, table.rows[index]!.id, false);
  if (!ref) throw new Error('row occurrence missing');
  return ref;
}

/** The table as it paints: column edges and row heights. */
function geometry(editor: DocxEditorInstance): unknown {
  const table = rowAt(editor, 0).table;
  return { columns: table.columnEdges, rows: table.rows.map((row) => row.box.height) };
}

/** Drag the divider under row `index` to `heightTwips`. */
function dragRowHeight(editor: DocxEditorInstance, index: number, heightTwips: number): ExecResult {
  const surface = editor.surface!;
  const layout = surface.layout();
  const state = surface.state();
  const plan = planTableRowHeightResize(
    {
      part: surface.session.part(),
      layout,
      storeRevision: surface.session.packageRevision(),
      selection: state.selection,
      cellSelection: state.cellSelection,
      themeColors: surface.session.documentThemeColors(),
      editable: true,
      viewing: false,
    },
    tableRowOccurrenceTargetFrom(layout.revision, rowAt(editor, index)),
    heightTwips
  );
  return surface.applyTableCommandPlan(plan);
}

/** Drag the divider between the two columns. */
function dragColumnDivider(editor: DocxEditorInstance): ExecResult {
  const ref = rowAt(editor, 0);
  const [left, right] = ref.row.cells.map((entry) => entry.gridColumnId!);
  const target = tableColumnDividerResizeTargetFrom(
    editor.surface!.layout().revision,
    ref,
    left!,
    right!
  );
  return editor.exec({
    type: 'commitTableColumnDividerResize',
    target,
    leftWidthTwips: 4000,
    rightWidthTwips: 2000,
  });
}

/** Drag the table's right edge. */
function dragRightEdge(editor: DocxEditorInstance): ExecResult {
  const ref = rowAt(editor, 0);
  const target = tableRightEdgeResizeTargetFrom(
    editor.surface!.layout().revision,
    ref,
    ref.row.cells.at(-1)!.gridColumnId!
  );
  return editor.exec({
    type: 'commitTableRightEdgeResize',
    target,
    columnWidthTwips: 4000,
    tableWidthTwips: 7000,
  });
}

const LANDED: ExecResult = { ok: true, changed: true };

interface Case {
  readonly name: string;
  readonly run: (editor: DocxEditorInstance) => void;
  /** Whether the document shows the gesture's effect, accepted or applied outright. */
  readonly applied: (xml: string) => boolean;
}

const CASES: readonly Case[] = [
  {
    name: 'row height',
    run: (editor) => expect(dragRowHeight(editor, 0, 900)).toEqual(LANDED),
    applied: (xml) => xml.includes('<w:trHeight w:hRule="exact" w:val="900"/>'),
  },
  {
    name: 'column divider',
    run: (editor) => expect(dragColumnDivider(editor)).toEqual(LANDED),
    applied: (xml) => xml.includes('<w:gridCol w:w="4000"/><w:gridCol w:w="2000"/>'),
  },
  {
    name: 'right edge',
    run: (editor) => expect(dragRightEdge(editor)).toEqual(LANDED),
    applied: (xml) => xml.includes('<w:tblW w:type="dxa" w:w="7000"/>'),
  },
];

const RECORDS = /<w:(trPrChange|tcPrChange|tblPrChange|tblGridChange)\b/;

describe.each(CASES)('$name', (entry) => {
  test('suggesting leaves one card, and Accept applies it as editing mode does', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    entry.run(editor);
    expect(cards(editor)).toEqual(['format']);
    expect(entry.applied(xmlOf(editor))).toBe(true);
    // The proposal shows the new geometry, and the record holding the old one does not paint.
    const editing = mount(docxOf(BODY), 'editing');
    entry.run(editing);
    expect(geometry(editor)).toEqual(geometry(editing));
    expect(editor.acceptReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    expect(xmlOf(editor)).toBe(xmlOf(editing));
  });

  test('Reject restores the document', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    const original = xmlOf(editor);
    entry.run(editor);
    expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
    expect(cards(editor)).toEqual([]);
    // A row with no `w:trPr` keeps the empty one its record needed, as a rejected row
    // deletion does; nothing else differs.
    expect(xmlOf(editor).replace('<w:tr><w:trPr/>', '<w:tr>')).toBe(original);
  });

  test('one undo step takes the proposal back', () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    const original = xmlOf(editor);
    entry.run(editor);
    expect(cards(editor)).toEqual(['format']);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(original);
    expect(cards(editor)).toEqual([]);
  });

  test('a save keeps the proposal, and reopening reads the same card', async () => {
    const editor = mount(docxOf(BODY), 'suggesting');
    entry.run(editor);
    const reopened = mount(new Uint8Array(await editor.save()), 'editing');
    expect(cards(reopened)).toEqual(['format']);
    expect(xmlOf(reopened)).toBe(xmlOf(editor));
  });

  test('editing mode applies it outright, where suggesting proposes it', () => {
    const editing = mount(docxOf(BODY), 'editing');
    entry.run(editing);
    expect(cards(editing)).toEqual([]);
    expect(entry.applied(xmlOf(editing))).toBe(true);
    expect(xmlOf(editing)).not.toMatch(RECORDS);
    const suggesting = mount(docxOf(BODY), 'suggesting');
    entry.run(suggesting);
    expect(cards(suggesting)).toEqual(['format']);
  });

  test('with formatting changes untracked, it applies bare where it is otherwise proposed', () => {
    const tracked = mount(docxOf(BODY), 'suggesting');
    entry.run(tracked);
    expect(cards(tracked)).toEqual(['format']);
    const muted = mount(docxOf(BODY), 'suggesting');
    muted.surface!.setRevisionMarkup(resolveRevisionMarkup({ trackFormatting: false }));
    entry.run(muted);
    const document = mount(docxOf(BODY, { doNotTrackFormatting: true }), 'suggesting');
    expect(document.surface!.session.trackingSettings().doNotTrackFormatting).toBe(true);
    entry.run(document);
    for (const editor of [muted, document]) {
      expect(cards(editor)).toEqual([]);
      expect(entry.applied(xmlOf(editor))).toBe(true);
      expect(xmlOf(editor)).not.toMatch(RECORDS);
    }
  });
});

// A ROW HOLDS ONE AUTHOR'S PENDING DECISION AT A TIME. Bob proposes a fill in the first row,
// which records on each of its cells. Bea's change to that row used to take some of his records
// over, and rejecting hers and then his reset those cells to `tcW auto`. It is refused instead.
describe("a row with another author's pending fill", () => {
  const FILL = { type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } } as const;
  const HELD: ExecResult = { ok: false, code: 'unsupported', reason: HELD_ROW_REASON };
  // Paragraphs in order: before, a1, a2, b1, b2, after.
  const A1 = 1;
  const B1 = 3;

  function caret(editor: DocxEditorInstance, paragraph: number): void {
    const position = { paragraphId: editor.surface!.session.paragraphIds()[paragraph]!, offset: 0 };
    editor.surface!.setSelection({ anchor: position, head: position });
  }
  const authors = (editor: DocxEditorInstance): unknown[] => [
    ...new Set(editor.getReviewItems().map((item) => item.kind === 'revision' && item.author)),
  ];
  const widths = (editor: DocxEditorInstance): string[] =>
    [...xmlOf(editor).matchAll(/<w:tcW w:type="(\w+)" w:w="(\d+)"\/>/g)].map(
      ([, type, width]) => `${type}:${width}`
    );

  /** Bob's fill pending in the first row, and Bea at the keyboard. */
  function withBobsFill(): DocxEditorInstance {
    const editor = mount(docxOf(BODY), 'suggesting');
    editor.setAuthor('Bob');
    caret(editor, A1);
    expect(editor.exec({ ...FILL, color: { kind: 'hex', value: '00FF00' } })).toEqual(LANDED);
    editor.setAuthor('Bea');
    return editor;
  }

  test("refuses another author's change to it, says why, and takes one to another row", () => {
    const editor = withBobsFill();
    const before = xmlOf(editor);
    caret(editor, A1);
    expect(editor.exec(FILL)).toEqual(HELD);
    expect(dragRowHeight(editor, 0, 900)).toEqual(HELD);
    expect(dragColumnDivider(editor)).toEqual(HELD);
    expect(dragRightEdge(editor)).toEqual(HELD);
    expect(xmlOf(editor)).toBe(before);
    expect(authors(editor)).toEqual(['Bob']);
    caret(editor, B1);
    expect(editor.exec(FILL)).toEqual(LANDED);
    expect(dragRowHeight(editor, 1, 900)).toEqual(LANDED);
    expect(authors(editor)).toEqual(['Bob', 'Bea']);
  });

  test("takes its own author's second change, which Reject takes back with the first", () => {
    const editor = withBobsFill();
    const original = xmlOf(mount(docxOf(BODY), 'editing'));
    expect(dragColumnDivider(editor)).toEqual(HELD);
    editor.setAuthor('Bob');
    expect(dragColumnDivider(editor)).toEqual(LANDED);
    expect(authors(editor)).toEqual(['Bob']);
    expect(editor.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
    expect(xmlOf(editor)).toBe(original);
  });

  for (const action of ['accept', 'reject'] as const) {
    test(`takes the change once the fill is answered (${action}), and no width is lost`, () => {
      const editor = withBobsFill();
      expect(dragColumnDivider(editor)).toEqual(HELD);
      const [fill] = editor.getReviewItems();
      const answered =
        action === 'accept'
          ? editor.acceptReviewItem(fill!.key)
          : editor.rejectReviewItem(fill!.key);
      expect(answered.ok).toBe(true);
      expect(dragColumnDivider(editor)).toEqual(LANDED);
      expect(authors(editor)).toEqual(['Bea']);
      expect(editor.rejectReviewItem(editor.getReviewItems()[0]!.key).ok).toBe(true);
      expect(widths(editor)).toEqual(Array(4).fill('dxa:3000'));
    });
  }
});
