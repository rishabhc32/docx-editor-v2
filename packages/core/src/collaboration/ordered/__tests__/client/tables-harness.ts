// Helpers for the table tests with real editors: a file with a three-column table, the
// editor's own table commands and drags, and what the editors show of the table.

import { expect } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { EditorModule } from '../../../../contracts/modules.ts';
import { planTableCommand } from '../../../../editor/table-command-plan.ts';
import type { OoxmlNode, OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { paragraphTextOf } from '../../../../store/store/tree-op-apply.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { normalizedBody, withoutRevisionIds } from '../matrix/compare.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  editorTexts,
  editorXml,
  sameAsRoom,
  screenFollowsList,
  surface,
  type Person,
} from './editor-harness.ts';

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;
const row = (...texts: string[]) => `<w:tr>${texts.map(cell).join('')}</w:tr>`;

/** A paragraph, a table of three columns 3000 twips wide, and a paragraph. */
export const TABLE_FILE = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    row('Fee', 'Rate', 'Total') +
    row('100', '5', '105') +
    '</w:tbl>' +
    p(t('Signed here.'))
);

/** Run one of the editor's table commands at the caret. */
export function tableCommand(person: Person, command: Readonly<Record<string, unknown>>) {
  return commandOn(surface(person), command);
}

/** Run one of the editor's table commands at an editor's caret (in a room or alone). */
export function commandOn(s: Surface, command: Readonly<Record<string, unknown>>) {
  const plan = planTableCommand({
    // The command shapes are the editor's; the test names them by their type
    command: command as never,
    part: s.session.part(),
    layout: s.layout(),
    storeRevision: s.session.packageRevision(),
    selection: s.state().selection,
    cellSelection: s.state().cellSelection,
    themeColors: s.session.documentThemeColors(),
    editable: s.session.editable,
    viewing: false,
  });
  return s.applyTableCommandPlan(plan);
}

const tc = (text: string, props: string) =>
  `<w:tc><w:tcPr>${props}</w:tcPr>${text ? p(t(text)) : '<w:p/>'}</w:tc>`;
const w = (width: number) => `<w:tcW w:w="${width}" w:type="dxa"/>`;

/**
 * Word's sample of a table with merged cells (Word for Mac, 2026-10-29): a cell over the first
 * two columns of row 1, and column 3 merged down rows 2 and 3.
 */
export const MERGED_FILE = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>' +
    `<w:tr>${tc('Fee and rate', w(6240) + '<w:gridSpan w:val="2"/>')}${tc('Total', w(3120))}</w:tr>` +
    `<w:tr>${tc('100', w(3120))}${tc('5', w(3120))}${tc('105', w(3120) + '<w:vMerge w:val="restart"/>')}</w:tr>` +
    `<w:tr>${tc('200', w(3120))}${tc('6', w(3120))}${tc('', w(3120) + '<w:vMerge/>')}</w:tr>` +
    '</w:tbl>' +
    p(t('Signed here.'))
);

/** Put the caret at the start of the paragraph whose text is `text`. */
export function caretIn(person: Person, text: string, offset = 0): void {
  const index = editorTexts(person).indexOf(text);
  if (index < 0)
    throw new Error(`no paragraph "${text}" in ${JSON.stringify(editorTexts(person))}`);
  caretAt(person, index, offset);
}

/** The editor's XML without the old properties change records hold. */
const shownXml = (person: Person): string =>
  editorXml(person).replace(/<w:(tblGridChange|tcPrChange|tblPrChange)\b.*?<\/w:\1>/g, '');

/** The grid widths the editor shows. */
export const gridOf = (person: Person): number[] =>
  [...shownXml(person).matchAll(/<w:gridCol w:w="(\d+)"/g)].map((m) => Number(m[1]));

/** The cell widths of the editor's first table row. */
export const firstRowWidths = (person: Person): number[] => {
  const first = /<w:tr>(.*?)<\/w:tr>/.exec(shownXml(person))?.[1] ?? '';
  return [...first.matchAll(/<w:tcW w:type="dxa" w:w="(\d+)"/g)].map((m) => Number(m[1]));
};

export function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect({ who: person.name, refusals: person.refusals }).toEqual({
      who: person.name,
      refusals: [],
    });
    expect({ who: person.name, room: sameAsRoom(person, room) }).toEqual({
      who: person.name,
      room: true,
    });
    expect({ who: person.name, screen: screenFollowsList(person) }).toEqual({
      who: person.name,
      screen: true,
    });
  }
  const [first, ...rest] = people;
  if (first) for (const other of rest) expect(editorXml(other)).toBe(editorXml(first));
}

/** A table's node ids in an editor's part. */
export interface TableIds {
  readonly table: string;
  readonly grid: readonly string[];
  readonly rows: readonly string[];
  readonly cells: readonly (readonly string[])[];
}

/** The first table of a part and the ids of its grid columns, rows and cells. */
export function tableIdsIn(root: OoxmlNode): TableIds {
  const stack: OoxmlNode[] = [root];
  for (let node = stack.pop(); node; node = stack.pop()) {
    if (node.kind === 'textValue') continue;
    if (node.localName === 'tbl') {
      const grid: string[] = [];
      const rows: string[] = [];
      const cells: string[][] = [];
      for (const child of node.children) {
        if (child.kind === 'textValue') continue;
        if (child.localName === 'tblGrid')
          for (const column of child.children) grid.push(column.id);
        if (child.localName !== 'tr') continue;
        rows.push(child.id);
        cells.push(
          child.children
            .filter((c) => c.kind !== 'textValue' && c.localName === 'tc')
            .map((c) => c.id)
        );
      }
      return { table: node.id, grid, rows, cells };
    }
    for (const child of [...node.children].reverse()) stack.push(child);
  }
  throw new Error('no table');
}

/** The editor's first table and its grid column ids. */
function tableIds(person: Person): TableIds {
  return tableIdsIn(surface(person).session.part().root);
}

/** Drag the edge between grid columns `k` and `k + 1` (or the right edge when `k` is last). */
export function drag(person: Person, k: number, widths: readonly [number, number]) {
  const { table, grid } = tableIds(person);
  const sourceRevision = surface(person).session.packageRevision();
  if (k === grid.length - 1)
    return tableCommand(person, {
      type: 'commitTableRightEdgeResize',
      target: { sourceRevision, tableId: table, gridColumnId: grid[k], isHeaderRepeat: false },
      columnWidthTwips: widths[0],
      tableWidthTwips: widths[1],
    });
  return tableCommand(person, {
    type: 'commitTableColumnDividerResize',
    target: {
      sourceRevision,
      tableId: table,
      leftGridColumnId: grid[k],
      rightGridColumnId: grid[k + 1],
      isHeaderRepeat: false,
    },
    leftWidthTwips: widths[0],
    rightWidthTwips: widths[1],
  });
}

/** An editor's surface. */
export type Surface = NonNullable<ReturnType<typeof createDocxEditor>['surface']>;

/** A body as the matrix compares it: attributes in order, a revision's `w:id` any id. */
export const comparedBody = (part: OoxmlPart): string =>
  withoutRevisionIds(normalizedBody(part, (id) => id));

/** Put an editor's caret in the paragraph whose text is `text`. */
export function caretInSurface(s: Surface, text: string, offset = 0): void {
  const part = s.session.part();
  const id = s.session
    .paragraphIds()
    .find((candidate) => paragraphTextOf(part, candidate) === text);
  if (!id) throw new Error(`no paragraph ${text}`);
  s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
}

/**
 * What the editor alone (no room) makes of the same steps, as the matrix compares bodies; with
 * `decide`, after it accepts or rejects every revision.
 */
export function alone(
  options: {
    readonly author: string;
    readonly mode: 'suggesting' | 'edit';
    readonly review: EditorModule;
    readonly file?: Uint8Array;
    readonly decide?: 'accept' | 'reject';
  },
  steps: (s: Surface) => void
): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: options.file ?? TABLE_FILE,
    author: options.author,
    mode: options.mode,
    modules: [options.review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    steps(s);
    if (options.decide !== undefined) {
      const command = {
        type: 'resolveAllReviewChanges',
        action: options.decide,
        scope: 'document',
      };
      expect(editor.exec(command as never).ok).toBe(true);
    }
    return comparedBody(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}
