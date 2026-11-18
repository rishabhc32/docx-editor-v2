// The matrix sets past spike 5's four: paragraph presses, structures (rows, columns, tables,
// links, control values, notes), and mixed content that three people edit.

import type { OoxmlProperty } from '../../../../store/store/tree-op-types.ts';
import type { OrderedParagraphOp, OrderedTableOp } from '../../index.ts';
import { alphaOmega, p, pictureRun, t } from '../word-fixtures.ts';
import {
  AFTER,
  ANA,
  BEFORE,
  BEN,
  CY,
  edits,
  kindNamed,
  OLIVE,
  present,
  scenario,
  WIDE_FROM,
  wideTo,
  type Action,
  type Scenario,
  type Who,
} from './actions.ts';
import { paragraphOf, type MatrixKind, type Selector } from './catalogue.ts';
import { cell, furniture, list, note, sdt, table, textBoxText } from './kinds.ts';

const props = (...list: OoxmlProperty[]): OrderedParagraphOp => ({
  op: 'setParagraphProperties',
  properties: list,
});
const val = (localName: string, value: string): OoxmlProperty => ({
  localName,
  attributes: { val: value },
});

/** Paragraph presses, by name. */
export const PRESS_OPS: Readonly<Record<string, OrderedParagraphOp>> = {
  'align right': props(val('jc', 'right')),
  justify: props(val('jc', 'both')),
  'align left': props(val('jc', 'left')),
  'space before and after': props({
    localName: 'spacing',
    attributes: { before: '240', after: '120' },
  }),
  'exact line spacing': props({
    localName: 'spacing',
    attributes: { line: '300', lineRule: 'exact' },
  }),
  'left and right indents': props({ localName: 'ind', attributes: { left: '720', right: '360' } }),
  'hanging indent': props({ localName: 'ind', attributes: { left: '720', hanging: '360' } }),
  'first-line indent': props({ localName: 'ind', attributes: { firstLine: '480' } }),
  'tab stops': {
    op: 'setParagraphTabStops',
    stops: [
      { positionTwips: 4680, alignment: 'center', leader: 'dot' },
      { positionTwips: 9000, alignment: 'right', leader: 'underscore' },
    ],
  },
  'keep with next': props({ localName: 'keepNext' }),
  'widow control off': props(val('widowControl', '0')),
  'contextual spacing': props({ localName: 'contextualSpacing' }),
  'outline level': props(val('outlineLvl', '2')),
  'shading fill': props({
    localName: 'shd',
    attributes: { val: 'clear', color: 'auto', fill: 'EEEEEE' },
  }),
  'Heading 1': props(val('pStyle', 'Heading1')),
  'Quote style': props(val('pStyle', 'Quote')),
  'bullets on': { op: 'setListNumbering', numId: '2' },
  'restart at 5': { op: 'setListNumbering', numId: '5' },
};

/** Presses that need a list paragraph. */
const LIST_OPS: Readonly<Record<string, OrderedParagraphOp>> = {
  'numbering off': { op: 'setListNumbering', numId: null },
  'list level 2': { op: 'setListLevel', level: 1 },
};

const PRESS_KINDS = [
  'bold',
  'list',
  'heading',
  'formatted',
  'markfmt',
  'section',
  'sdtblock',
  'table',
  'bullet',
  'multilevel',
  'listrun',
  'listcell',
  'right',
  'tabstops',
  'keeps',
  'borders',
  'hanging',
  'bidi',
  'quote',
  'dropcap',
  'pprchange',
  'header',
  'footnote',
  'boxtext',
];
const LIST_KINDS = new Set([
  'list',
  'bullet',
  'multilevel',
  'roman',
  'restart',
  'listrun',
  'listcell',
]);

const press = (who: Who, kind: MatrixKind, op: OrderedParagraphOp): Action => ({
  ...who,
  p: paragraphOf(kind),
  kind: 'paragraph',
  ops: [op],
});

const PRESS_PAIRS: readonly (readonly [string, string, Who])[] = [
  ['align right', 'Heading 1', BEN],
  ['Heading 1', 'left and right indents', BEN],
  ['bullets on', 'justify', OLIVE],
  ['tab stops', 'keep with next', BEN],
];

export const PRESSES: readonly Scenario[] = PRESS_KINDS.map(kindNamed).flatMap((kind) => {
  const ops = { ...PRESS_OPS, ...(LIST_KINDS.has(kind.key) ? LIST_OPS : {}) };
  const singles = Object.entries(ops).map(([name, op]) =>
    scenario('presses', kind, name, [press(ANA, kind, op)])
  );
  const pairs = PRESS_PAIRS.map(([first, second, who]) => {
    const one = PRESS_OPS[first];
    const two = PRESS_OPS[second];
    if (!one || !two) throw new Error(`no press ${first} or ${second}`);
    return scenario('presses', kind, `Ana: ${first}, ${who.author}: ${second}`, [
      press(ANA, kind, one),
      press(who, kind, two),
    ]);
  });
  // A press, then text typed and struck in the paragraph by another person
  const edited = scenario('presses', kind, 'Ana: align right, Ben deletes across', [
    press(ANA, kind, PRESS_OPS['align right'] ?? props()),
    edits(BEN)['delete across']?.(paragraphOf(kind), kind) ?? null,
  ]);
  return [...singles, ...pairs, edited].filter(present);
});

// ---- structures ----

const TABLE_KINDS = [
  'table',
  'span',
  'vmerge',
  'nested',
  'headrow',
  'listcell',
  'tblborders',
  'widths',
  'tblstyle',
  'rtltable',
  'floattable',
  'textdir',
  'cantsplit',
  'mergeboth',
  'cellrev',
  'pictable',
];

const TABLE_PRESSES: Readonly<Record<string, OrderedTableOp>> = {
  'cell fill': { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } },
  'row height': { op: 'setTableRowHeight', heightTwips: 600 },
  'table centered': { op: 'setTableProperties', justification: 'center' },
  'cell vertical alignment': { op: 'setTableCellVerticalAlignment', alignment: 'center' },
  'cell borders': {
    op: 'setTableCellBorders',
    scope: 'all',
    spec: { style: 'single', size: 8, color: { kind: 'hex', value: 'FF0000' } },
  },
  'table style': { op: 'setTableProperties', styleId: 'TableGrid', headerRowCount: 1 },
};

/** Table kinds where the caret's cell has a column to its right: an edge to drag. */
const EDGE_KINDS = new Set(['table', 'tblborders', 'widths', 'vmerge', 'mergeboth']);

/**
 * The drag the editor offers at the caret's cell: the edge on its right, or, where the cell
 * reaches the table's right edge, that edge.
 */
const dragOf = (kind: MatrixKind): readonly [string, OrderedTableOp] => {
  if (EDGE_KINDS.has(kind.key))
    return [
      'column widths',
      { op: 'setTableColumnWidths', leftWidthTwips: 4000, rightWidthTwips: 2000 },
    ];
  // The table's width is its columns' sum: the kinds' columns are 3000 wide, the last one 4500
  const grid = /<w:tblGrid>(.*?)<\/w:tblGrid>/.exec(kind.blocks ?? '')?.[1] ?? '';
  const columns = Math.max(1, grid.split('<w:gridCol ').length - 1);
  const tableWidthTwips = 4500 + 3000 * (columns - 1);
  return ['right edge', { op: 'setTableRightEdgeWidth', columnWidthTwips: 4500, tableWidthTwips }];
};

function tableScenarios(kind: MatrixKind): (Scenario | null)[] {
  const at = paragraphOf(kind);
  const row = (who: Who, action: Partial<Action>): Action =>
    ({ ...who, p: at, ...action }) as Action;
  const tablePress = (who: Who, name: string): Action => {
    const op = TABLE_PRESSES[name];
    if (!op) throw new Error(`no table press ${name}`);
    return row(who, { kind: 'tableProperty', op });
  };
  const [dragName, drag] = dragOf(kind);
  const list: (Scenario | null)[] = [
    scenario('structures', kind, 'insert row above', [
      row(ANA, { kind: 'insertRow', where: 'above' }),
    ]),
    scenario('structures', kind, 'insert row below', [
      row(ANA, { kind: 'insertRow', where: 'below' }),
    ]),
    scenario('structures', kind, 'delete row', [row(ANA, { kind: 'deleteRow' })]),
    scenario('structures', kind, 'insert column left', [
      row(ANA, { kind: 'insertColumn', where: 'left' }),
    ]),
    scenario('structures', kind, 'insert column right', [
      row(ANA, { kind: 'insertColumn', where: 'right' }),
    ]),
    scenario('structures', kind, 'delete column', [row(ANA, { kind: 'deleteColumn' })]),
    ...Object.entries(TABLE_PRESSES).map(([name, op]) =>
      scenario('structures', kind, name, [row(ANA, { kind: 'tableProperty', op })])
    ),
    scenario('structures', kind, dragName, [row(ANA, { kind: 'tableProperty', op: drag })]),
    scenario('structures', kind, `Olive: ${dragName}`, [
      { ...OLIVE, p: at, kind: 'tableProperty', op: drag },
    ]),
    scenario('structures', kind, 'Ana deletes the row, Ben deletes it too', [
      row(ANA, { kind: 'deleteRow' }),
      row(BEN, { kind: 'deleteRow' }),
    ]),
    scenario('structures', kind, 'Ana types in the row, Ben deletes the row', [
      edits(ANA)['type before']?.(at, kind) ?? null,
      row(BEN, { kind: 'deleteRow' }),
    ]),
    scenario('structures', kind, 'Ana inserts a row, Olive deletes the row below it', [
      row(ANA, { kind: 'insertRow', where: 'above' }),
      { ...OLIVE, p: at + 1, kind: 'deleteRow' },
    ]),
    scenario('structures', kind, 'Olive inserts a column', [
      { ...OLIVE, p: at, kind: 'insertColumn', where: 'right' },
    ]),
    scenario('structures', kind, 'Olive deletes a column', [
      { ...OLIVE, p: at, kind: 'deleteColumn' },
    ]),
    scenario('structures', kind, 'Ana inserts a column, Ben inserts one too', [
      row(ANA, { kind: 'insertColumn', where: 'right' }),
      row(BEN, { kind: 'insertColumn', where: 'left' }),
    ]),
    scenario('structures', kind, 'Ana deletes the column, Ben inserts one beside it', [
      row(ANA, { kind: 'deleteColumn' }),
      row(BEN, { kind: 'insertColumn', where: 'right' }),
    ]),
    scenario('structures', kind, 'Ana fills the cell, Ben types inside', [
      row(ANA, {
        kind: 'tableProperty',
        op: TABLE_PRESSES['cell fill'] ?? { op: 'setTableRowHeight', heightTwips: 1 },
      }),
      edits(BEN)['type after']?.(at, kind) ?? null,
    ]),
    // A second change of the same cell's or table's properties: Word keeps one record (task 4.4)
    scenario('structures', kind, 'Ana fills the cell, Ben: cell borders', [
      tablePress(ANA, 'cell fill'),
      tablePress(BEN, 'cell borders'),
    ]),
    scenario('structures', kind, 'Ana: table centered, Ben: table style', [
      tablePress(ANA, 'table centered'),
      tablePress(BEN, 'table style'),
    ]),
    scenario('structures', kind, 'Ana fills the cell, Ben gives it the same fill', [
      tablePress(ANA, 'cell fill'),
      tablePress(BEN, 'cell fill'),
    ]),
    scenario('structures', kind, 'Ana sets the row height, then fills the cell', [
      tablePress(ANA, 'row height'),
      tablePress(ANA, 'cell fill'),
    ]),
  ];
  return list;
}

const insertTable = (kind: MatrixKind) =>
  scenario('structures', kind, 'new table before it', [
    { ...ANA, p: paragraphOf(kind), kind: 'insertTable', rows: 2, cols: 2 },
  ]);

const link = (kind: MatrixKind, inside: Selector) => [
  scenario('structures', kind, 'link target', [
    { ...ANA, p: paragraphOf(kind), kind: 'link', at: inside, url: 'https://example.com/new' },
  ]),
  scenario('structures', kind, 'Ben types in it, Ana changes the target', [
    { ...BEN, p: paragraphOf(kind), kind: 'insert', at: inside, text: 'Z' },
    {
      ...ANA,
      p: paragraphOf(kind),
      kind: 'link',
      at: { find: inside.find.slice(0, 2), edge: 'start', delta: 1 },
      url: 'https://example.com/new',
    },
  ]),
];

const control = (key: string, value: string) => {
  const kind = kindNamed(key);
  return [
    scenario('structures', kind, `value ${value}`, [
      { ...ANA, p: paragraphOf(kind), kind: 'control', value },
    ]),
    scenario('structures', kind, `value ${value}, Ben deletes across`, [
      { ...ANA, p: paragraphOf(kind), kind: 'control', value },
      {
        ...BEN,
        p: paragraphOf(kind),
        kind: 'deleteRange',
        from: WIDE_FROM,
        to: wideTo(paragraphOf(kind)),
      },
    ]),
  ];
};

const notes = (key: string) => {
  const kind = kindNamed(key);
  return (['footnote', 'endnote'] as const).map((note) =>
    scenario('structures', kind, `new ${note}`, [
      { ...ANA, p: paragraphOf(kind), kind: 'insertNote', at: AFTER, note },
    ])
  );
};

/** Group 9.5: a page field a party proposes, and the organiser's "Page X of Y" and bookmark. */
const fields = (key: string) => {
  const kind = kindNamed(key);
  const p = paragraphOf(kind);
  return [
    scenario('structures', kind, 'page field after it', [
      { ...ANA, p, kind: 'pageField', at: AFTER, field: 'PAGE' },
    ]),
    scenario('structures', kind, '"Page X of Y" before it, direct', [
      { ...OLIVE, p, kind: 'pageField', at: BEFORE, field: 'PAGE_X_OF_Y' },
    ]),
    scenario('structures', kind, 'bookmark around it, direct, Ben types inside', [
      { ...OLIVE, p, kind: 'bookmark', at: BEFORE, to: AFTER, name: 'Mark' },
      edits(BEN)['type before']?.(p, kind) ?? null,
    ]),
  ];
};

export const STRUCTURES: readonly Scenario[] = [
  ...TABLE_KINDS.map(kindNamed).flatMap(tableScenarios),
  ...['bold', 'list', 'heading', 'bullet', 'picture'].map(kindNamed).map(insertTable),
  ...link(kindNamed('link'), { find: 'Docs', edge: 'start', delta: 2 }),
  ...link(kindNamed('anchorlink'), { find: 'Jump', edge: 'start', delta: 2 }),
  ...control('dropdown', '2'),
  ...control('combo', 'r'),
  ...control('date', '2026-03-04'),
  ...control('checkbox', 'true'),
  ...control('dropprompt', '2'),
  ...control('dateprompt', '2026-03-04'),
  ...['bold', 'list', 'table', 'heading', 'picture'].flatMap(notes),
  ...['bold', 'link', 'pagefield', 'heading', 'picture', 'ins'].flatMap(fields),
].filter(present);

// ---- three people on mixed content ----

const field = (instr: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';
const inside = (find: string, delta: number): Selector => ({ find, edge: 'start', delta });

/** Mixed content: what a contract holds side by side. */
const MIXED: readonly (MatrixKind & { readonly parts: readonly string[] })[] = [
  {
    key: 'listcell+picture+footnote',
    parts: ['listcell', 'bullet', 'pictable', 'fnref'],
    name: 'bulleted item in a table cell with a picture and a footnote',
    group: 'table',
    blocks: table(
      1,
      `<w:tr>${cell(p(alphaOmega(t('Fee ') + pictureRun(31) + '<w:r><w:footnoteReference w:id="1"/></w:r>' + t(' due')), list(2)))}</w:tr>`
    ),
    text: 'Fee ￼￼ due',
    inside: inside('Fee', 2),
  },
  {
    key: 'heading+field+bookmark',
    parts: ['heading', 'pagefield', 'bookmark'],
    name: 'heading with a page field inside a bookmark',
    group: 'block',
    blocks: p(
      alphaOmega(
        `<w:bookmarkStart w:id="4" w:name="h1"/>${t('Page ')}${field('PAGE', t('1'))}<w:bookmarkEnd w:id="4"/>`
      ),
      '<w:pStyle w:val="Heading1"/>'
    ),
    text: 'Page ￼',
    inside: inside('Page', 2),
  },
  {
    key: 'justified+link+control+math',
    parts: ['justify', 'tabstops', 'link', 'plain', 'math'],
    name: 'justified paragraph with tab stops, a link, a control and an equation',
    group: 'block',
    blocks: p(
      alphaOmega(
        `<w:hyperlink r:id="rIdLink">${t('Docs', '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>${t(' and ')}` +
          sdt('<w:alias w:val="Name"/><w:text/>', t('Name')) +
          `<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x=1</m:t></m:r></m:oMath>`
      ),
      '<w:jc w:val="both"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8000"/></w:tabs>'
    ),
    text: 'Docs and Name￼',
    inside: inside(' and ', 2),
  },
  {
    key: 'list+revision+comment',
    parts: ['list', 'ins', 'comment'],
    name: 'numbered item with an insertion from Word and a comment',
    group: 'revision',
    blocks: p(
      alphaOmega(
        `<w:ins w:id="941" ${REV}>${t('earlier ')}</w:ins><w:commentRangeStart w:id="0"/>${t('noted')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>`
      ),
      list(1)
    ),
    text: 'earlier noted',
    inside: inside('noted', 2),
  },
  {
    parts: ['header', 'picheader', 'pagefield'],
    ...furniture(
      'header+picture+field',
      'header with a picture and a page field',
      'header',
      pictureRun(32) + t(' page ') + field('PAGE', t('1')),
      '￼ page ￼'
    ),
    inside: inside(' page', 2),
  },
  {
    parts: ['footnote', 'link', 'bold'],
    ...note(
      'footnote+link+bold',
      'footnote with a link and bold text',
      'footnote',
      `<w:hyperlink r:id="rIdLink">${t('Docs')}</w:hyperlink>${t(' bold', '<w:b/>')}`
    ),
    text: 'Docs bold',
    inside: inside(' bold', 2),
  },
  {
    parts: ['boxtext', 'link', 'bold'],
    ...textBoxText(
      'boxtext+link+bold',
      'text box with a link and bold text',
      `<w:hyperlink r:id="rIdLink">${t('Docs')}</w:hyperlink>${t(' bold', '<w:b/>')}`,
      'Docs bold'
    ),
    inside: inside(' bold', 2),
  },
];

const THREE: readonly (readonly [string, readonly (readonly [Who, string])[]])[] = [
  [
    'Ana types inside, Ben deletes across, Cy bolds across',
    [
      [ANA, 'type inside'],
      [BEN, 'delete across'],
      [CY, 'bold across'],
    ],
  ],
  [
    'Ana deletes it, Ben types before, Cy replaces across',
    [
      [ANA, 'delete it'],
      [BEN, 'type before'],
      [CY, 'replace across'],
    ],
  ],
  [
    'Ana bolds across, Ben italics it, Cy deletes it',
    [
      [ANA, 'bold across'],
      [BEN, 'italic it'],
      [CY, 'delete it'],
    ],
  ],
  [
    'Ana types after, Cy deletes across, Olive types before',
    [
      [ANA, 'type after'],
      [CY, 'delete across'],
      [OLIVE, 'type before'],
    ],
  ],
  [
    'Ana Enter after, Ben centers, Cy types before',
    [
      [ANA, 'Enter after'],
      [BEN, 'center'],
      [CY, 'type before'],
    ],
  ],
];

export const TRIO: readonly Scenario[] = MIXED.flatMap((kind) =>
  THREE.map(([name, steps]) =>
    scenario(
      'trio',
      kind,
      name,
      steps.map(([who, edit]) => edits(who)[edit]?.(paragraphOf(kind), kind) ?? null),
      kind.parts
    )
  ).filter(present)
);
