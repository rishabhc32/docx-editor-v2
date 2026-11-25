// A decision after a fold gives the Word file the same decision gives in the room (design D14).
// A pending suggestion splits the run it sits in, and a fold writes each piece as a run of its
// own; a decision that takes the suggestion away draws the room's run whole again. A new list's
// definition or style goes with its suggestion. A table record keeps what it took from another. A
// suggestion keeps its version, so a decision that names what the room showed applies.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { orderedCardSeen } from '../change/apply-decide.ts';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  type OrderedWordFile,
} from '../client/open.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  DIRECT,
  EDIT,
  markOf,
  opens,
  partsOf,
} from './change-apply-fixture.ts';
import { checkFold, checkpointBytes, folded, mediaStore } from './fold-check.ts';
import {
  alphaOmega,
  bodyOf,
  docx,
  MATRIX_EXTRAS,
  MATRIX_ITEMS,
  p,
  t,
  type PackageExtras,
} from './word-fixtures.ts';

const MANIFEST = 'ordered-checkpoint.json';

/** A run as Word writes it, with its revision session and its own properties. */
const run = (text: string, rPr = '') =>
  `<w:r w:rsidR="00A96A35"><w:rPr><w:sz w:val="24"/>${rPr}</w:rPr><w:t>${text}</w:t></w:r>`;

/** The file's own format record on a run, by another author. */
const FILE_FORMAT =
  '<w:rPrChange w:id="7" w:author="Author" w:date="2026-07-07T18:18:00Z">' +
  '<w:rPr><w:color w:val="000000"/></w:rPr></w:rPrChange>';

async function foldOf(
  body: string,
  edit: (doc: OrderedDocument) => void,
  label: string,
  extras?: PackageExtras
) {
  const { document: doc, base } = openOrderedWordFile(docx(body, extras), { file: 1 });
  edit(doc);
  await checkFold(doc, base, mediaStore(base), label);
}

describe('a decision after a fold, on a suggestion inside a run', () => {
  test('a suggested word', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
      },
      'a word inside a run'
    );
  });

  test('a suggested deletion', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 6), to: at(doc, 0, 10) });
      },
      'a deletion inside a run'
    );
  });

  test("a suggested word in a run with the file's format record", async () => {
    await foldOf(
      p(run('Lorem ipsum dolor.', FILE_FORMAT)),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 7), text: 'X ' });
      },
      'a word inside a formatted run'
    );
  });

  // The room's run goes on in the new paragraph: its piece there is drawn with new ids
  test('a suggested paragraph break, beside a deletion in a combo box', async () => {
    const combo = MATRIX_ITEMS.find((item) => item.key === 'combo');
    await foldOf(
      p(run('Cloud and serverless.')) + bodyOf(combo!),
      (doc) => {
        apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 2, 6), to: at(doc, 2, 7) });
        apply(doc, 'bob:1', { kind: 'split', ...EDIT, place: at(doc, 0, 6) }, 'Bob');
      },
      'a paragraph break inside a run'
    );
  });

  // Bob's format holds the run's one change record, and Ana's format on her own new words shows
  // with no record of its own: a reject of Bob's format takes Ana's away too, as Word's does
  for (const [label, from] of [
    ['on her own new words', 6],
    ['from before her new words', 2],
  ] as const) {
    test(`a format ${label}, over another person's format`, async () => {
      await foldOf(
        p(run('Cloud and serverless.')),
        (doc) => {
          apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
          const bold = [{ localName: 'b' }];
          const range = { from: at(doc, 0, 2), to: at(doc, 0, 8) };
          apply(doc, 'bob:1', { kind: 'format', ...EDIT, ...range, properties: bold }, 'Bob');
          const italic = [{ localName: 'i' }];
          const own = { from: at(doc, 0, from), to: at(doc, 0, 8) };
          apply(doc, 'ana:2', { kind: 'format', ...EDIT, ...own, properties: italic });
        },
        `a format ${label}`
      );
    });
  }
});

describe("a checkpoint's room run numbers", () => {
  /** The room run numbers (`oc:run`) the main document of a checkpoint holds, in order. */
  async function numbersOf(edit: (doc: OrderedDocument) => void) {
    const body = p(run('Cloud and serverless.')) + p(run('Untouched.'));
    const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
    edit(doc);
    const bytes = await checkpointBytes(doc, base, mediaStore(base));
    const xml = strFromU8(unzipSync(bytes)['word/document.xml']!);
    return {
      numbers: [...xml.matchAll(/ oc:run="(\d+)"/g)].map((m) => m[1]),
      declared: xml.includes('xmlns:oc='),
    };
  }

  test('only on the pieces of a run a suggestion splits: each run and its text', async () => {
    const { numbers, declared } = await numbersOf((doc) => {
      apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
    });
    expect({ numbers, declared }).toEqual({
      numbers: [numbers[0], numbers[1], numbers[0], numbers[1]],
      declared: true,
    });
    expect(numbers[0]).not.toBe(numbers[1]);
  });

  // The piece after the break is drawn in another paragraph than its own while the break is
  // pending; a reject draws both pieces in one text element again, live and after a fold
  test('a paragraph break inside a run, rejected: one text element', async () => {
    const { document: doc, base } = openOrderedWordFile(docx(p(run('Cloud and serverless.'))), {
      file: 1,
    });
    const blobs = mediaStore(base);
    apply(doc, 'bob:1', { kind: 'split', ...EDIT, place: at(doc, 0, 6) }, 'Bob');
    const once = await folded(doc, base, 2, blobs);
    const twice = await folded(once.document, once.base, 3, blobs);
    const texts = [{ document: doc, base }, once, twice].map((copy) => {
      const suggestions = orderedCardSeen(copy.document, 'S:bob:1.1')!;
      const decide = { kind: 'decide', cards: [{ suggestions }], status: 'rejected' } as const;
      apply(copy.document, 'olive:1', decide, 'Olive');
      const bytes = orderedPackageBytes(copy.document, copy.base, { blobs });
      const xml = strFromU8(unzipSync(bytes)['word/document.xml']!);
      return [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]);
    });
    const one = ['Cloud and serverless.'];
    expect(texts).toEqual([one, one, one]);
  });

  test('none, and no namespace, where no suggestion splits a run', async () => {
    const drawn = await numbersOf((doc) => {
      const text = ' Yes.';
      apply(doc, 'olive:1', { kind: 'insert', ...DIRECT, place: at(doc, 0, 21), text }, 'Olive');
    });
    expect(drawn).toEqual({ numbers: [], declared: false });
  });
});

describe('a decision after a fold, on a new list', () => {
  /** A new list on paragraph `k`: Ana's suggestion, or Olive's direct edit. */
  const listOn = (doc: OrderedDocument, k: number, id = 'ana:1') =>
    apply(doc, id, {
      kind: 'new-list',
      ...(id.startsWith('olive') ? DIRECT : EDIT),
      list: {
        abstract: {
          e: 'w:abstractNum',
          c: [
            {
              e: 'w:lvl',
              a: { 'w:ilvl': '0' },
              c: [
                { e: 'w:numFmt', a: { 'w:val': 'lowerLetter' } },
                { e: 'w:lvlText', a: { 'w:val': '(%1)' } },
              ],
            },
          ],
        },
      },
      paragraphs: [{ mark: markOf(doc, k), level: 0 }],
    });
  const newList = (doc: OrderedDocument) => listOn(doc, 0);
  const ITEMS = p(run('First item.')) + p(run('Second item.'));

  test('in a file with no lists: the numbering part goes with it', async () => {
    await foldOf(ITEMS, newList, 'a new list, no numbering part');
  });

  test("in a file with lists of its own: the file's stay", async () => {
    await foldOf(ITEMS, newList, 'a new list, a numbering part', MATRIX_EXTRAS);
  });

  test('beside a direct new list: the definitions keep their order', async () => {
    await foldOf(
      ITEMS,
      (doc) => {
        listOn(doc, 1, 'olive:1');
        listOn(doc, 0);
      },
      'a new list beside a direct one'
    );
  });
});

describe('a decision after a fold, on a pasted style', () => {
  const STYLE = {
    e: 'w:style',
    a: { 'w:type': 'paragraph', 'w:styleId': 'PastedClause' },
    c: [
      { e: 'w:name', a: { 'w:val': 'Pasted Clause' } },
      { e: 'w:rPr', c: [{ e: 'w:i' }] },
    ],
  } as const;
  const BLOCK = {
    e: 'w:p',
    c: [
      { e: 'w:pPr', c: [{ e: 'w:pStyle', a: { 'w:val': 'PastedClause' } }] },
      { e: 'w:r', c: [{ e: 'w:t', c: ['Pasted.'] }] },
    ],
  } as const;

  test('whole paragraphs with their style: a reject takes the style away', async () => {
    await foldOf(
      p(run('Cloud and serverless.')),
      (doc) => {
        const place = at(doc, 0, 6);
        const paste = { place, blocks: [BLOCK], styles: [STYLE], whole: true };
        apply(doc, 'ana:1', { kind: 'insert-fragment', ...EDIT, ...paste });
      },
      'a pasted style'
    );
  });
});

describe('a decision after a fold, beside a pending control value', () => {
  // A check box's value record names elements of another namespace than the main one: a fold
  // keeps their prefix, so the part a decision edits still reads
  const xmlOf = (key: string) => MATRIX_ITEMS.find((item) => item.key === key)?.xml ?? '';

  test('an accepted deletion empties a drop-down list beside a ticked check box', async () => {
    await foldOf(
      p(alphaOmega(xmlOf('dropdown'))) + p(alphaOmega(xmlOf('checkbox'))),
      (doc) => {
        apply(doc, 'ana:1', {
          kind: 'control-value',
          ...EDIT,
          control: opens(doc, 'sdt')[1]!,
          value: { kind: 'checkbox', checked: true },
        });
        apply(doc, 'bob:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 6), to: at(doc, 0, 9) });
      },
      'a drop-down list beside a check box'
    );
  });
});

describe('a decision after a fold, on a value in a pending control', () => {
  // A direct value in a control a pending insertion put in is part of that insertion (design
  // D9): a reject takes the control with its value, live as after a fold. A party's value in its
  // own pending control is part of that insertion too, and its undo takes back that value alone
  const WRAPS = [
    ['date', { kind: 'date', iso: '2026-09-15' }],
    ['comboBox', { kind: 'text', text: 'Net 30' }],
    ['plainText', { kind: 'text', text: 'Acme' }],
    ['richText', { kind: 'text', text: 'Acme' }],
  ] as const;
  const wrapped = (
    doc: OrderedDocument,
    type: (typeof WRAPS)[number][0],
    who: string,
    mode: typeof EDIT | typeof DIRECT
  ) => {
    const wrap = { from: at(doc, 0, 6), to: at(doc, 0, 12), control: { type, tag: 'c.86' } };
    apply(doc, `${who}:1`, { kind: 'wrap-control', ...EDIT, ...wrap }, who);
    return { ...mode, control: opens(doc, 'sdt')[0]! };
  };

  for (const [type, value] of WRAPS) {
    test(`the organiser's direct value in their own suggested ${type} wrap`, async () => {
      await foldOf(
        p(run('Line: Revenue')),
        (doc) => {
          const control = wrapped(doc, type, 'Olive', DIRECT);
          apply(doc, 'Olive:2', { kind: 'control-value', ...control, value }, 'Olive');
        },
        `a direct value in a suggested ${type} wrap`
      );
    });

    test(`a party's suggested value in its own suggested ${type} wrap`, async () => {
      await foldOf(
        p(run('Line: Revenue')),
        (doc) => {
          apply(doc, 'Ana:2', { kind: 'control-value', ...wrapped(doc, type, 'Ana', EDIT), value });
        },
        `a suggested value in a suggested ${type} wrap`
      );
    });

    test(`a party's undo of each value in its own suggested ${type} wrap`, async () => {
      const again =
        value.kind === 'date'
          ? ({ kind: 'date', iso: '2026-10-01' } as const)
          : ({ kind: 'text', text: 'Beta' } as const);
      await foldOf(
        p(run('Line: Revenue')),
        (doc) => {
          const control = wrapped(doc, type, 'Ana', EDIT);
          const wrap = partsOf(doc);
          apply(doc, 'Ana:2', { kind: 'control-value', ...control, value }, 'Ana');
          const first = partsOf(doc);
          apply(doc, 'Ana:3', { kind: 'control-value', ...control, value: again }, 'Ana');
          apply(doc, 'Ana:4', { kind: 'retract', suggestions: [], undo: ['Ana:3'] }, 'Ana');
          expect(partsOf(doc)).toEqual(first);
          apply(doc, 'Ana:5', { kind: 'retract', suggestions: [], undo: ['Ana:2'] }, 'Ana');
          expect(partsOf(doc)).toEqual(wrap);
        },
        `each suggested value in a suggested ${type} wrap undone`
      );
    });
  }

  const ITEMS =
    '<w:listItem w:displayText="Net 30" w:value="30"/>' +
    '<w:listItem w:displayText="Net 60" w:value="60"/>';
  const CHECK_BOX =
    '<w14:checkbox><w14:checked w14:val="0"/>' +
    '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
    '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>';
  const ANAS = [
    [
      'date',
      '<w:date><w:dateFormat w:val="M/d/yyyy"/></w:date>',
      { kind: 'date', iso: '2026-09-15' },
    ],
    ['drop-down', `<w:dropDownList>${ITEMS}</w:dropDownList>`, { kind: 'listItem', value: '60' }],
    ['combo box', `<w:comboBox>${ITEMS}</w:comboBox>`, { kind: 'text', text: 'Net 90' }],
    ['check box', CHECK_BOX, { kind: 'checkbox', checked: true }],
    ['plain text', '<w:text/>', { kind: 'text', text: 'Beta' }],
    ['rich text', '', { kind: 'text', text: 'Beta' }],
  ] as const;

  for (const [kind, props, value] of ANAS) {
    test(`the organiser's direct value in another person's pending ${kind} control`, async () => {
      await foldOf(
        p(
          run('Term: ') +
            '<w:ins w:id="1" w:author="Ana" w:date="2026-10-28T07:00:00Z"><w:sdt><w:sdtPr>' +
            `<w:id w:val="5"/>${props}</w:sdtPr><w:sdtContent>${run('Net 30')}</w:sdtContent>` +
            '</w:sdt></w:ins>'
        ),
        (doc) => {
          const control = opens(doc, 'sdt')[0]!;
          apply(doc, 'Olive:1', { kind: 'control-value', ...DIRECT, control, value }, 'Olive');
        },
        `a direct value in a pending ${kind} control`
      );
    });
  }
});

describe("a fold of a table's records", () => {
  const ANA_1 = '2026-10-28T09:01:00Z';

  // Two dates: each is decided on its own
  test('two row heights of one person on a row with no properties', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    for (const [id, heightTwips, date] of [
      ['ana:1', 300, ANA_1],
      ['ana:2', 480, '2026-10-28T09:02:00Z'],
    ] as const) {
      const on = [addressOf(opens(doc, 'tr')[0]!)];
      const op = { op: 'setTableRowHeight', heightTwips } as const;
      apply(doc, id, { kind: 'table-property', ...EDIT, on, op }, 'Ana', date);
    }

    await checkFold(doc, base, mediaStore(base), 'two row heights');
  });

  // The organiser's column shows on top of the pending row height's table record, which Word's
  // record holds as the grid before both: the fold keeps the new grid
  test("a row height, then the organiser's direct column", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const on = [addressOf(opens(doc, 'tr')[0]!)];
    const op = { op: 'setTableRowHeight', heightTwips: 480 } as const;
    apply(doc, 'ana:1', { kind: 'table-property', ...EDIT, on, op }, 'Ana', ANA_1);
    const cell = addressOf(opens(doc, 'tc')[0]!);
    apply(doc, 'olive:1', { kind: 'insert-column', ...DIRECT, cell, where: 'left' }, 'Olive');
    await checkFold(doc, base, mediaStore(base), "a row height, then the organiser's column");
  });

  // A direct column beside the organiser's own new cell is made over the row height's grid
  // record, so a reject of the height takes the direct grid away too, after a fold as live; and
  // the row the reject leaves with no properties writes none
  test("a row height, a column, then the organiser's direct column beside its new cell", async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const height = (row: number) =>
      ({
        kind: 'table-property',
        on: [addressOf(opens(doc, 'tr')[row]!)],
        op: { op: 'setTableRowHeight', heightTwips: 500 },
      }) as const;
    const right = (k: number) =>
      ({ cell: addressOf(opens(doc, 'tc')[k]!), where: 'right' }) as const;
    apply(doc, 'ana:1', { ...height(1), ...EDIT }, 'Ana', ANA_1);
    apply(doc, 'olive:1', { kind: 'insert-column', ...EDIT, ...right(3) }, 'Olive');
    apply(doc, 'olive:2', { ...height(0), ...DIRECT }, 'Olive');
    apply(doc, 'olive:3', { kind: 'insert-column', ...DIRECT, ...right(4) }, 'Olive');
    await checkFold(doc, base, mediaStore(base), 'a direct column over a row height');
  });

  // The organiser's direct column deletion after a fold, over a row height made before it: a
  // decision on the height fits the grid to the cells, as the engine's does, after another fold
  // as live (the fold holds the deleted cells no more)
  test("a row height, a fold, then the organiser's direct column deletion", async () => {
    const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const blobs = mediaStore(first.base);
    const height = (doc: OrderedDocument, row: number) =>
      ({
        kind: 'table-property',
        ...EDIT,
        on: [addressOf(opens(doc, 'tr')[row]!)],
        op: { op: 'setTableRowHeight', heightTwips: 400 },
      }) as const;
    apply(first.document, 'ana:1', height(first.document, 1), 'Ana', ANA_1);
    const row = addressOf(opens(first.document, 'tr')[0]!);
    apply(first.document, 'olive:1', { kind: 'insert-row', ...EDIT, row, where: 'below' }, 'Olive');
    const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
    apply(doc, 'ana:2', height(doc, 2), 'Ana', ANA_1);
    const cell = addressOf(opens(doc, 'tc')[5]!);
    apply(doc, 'olive:2', { kind: 'delete-column', ...DIRECT, cell }, 'Olive');
    await checkFold(doc, base, blobs, 'a direct column deletion over a row height');
  });

  // The same for the organiser's table alignment: it goes in the schema's place, before borders
  test("a row height, then the organiser's table alignment", async () => {
    const cell = (text: string) =>
      `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(run(text))}</w:tc>`;
    const table =
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' +
      '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('Fee')}${cell('Due')}</w:tr></w:tbl>`;
    const { document: doc, base } = openOrderedWordFile(docx(table + p(run('After.'))), {
      file: 1,
    });
    const on = [addressOf(opens(doc, 'tr')[0]!)];
    const op = { op: 'setTableRowHeight', heightTwips: 480 } as const;
    apply(doc, 'ana:1', { kind: 'table-property', ...EDIT, on, op }, 'Ana', ANA_1);
    const align = { op: 'setTableProperties', justification: 'center' } as const;
    const tbl = [addressOf(opens(doc, 'tbl')[0]!)];
    apply(doc, 'olive:1', { kind: 'table-property', ...DIRECT, on: tbl, op: align }, 'Olive');
    await checkFold(doc, base, mediaStore(base), "a row height, then the organiser's alignment");
  });

  // A row height records each cell and the grid, and sets nothing there: a decision on a column
  // gives back or gives on the widths the column holds, after a fold too
  const rowHeight = (doc: OrderedDocument, id: string, row: number, date: string) => {
    const on = [addressOf(opens(doc, 'tr')[row]!)];
    const op = { op: 'setTableRowHeight', heightTwips: 480 } as const;
    apply(doc, id, { kind: 'table-property', ...EDIT, on, op }, 'Bob', date);
  };
  const column = (doc: OrderedDocument, kind: 'insert-column' | 'delete-column') => {
    const cell = addressOf(opens(doc, 'tc')[1]!);
    const body =
      kind === 'insert-column'
        ? ({ kind, ...EDIT, cell, where: 'left' } as const)
        : ({ kind, ...EDIT, cell } as const);
    apply(doc, 'ana:1', body, 'Ana', ANA_1);
  };

  for (const kind of ['insert-column', 'delete-column'] as const) {
    test(`${kind}, then a row height on each row`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      column(doc, kind);
      rowHeight(doc, 'bob:1', 0, '2026-10-28T09:06:00Z');
      rowHeight(doc, 'bob:2', 1, '2026-10-28T09:07:00Z');
      await checkFold(doc, base, mediaStore(base), `${kind}, then row heights`);
    });

    // The room's shape: the later row height is made on a folded copy
    test(`${kind} and a row height, folded, then a row height on the other row`, async () => {
      const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const blobs = mediaStore(first.base);
      column(first.document, kind);
      rowHeight(first.document, 'bob:1', 0, '2026-10-28T09:06:00Z');
      const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
      rowHeight(doc, 'bob:2', 1, '2026-10-28T09:07:00Z');
      await checkFold(doc, base, blobs, `${kind}, folded, then a row height`);
    });

    test(`${kind}, folded, then a row height`, async () => {
      const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const blobs = mediaStore(first.base);
      column(first.document, kind);
      const { document: doc, base } = await folded(first.document, first.base, 2, blobs);
      rowHeight(doc, 'bob:1', 1, '2026-10-28T09:06:00Z');
      await checkFold(doc, base, blobs, `${kind}, folded, then a row height`);
    });

    // A rejected row record leaves the row's cells that hold a record as they are: the fold keeps
    // the record on each cell, not only on the cells the column's pending widths are kept on
    test(`${kind}, then a row height rejected`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      column(doc, kind);
      rowHeight(doc, 'bob:1', 0, '2026-10-28T09:06:00Z');
      const suggestions = orderedCardSeen(doc, 'S:bob:1.1')!;
      const decide = { kind: 'decide', cards: [{ suggestions }], status: 'rejected' } as const;
      apply(doc, 'olive:1', decide, 'Olive');
      await checkFold(doc, base, mediaStore(base), `${kind}, then a row height rejected`);
    });
  }

  // A table width on the first row's second cell: it takes Bob's table records
  const tableWidth = (doc: OrderedDocument, id: string, columnWidthTwips: number) => {
    const on = [addressOf(opens(doc, 'tc')[1]!)];
    const tableWidthTwips = columnWidthTwips + 3000;
    const op = { op: 'setTableRightEdgeWidth', columnWidthTwips, tableWidthTwips } as const;
    apply(doc, id, { kind: 'table-property', ...EDIT, on, op }, 'Bob');
  };

  // A table width takes the row height's table record. Once the row height is decided, that
  // record takes over nothing, but it stays apart: the file's one change record holds it, and a
  // fold gives it back
  for (const status of ['accepted', 'rejected'] as const) {
    test(`a row height, a table width on its row, then the row height ${status}`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      rowHeight(doc, 'bob:1', 0, '2026-10-28T09:06:00Z');
      tableWidth(doc, 'bob:2', 2600);
      const suggestions = orderedCardSeen(doc, 'S:bob:1.1')!;
      apply(doc, 'olive:1', { kind: 'decide', cards: [{ suggestions }], status }, 'Olive');
      await checkFold(doc, base, mediaStore(base), `a table width over a ${status} row height`);

      // The reader stays strict: without the takeover, the records do not fit the file
      const entries = unzipSync(await checkpointBytes(doc, base, mediaStore(base)));
      const manifest = JSON.parse(strFromU8(entries[MANIFEST]!)) as { takeovers: unknown[] };
      expect(manifest.takeovers).not.toEqual([]);
      entries[MANIFEST] = strToU8(JSON.stringify({ ...manifest, takeovers: [] }));
      expect(() => openOrderedCheckpoint(zipSync(entries), { file: 2 })).toThrow(
        'checkpoint-manifest'
      );
    });
  }

  // A take-back gives a suggestion back the records another took from it, two of one target:
  // Word's one change record holds them, and a fold reads one record that sets what both set
  test('two table widths over a row height, then the second taken back', async () => {
    const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
    rowHeight(doc, 'bob:1', 0, '2026-10-28T09:06:00Z');
    tableWidth(doc, 'bob:2', 2600);
    tableWidth(doc, 'bob:3', 2400);
    apply(doc, 'bob:4', { kind: 'retract', suggestions: ['S:bob:3.1'] }, 'Bob');
    await checkFold(doc, base, mediaStore(base), 'a table width taken back over another');
  });

  // A decision on a column gives its width back as the engine's reject does, from the table
  // before the decision: the cell the column took width from at the width it lent, the others as
  // they show, live as after a fold
  const right = (doc: OrderedDocument, k: number) =>
    ({ cell: addressOf(opens(doc, 'tc')[k]!), where: 'right' }) as const;
  for (const mode of [EDIT, DIRECT]) {
    test(`a column, then the organiser's ${mode.mode} column beside its new cell`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      apply(doc, 'ana:1', { kind: 'insert-column', ...EDIT, ...right(doc, 0) }, 'Ana', ANA_1);
      apply(doc, 'olive:1', { kind: 'insert-column', ...mode, ...right(doc, 1) }, 'Olive');
      await checkFold(doc, base, mediaStore(base), `a ${mode.mode} column beside a new cell`);
    });
  }

  // The organiser deletes the column before the party's column, or the one it took width from
  for (const k of [0, 1]) {
    test(`a column, then the organiser's deletion of column ${k + 1}`, async () => {
      const { document: doc, base } = openOrderedWordFile(APPLY_FILE, { file: 1 });
      apply(doc, 'ben:1', { kind: 'insert-column', ...EDIT, ...right(doc, 1) }, 'Ben', ANA_1);
      const cell = addressOf(opens(doc, 'tc')[k]!);
      apply(doc, 'olive:1', { kind: 'delete-column', ...DIRECT, cell }, 'Olive');
      await checkFold(doc, base, mediaStore(base), `a column, then column ${k + 1} deleted`);
    });
  }
});

describe("a suggestion's version after a fold", () => {
  // Each pending suggestion's card as the organiser sends it: its suggestions and their versions
  const seenOf = (doc: OrderedDocument) =>
    [...doc.suggestions.values()]
      .filter((s) => s.status === 'pending')
      .map((s) => orderedCardSeen(doc, s.id));

  // Words a person adds to their own insertion count on it; a table op's records, which an open
  // gives back their room facts, count on nothing
  test('typing that joins an insertion, beside a column and a row height', async () => {
    const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
    const { document: doc } = first;
    const blobs = mediaStore(first.base);
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'net ' });
    apply(doc, 'ana:2', { kind: 'insert', ...EDIT, place: at(doc, 0, 8), text: 'x' });
    const cell = addressOf(opens(doc, 'tc')[1]!);
    apply(doc, 'bob:1', { kind: 'insert-column', ...EDIT, cell, where: 'left' }, 'Bob');
    const on = [addressOf(opens(doc, 'tr')[0]!)];
    const op = { op: 'setTableRowHeight', heightTwips: 480 } as const;
    apply(doc, 'bob:2', { kind: 'table-property', ...EDIT, on, op }, 'Bob');
    const seen = seenOf(doc);
    expect(seen.some((card) => card?.some(({ version }) => version > 0))).toBe(true);

    const once = await folded(doc, first.base, 2, blobs);
    const twice = await folded(once.document, once.base, 3, blobs);
    expect({ once: seenOf(once.document), twice: seenOf(twice.document) }).toEqual({
      once: seen,
      twice: seen,
    });

    // The organiser decides each card as the room showed it
    for (const suggestions of seen) {
      apply(
        twice.document,
        `olive:${seen.indexOf(suggestions) + 1}`,
        { kind: 'decide', cards: [{ suggestions: suggestions! }], status: 'accepted' },
        'Olive'
      );
    }

    expect(seenOf(twice.document)).toEqual([]);
  });
});

/** The body as written once the organiser decided every pending card. */
function decidedBody(
  { document: doc, base }: OrderedWordFile,
  blobs: ReturnType<typeof mediaStore>,
  status: 'accepted' | 'rejected'
): string {
  let n = 0;
  for (const { id, status: now } of [...doc.suggestions.values()]) {
    const suggestions = now === 'pending' ? orderedCardSeen(doc, id) : null;
    if (suggestions) {
      n += 1;
      apply(doc, `olive:${100 + n}`, { kind: 'decide', cards: [{ suggestions }], status }, 'Olive');
    }
  }

  const xml = strFromU8(unzipSync(orderedPackageBytes(doc, base, { blobs }))['word/document.xml']!);
  return xml.replace(/ (?:w14:paraId|w14:textId|w:id|w:date)="[^"]*"/g, '');
}

describe('a decision after a fold, on properties drawn for a change record', () => {
  // Compose draws a `w:pPr` or a `w:rPr` where the paragraph or the run has none, only to hold a
  // change record: a decision that takes the record away draws none again, after a fold too.
  // The fold check reads an empty one as none, so this compares the decided body as written
  const FILE = docx(p(t('First para.')) + p(t('Second para.')) + p(t('Third para.')));
  const right = [{ localName: 'jc', attributes: { val: 'right' } }];
  const SHAPES: [string, (doc: OrderedDocument) => void][] = [
    ['a join', (doc) => apply(doc, 'ana:1', { kind: 'join', ...EDIT, mark: markOf(doc, 0) })],
    [
      'a paragraph format',
      (doc) =>
        apply(doc, 'ana:1', {
          kind: 'paragraph-property',
          ...EDIT,
          marks: [markOf(doc, 2)],
          properties: right,
        }),
    ],
    [
      'a paragraph mark format',
      (doc) =>
        apply(doc, 'ana:1', {
          kind: 'mark-format',
          ...EDIT,
          marks: [markOf(doc, 2)],
          properties: [{ localName: 'b' }],
        }),
    ],
    [
      'a run format',
      (doc) =>
        apply(doc, 'ana:1', {
          kind: 'format',
          ...EDIT,
          from: at(doc, 0, 0),
          to: at(doc, 0, 5),
          properties: [{ localName: 'b' }],
        }),
    ],
    [
      'a paragraph break, and a format on the paragraph after it',
      (doc) => {
        apply(doc, 'ana:1', { kind: 'split', ...EDIT, place: at(doc, 1, 6) });
        apply(
          doc,
          'olive:1',
          { kind: 'paragraph-property', ...EDIT, marks: [markOf(doc, 2)], properties: right },
          'Olive'
        );
      },
    ],
  ];

  for (const [name, edit] of SHAPES) {
    for (const status of ['accepted', 'rejected'] as const) {
      test(`${name}, ${status}`, async () => {
        const first = openOrderedWordFile(FILE, { file: 1 });
        const blobs = mediaStore(first.base);
        edit(first.document);
        const once = await folded(first.document, first.base, 2, blobs);
        const twice = await folded(once.document, once.base, 3, blobs);
        const [live, ...copies] = [first, once, twice].map((copy) =>
          decidedBody(copy, blobs, status)
        );
        expect(copies).toEqual([live, live]);
      });
    }
  }
});

describe('a decision after a fold, on a section', () => {
  // The organiser's direct break gives the body's section a type over the organiser's own pending
  // page setup: Word keeps one record for both, so a reject of the page setup takes the type too
  // and lays back what the record states, as the engine's reject of the saved file does
  for (const status of ['accepted', 'rejected'] as const) {
    test(`a page setup, then the organiser's direct break, ${status}`, async () => {
      const first = openOrderedWordFile(APPLY_FILE, { file: 1 });
      const { document: doc } = first;
      const blobs = mediaStore(first.base);
      const page = { marginLeftTwips: 1800 };
      const pageAt = markOf(doc, 9);
      const setup = { kind: 'section-property', ...EDIT, sections: [], page, pageAt } as const;
      apply(doc, 'olive:1', setup, 'Olive');
      const mark = markOf(doc, 0);
      const brk = { kind: 'section-break', ...DIRECT, mark, type: 'continuous' } as const;
      apply(doc, 'olive:2', brk, 'Olive');
      const once = await folded(doc, first.base, 2, blobs);
      expect(decidedBody(once, blobs, status)).toBe(decidedBody(first, blobs, status));
    });
  }
});
