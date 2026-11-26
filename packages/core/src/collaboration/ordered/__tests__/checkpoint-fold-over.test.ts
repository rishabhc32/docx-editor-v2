// A write made over a pending record with no change record of its own (a direct one, or one on
// its author's own new words) stays when the record's author takes the record back, after a fold
// too: Word keeps one change record, so the checkpoint keeps the room's records (design D14). The
// manifest is untrusted: an entry that does not fit the file refuses the whole checkpoint.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  openOrderedCheckpoint,
  openOrderedWordFile,
  OrderedWordFileError,
} from '../client/open.ts';
import { orderedCardSeen } from '../change/apply-decide.ts';
import { writeOrderedCheckpointVersion5 } from '../compose/checkpoint.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  at,
  DIRECT,
  EDIT,
  markOf,
  opens,
  partsOf,
} from './change-apply-fixture.ts';
import { checkFold, checkpointBytes, folded, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

type Mode = typeof EDIT | typeof DIRECT;

const TEXT = p(t('Cloud and serverless.')) + p(t('After.'));

/** A cell of a fixed width, holding one paragraph. */
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

const TABLE =
  p(t('Before.')) +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  p(t('After.'));

/** A format of words of the first paragraph, as `mode` makes it. */
const format = (
  doc: OrderedDocument,
  mode: Mode,
  from: number,
  to: number,
  localName: string,
  attributes?: Record<string, string>
) =>
  ({
    kind: 'format',
    ...mode,
    from: at(doc, 0, from),
    to: at(doc, 0, to),
    properties: [attributes === undefined ? { localName } : { localName, attributes }],
  }) as const;

/** Ana's new words, and Olive's suggested bold over part of them and the words before. */
function boldOverNewWords(doc: OrderedDocument): void {
  apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
  apply(doc, 'olive:1', format(doc, EDIT, 2, 8, 'b'), 'Olive');
}

/** A write of the first paragraph's properties. */
const paragraphWrite = (
  doc: OrderedDocument,
  mode: Mode,
  localName: string,
  attributes: Record<string, string>
) =>
  ({
    kind: 'paragraph-property',
    ...mode,
    marks: [markOf(doc, 0)],
    properties: [{ localName, attributes }],
  }) as const;

/** A write of the page setup of the first paragraph's section. */
const pageWrite = (doc: OrderedDocument, mode: Mode, page: Record<string, number | string>) =>
  ({ kind: 'section-property', ...mode, sections: [], page, pageAt: markOf(doc, 0) }) as const;

/** A section that states its portrait page, a value Word's reject of its change record drops. */
const PORTRAIT = '<w:sectPr><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/></w:sectPr>';

/** Ana's A4 page and Olive's direct Letter page over it, neither naming the orientation. */
function letterOverA4(doc: OrderedDocument): void {
  apply(doc, 'ana:1', pageWrite(doc, EDIT, { pageWidthTwips: 11906, pageHeightTwips: 16838 }));
  const letter = { pageWidthTwips: 12240, pageHeightTwips: 15840 };
  apply(doc, 'olive:1', pageWrite(doc, DIRECT, letter), 'Olive');
}

/** The first container of `token`, as a table write names it. */
const first = (doc: OrderedDocument, token: string) => [addressOf(opens(doc, token)[0]!)];

/**
 * Ben's indent, his Enter at the end of its paragraph, which copies the record to his new mark,
 * and his press on that mark, with no change record of its own, each on its own day.
 */
function pressOverOwnRecord(doc: OrderedDocument): void {
  const day = (n: number) => `2026-10-2${n}T09:30:00Z`;
  apply(doc, 'ben:1', paragraphWrite(doc, EDIT, 'ind', { left: '720' }), 'Ben', day(1));
  apply(doc, 'ben:2', { kind: 'split', ...EDIT, place: { mark: markOf(doc, 0) } }, 'Ben', day(2));
  apply(doc, 'ben:3', paragraphWrite(doc, EDIT, 'jc', { val: 'right' }), 'Ben', day(3));
}

const PRESS = "a press on its author's own new paragraph, over their pending record";

/** Olive's decision of the card that holds `suggestion`. */
const decideCard = (doc: OrderedDocument, suggestion: string, status: 'accepted' | 'rejected') =>
  apply(
    doc,
    'olive:1',
    { kind: 'decide', cards: [{ suggestions: orderedCardSeen(doc, suggestion) ?? [] }], status },
    'Olive'
  );

/** Each room: its file, the edits, and the suggestion taken back with its author. */
const ROOMS: readonly (readonly [string, string, (doc: OrderedDocument) => void, string])[] = [
  [
    "a direct format over another's pending format",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'olive:2', format(doc, DIRECT, 6, 8, 'i'), 'Olive');
    },
    'Olive',
  ],
  [
    "a direct format on words of another's pending format, outside the new words",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'olive:2', format(doc, DIRECT, 2, 4, 'i'), 'Olive');
    },
    'Olive',
  ],
  [
    "a format on its author's own new words, over another's pending format",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'ana:2', format(doc, EDIT, 6, 8, 'i'));
    },
    'Olive',
  ],
  [
    "a direct paragraph write over another's pending one",
    TEXT,
    (doc) => {
      apply(doc, 'ana:1', paragraphWrite(doc, EDIT, 'jc', { val: 'center' }));
      apply(doc, 'olive:1', paragraphWrite(doc, DIRECT, 'ind', { left: '720' }), 'Olive');
    },
    'Ana',
  ],
  [PRESS, TEXT, pressOverOwnRecord, 'Ben'],
  [
    "a direct page write over another's pending one",
    TEXT,
    (doc) => {
      apply(doc, 'ana:1', pageWrite(doc, EDIT, { marginLeftTwips: 1800 }));
      const margins = { marginTopTwips: 720, marginBottomTwips: 720 };
      apply(doc, 'olive:1', pageWrite(doc, DIRECT, margins), 'Olive');
    },
    'Ana',
  ],
  // The section before Ana's write states its portrait page, and the writes keep it
  [
    "a direct page size over another's pending one, in a paragraph's portrait section",
    p(t('Cloud and serverless.'), PORTRAIT) + p(t('After.')),
    letterOverA4,
    'Ana',
  ],
  [
    "a direct page size over another's pending one, in a portrait body",
    TEXT + PORTRAIT,
    letterOverA4,
    'Ana',
  ],
  [
    "a direct table write over another's pending one",
    TABLE,
    (doc) => {
      const op = (justification: 'center' | 'right') =>
        ({ op: 'setTableProperties', justification }) as const;
      apply(doc, 'ana:1', {
        kind: 'table-property',
        ...EDIT,
        on: first(doc, 'tbl'),
        op: op('center'),
      });
      const right = {
        kind: 'table-property',
        ...DIRECT,
        on: first(doc, 'tbl'),
        op: op('right'),
      } as const;
      apply(doc, 'olive:1', right, 'Olive');
    },
    'Ana',
  ],
  [
    "a direct cell fill over another's pending one",
    TABLE,
    (doc) => {
      const op = (value: string) =>
        ({ op: 'setTableCellFill', color: { kind: 'hex', value } }) as const;
      apply(doc, 'ana:1', {
        kind: 'table-property',
        ...EDIT,
        on: first(doc, 'tc'),
        op: op('FFCC00'),
      });
      const fill = {
        kind: 'table-property',
        ...DIRECT,
        on: first(doc, 'tc'),
        op: op('00CCFF'),
      } as const;
      apply(doc, 'olive:1', fill, 'Olive');
    },
    'Ana',
  ],
  [
    "a direct format between two pending ones of one run's record",
    TEXT,
    (doc) => {
      const color = (val: string) => ({ val });
      apply(doc, 'ben:1', format(doc, EDIT, 0, 5, 'color', color('FF0000')), 'Ben');
      apply(doc, 'olive:1', format(doc, DIRECT, 0, 5, 'color', color('0000FF')), 'Olive');
      apply(doc, 'ana:1', format(doc, EDIT, 0, 5, 'color', color('00FF00')));
    },
    'Ben',
  ],
  [
    "a direct paragraph write between two pending ones of the mark's record",
    TEXT,
    (doc) => {
      apply(doc, 'ben:1', paragraphWrite(doc, EDIT, 'jc', { val: 'center' }), 'Ben');
      apply(doc, 'olive:1', paragraphWrite(doc, DIRECT, 'jc', { val: 'right' }), 'Olive');
      apply(doc, 'ana:1', paragraphWrite(doc, EDIT, 'jc', { val: 'both' }));
    },
    'Ben',
  ],
  [
    "a direct page write between two pending ones of the section's record",
    TEXT,
    (doc) => {
      apply(doc, 'ben:1', pageWrite(doc, EDIT, { marginLeftTwips: 1800 }), 'Ben');
      apply(doc, 'olive:1', pageWrite(doc, DIRECT, { orientation: 'landscape' }), 'Olive');
      apply(doc, 'ana:1', pageWrite(doc, EDIT, { orientation: 'portrait' }));
    },
    'Ben',
  ],
];

describe('a write made between two pending ones of one change record', () => {
  // Word shows the latest write on top: the later pending orientation, not the direct one
  test('the later pending write shows on top of the direct one', () => {
    const { document: doc } = openOrderedWordFile(docx(TEXT), { file: 1 });
    const [, , edit] = ROOMS.find(([name]) => name.startsWith('a direct page write between'))!;
    edit(doc);
    const shown = /<w:sectPr>(.*?)<w:sectPrChange /.exec(partsOf(doc)[0] ?? '')?.[1] ?? '';
    expect(shown).not.toContain('landscape');
  });
});

// After a fold the pending page write is the file's record, which states the section before it:
// a direct write over it keeps that statement in the record, live and after another fold
describe('a direct page write after a fold, over a pending one', () => {
  test('a portrait page over a landscape one', async () => {
    const opened = openOrderedWordFile(docx(TEXT), { file: 1 });
    const landscape = pageWrite(opened.document, EDIT, { orientation: 'landscape' });
    apply(opened.document, 'ben:1', landscape, 'Ben');
    const blobs = mediaStore(opened.base);
    const { document: doc, base } = await folded(opened.document, opened.base, 2, blobs);
    apply(doc, 'olive:1', pageWrite(doc, DIRECT, { orientation: 'portrait' }), 'Olive');
    await checkFold(doc, base, blobs, 'a direct portrait page after a fold');
  });

  // Ben's margins take his landscape record over: the taken record keeps the file's record
  test('a portrait page over a landscape one another suggestion took over', async () => {
    const opened = openOrderedWordFile(docx(TEXT), { file: 1 });
    const landscape = pageWrite(opened.document, EDIT, { orientation: 'landscape' });
    apply(opened.document, 'ben:1', landscape, 'Ben');
    const blobs = mediaStore(opened.base);
    const { document: doc, base } = await folded(opened.document, opened.base, 2, blobs);
    const margins = { marginTopTwips: 720, marginBottomTwips: 720 };
    apply(doc, 'ben:2', pageWrite(doc, EDIT, margins), 'Ben');
    apply(doc, 'olive:1', pageWrite(doc, DIRECT, { orientation: 'portrait' }), 'Olive');
    await checkFold(doc, base, blobs, 'a direct portrait page over a taken record');
  });
});

// After a fold the pending page write is the file's record; a direct break copies it, and the
// type a first break gave the section goes with that record's reject on the copy too, as Word's
describe('a direct section break after a fold, in a section with a pending page write', () => {
  test('two continuous breaks', async () => {
    const opened = openOrderedWordFile(docx(TEXT + p(t('Signed.'))), { file: 1 });
    apply(opened.document, 'ana:1', pageWrite(opened.document, EDIT, { marginTopTwips: 720 }));
    const blobs = mediaStore(opened.base);
    const { document: doc, base } = await folded(opened.document, opened.base, 2, blobs);
    for (const [id, k] of [
      ['olive:1', 0],
      ['olive:2', 1],
    ] as const) {
      const mark = markOf(doc, k);
      apply(doc, id, { kind: 'section-break', ...DIRECT, mark, type: 'continuous' }, 'Olive');
    }

    await checkFold(doc, base, blobs, 'two continuous breaks after a fold');
  });
});

describe('a write made over a pending record, the record taken back after a fold', () => {
  for (const [label, body, edit, author] of ROOMS) {
    test(label, async () => {
      const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
      edit(doc);
      const blobs = mediaStore(base);
      await checkFold(doc, base, blobs, label);

      // The author's first suggestion is the record the others were made over
      const takenBack = `S:${author.toLowerCase()}:1.1`;
      await checkFold(doc, base, blobs, `${label}, taken back`, (each) => {
        apply(
          each,
          `${author.toLowerCase()}:9`,
          { kind: 'retract', suggestions: [takenBack] },
          author
        );
      });
    });
  }
});

// Seed 101 of the fold shapes: a reject of the indent alone takes Ben's press away too, live as
// after a fold
test("a press on its author's own new paragraph, its record rejected alone", async () => {
  const { document: doc, base } = openOrderedWordFile(docx(TEXT), { file: 1 });
  const blobs = mediaStore(base);
  pressOverOwnRecord(doc);
  await checkFold(doc, base, blobs, 'a press on its own new paragraph');
  decideCard(doc, 'S:ben:1.1', 'rejected');
  const once = await folded(doc, base, 2, blobs);
  expect([doc, once.document].map((copy) => partsOf(copy)[0])).toEqual([
    expect.not.stringContaining('<w:jc '),
    expect.not.stringContaining('<w:jc '),
  ]);
});

// A reject of Ana's italic takes Olive's color made over it away too, and the room draws the words
// in one run again; a fold splits no run there, as the color splits it only while Ana's shows
test("a direct format over another's pending format, the format rejected: one run", async () => {
  const opened = openOrderedWordFile(docx(TEXT), { file: 1 });
  const blobs = mediaStore(opened.base);
  apply(opened.document, 'ana:1', format(opened.document, EDIT, 0, 5, 'i'));
  const red = format(opened.document, DIRECT, 1, 4, 'color', { val: 'C00000' });
  apply(opened.document, 'olive:1', red, 'Olive');
  const once = await folded(opened.document, opened.base, 2, blobs);
  const twice = await folded(once.document, once.base, 3, blobs);
  const runs = [opened, once, twice].map((copy) => {
    decideCard(copy.document, 'S:ana:1.1', 'rejected');
    return [...(partsOf(copy.document)[0] ?? '').matchAll(/<w:r>.*?<\/w:r>/g)].map((m) => m[0]);
  });
  const one = ['<w:r><w:t>Cloud and serverless.</w:t></w:r>', '<w:r><w:t>After.</w:t></w:r>'];
  expect(runs).toEqual([one, one, one]);
});

// The press is decided with Ben's new paragraph: after its accept, his indent still pending, a
// fold keeps the press as the room holds it, and the checkpoint opens
test("a press on its author's own new paragraph, the paragraph accepted", async () => {
  const { document: doc, base } = openOrderedWordFile(docx(TEXT), { file: 1 });
  const blobs = mediaStore(base);
  pressOverOwnRecord(doc);
  decideCard(doc, 'S:ben:2.1', 'accepted');
  await checkFold(doc, base, blobs, 'a press on its accepted new paragraph');
});

const MANIFEST = 'ordered-checkpoint.json';
const FORMAT = "a direct format over another's pending format";
const PARAGRAPH = "a direct paragraph write over another's pending one";
const PAGE = "a direct page write over another's pending one";

/** A manifest's entry of made-over records, as JSON. */
interface Entry {
  readonly label: number;
  readonly target: string;
  readonly records: readonly Record<string, unknown>[];
}

/** A room's checkpoint of version 5 with its manifest's text edited. */
async function crafted(label: string, edit: (json: string) => string): Promise<Uint8Array> {
  const [, body, change] = ROOMS.find(([name]) => name === label)!;
  const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
  change(doc);
  const blobs = mediaStore(base);
  const entries = unzipSync(
    await checkpointBytes(doc, base, blobs, writeOrderedCheckpointVersion5)
  );
  entries[MANIFEST] = strToU8(edit(strFromU8(entries[MANIFEST]!)));
  return zipSync(entries);
}

/** A room's checkpoint with its manifest's made-over records edited. */
const withMadeOver = (label: string, edit: (entries: Entry[]) => unknown) =>
  crafted(label, (json) => {
    const manifest = JSON.parse(json) as { madeOver: Entry[] };
    return JSON.stringify({ ...manifest, madeOver: edit(manifest.madeOver) });
  });

/** Each entry's records made over a pending one edited. */
const withMade =
  (edit: (record: Record<string, unknown>) => Record<string, unknown>) => (entries: Entry[]) =>
    entries.map((entry) => ({
      ...entry,
      records: entry.records.map((record) => ('over' in record ? edit(record) : record)),
    }));

/** The reason a checkpoint does not open, or `opens`. */
function openReason(bytes: Uint8Array): string {
  try {
    openOrderedCheckpoint(bytes, { file: 2 });
    return 'opens';
  } catch (error) {
    return error instanceof OrderedWordFileError ? error.reason : String(error);
  }
}

describe("a version 5 checkpoint's made-over records are untrusted", () => {
  test('a checkpoint with none opens as the file holds it', async () => {
    for (const label of [FORMAT, PARAGRAPH, PAGE]) {
      const bytes = await crafted(label, (json) => {
        const { madeOver: _none, ...rest } = JSON.parse(json) as Record<string, unknown>;
        return JSON.stringify(rest);
      });
      expect({ label, reason: openReason(bytes) }).toEqual({ label, reason: 'opens' });
    }
  });

  test('each forged entry refuses the whole checkpoint', async () => {
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const forged: [string, Promise<Uint8Array>][] = [
      [
        'a label on no element',
        withMadeOver(FORMAT, (list) => list.map((e) => ({ ...e, label: 999 }))),
      ],
      ['a change record named twice', withMadeOver(FORMAT, (list) => [...list, ...list])],
      ['no records', withMadeOver(FORMAT, (list) => list.map((e) => ({ ...e, records: [] })))],
      [
        "another target than the label's change record",
        withMadeOver(FORMAT, (list) =>
          list.map((e) => ({
            ...e,
            target: 'pPr',
            records: e.records.map(({ format: _format, element: _element, ...r }) => r),
          }))
        ),
      ],
      [
        "another target's write",
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, sectionOp: { op: 'type', type: 'continuous' } }))
        ),
      ],
      [
        'the properties of another element than its target',
        withMadeOver(
          FORMAT,
          withMade(({ format: _format, ...r }) => ({ ...r, element: `<w:pPr xmlns:w="${W}"/>` }))
        ),
      ],
      [
        'a write beside an element',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, element: `<w:rPr xmlns:w="${W}"/>` }))
        ),
      ],
      [
        'children a run record sets',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, merge: [] }))
        ),
      ],
      [
        "a table's cells",
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, tableCells: [0] }))
        ),
      ],
      [
        'a suggestion the list does not hold',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, suggestion: 99 }))
        ),
      ],
      [
        'made over a suggestion the list does not hold',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, over: [99] }))
        ),
      ],
      [
        'taken from a suggestion the list does not hold',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, replaced: 0, replacedEarlier: [99] }))
        ),
      ],
      [
        'taken earlier from a suggestion, and not after',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, replacedEarlier: [0] }))
        ),
      ],
      [
        'a record on its own new words of a decided suggestion',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, untracked: true }))
        ),
      ],
      [
        'a property no format writes',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, format: { properties: [{ localName: 'pStyle' }] } }))
        ),
      ],
      [
        'a field a property does not have',
        withMadeOver(
          FORMAT,
          withMade((r) => ({ ...r, format: { properties: [{ localName: 'i', val: 1 }] } }))
        ),
      ],
      [
        'a list number that is no number',
        withMadeOver(
          PARAGRAPH,
          withMade((r) => ({ ...r, paragraphOp: { op: 'setListNumbering', numId: 'x' } }))
        ),
      ],
      // Ben's press on his own new paragraph read as a direct one: his suggestion is pending
      [
        'a press on its own new paragraph as a direct record',
        withMadeOver(
          PRESS,
          withMade(({ untracked: _untracked, ...r }) => r)
        ),
      ],
      [
        "a margin past the engine's bound",
        withMadeOver(
          PAGE,
          withMade((r) => ({ ...r, sectionOp: { op: 'page', marginTopTwips: 99_999_999 } }))
        ),
      ],
      [
        'a key that reaches a prototype',
        crafted(FORMAT, (json) => json.replace('"over":', '"__proto__":{},"over":')),
      ],
    ];
    for (const [name, bytes] of forged) {
      expect({ name, reason: openReason(await bytes) }).toEqual({
        name,
        reason: 'checkpoint-manifest',
      });
    }
  });
});
