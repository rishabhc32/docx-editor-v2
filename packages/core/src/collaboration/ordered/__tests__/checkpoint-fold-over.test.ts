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
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, at, DIRECT, EDIT, markOf, opens } from './change-apply-fixture.ts';
import { checkFold, checkpointBytes, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

type Mode = typeof EDIT | typeof DIRECT;

const TEXT = p(t('Cloud and serverless.')) + p(t('After.'));

/** The text with its body's own section, as Word writes one. */
const SECTION =
  TEXT +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" ' +
  'w:footer="720" w:gutter="0"/></w:sectPr>';

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
const format = (doc: OrderedDocument, mode: Mode, from: number, to: number, localName: string) =>
  ({
    kind: 'format',
    ...mode,
    from: at(doc, 0, from),
    to: at(doc, 0, to),
    properties: [{ localName }],
  }) as const;

/**
 * Ana's new words, and Olive's suggested bold over the whole paragraph: each run whole, as a
 * take-back joins a run a format split only in the room (another difference).
 */
function boldOverNewWords(doc: OrderedDocument): void {
  apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 6), text: 'net ' });
  apply(doc, 'olive:1', format(doc, EDIT, 0, 25, 'b'), 'Olive');
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
const pageWrite = (doc: OrderedDocument, mode: Mode, page: Record<string, number>) =>
  ({ kind: 'section-property', ...mode, sections: [], page, pageAt: markOf(doc, 0) }) as const;

/** The first container of `token`, as a table write names it. */
const first = (doc: OrderedDocument, token: string) => [addressOf(opens(doc, token)[0]!)];

/** Each room: its file, the edits, and the suggestion taken back with its author. */
const ROOMS: readonly (readonly [string, string, (doc: OrderedDocument) => void, string])[] = [
  [
    "a direct format over another's pending format",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'olive:2', format(doc, DIRECT, 6, 10, 'i'), 'Olive');
    },
    'Olive',
  ],
  [
    "a direct format on words of another's pending format, outside the new words",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'olive:2', format(doc, DIRECT, 0, 6, 'i'), 'Olive');
    },
    'Olive',
  ],
  [
    "a format on its author's own new words, over another's pending format",
    TEXT,
    (doc) => {
      boldOverNewWords(doc);
      apply(doc, 'ana:2', format(doc, EDIT, 6, 10, 'i'));
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
  [
    "a direct page write over another's pending one",
    SECTION,
    (doc) => {
      apply(doc, 'ana:1', pageWrite(doc, EDIT, { marginLeftTwips: 1800 }));
      const margins = { marginTopTwips: 720, marginBottomTwips: 720 };
      apply(doc, 'olive:1', pageWrite(doc, DIRECT, margins), 'Olive');
    },
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
];

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

const MANIFEST = 'ordered-checkpoint.json';
const FORMAT = "a direct format over another's pending format";
const PARAGRAPH = "a direct paragraph write over another's pending one";
const PAGE = "a direct page write over another's pending one";

/** A manifest's entry of made-over records, as JSON. */
interface Entry {
  readonly label: number;
  readonly properties: string;
  readonly records: readonly Record<string, unknown>[];
}

/** A room's checkpoint with its manifest's text edited. */
async function crafted(label: string, edit: (json: string) => string): Promise<Uint8Array> {
  const [, body, change] = ROOMS.find(([name]) => name === label)!;
  const { document: doc, base } = openOrderedWordFile(docx(body), { file: 1 });
  change(doc);
  const entries = unzipSync(await checkpointBytes(doc, base, mediaStore(base)));
  entries[MANIFEST] = strToU8(edit(strFromU8(entries[MANIFEST]!)));
  return zipSync(entries);
}

/** A room's checkpoint with its manifest's made-over records edited. */
const withMadeOver = (label: string, edit: (entries: Entry[]) => unknown) =>
  crafted(label, (json) => {
    const manifest = JSON.parse(json) as { madeOver: Entry[] };
    return JSON.stringify({ ...manifest, madeOver: edit(manifest.madeOver) });
  });

/** Each entry's records edited. */
const withRecords =
  (edit: (record: Record<string, unknown>) => Record<string, unknown>) => (entries: Entry[]) =>
    entries.map((entry) => ({ ...entry, records: entry.records.map(edit) }));

/** The reason a checkpoint does not open, or `opens`. */
function openReason(bytes: Uint8Array): string {
  try {
    openOrderedCheckpoint(bytes, { file: 2 });
    return 'opens';
  } catch (error) {
    return error instanceof OrderedWordFileError ? error.reason : String(error);
  }
}

describe("a checkpoint's made-over records are untrusted", () => {
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
        "the properties of another change record than the label's",
        withMadeOver(FORMAT, (list) =>
          list.map((e) => ({
            ...e,
            properties: `<w:pPr xmlns:w="${W}"/>`,
            records: e.records.map(({ format: _format, ...r }) => ({
              ...r,
              paragraphOp: { op: 'setListLevel', level: 1 },
            })),
          }))
        ),
      ],
      [
        'a write of another kind than its properties',
        withMadeOver(
          FORMAT,
          withRecords(({ format: _format, ...r }) => ({
            ...r,
            sectionOp: { op: 'type', type: 'continuous' },
          }))
        ),
      ],
      [
        'a suggestion the list does not hold',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, suggestion: 99 }))
        ),
      ],
      [
        'made over a suggestion the list does not hold',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, over: [99] }))
        ),
      ],
      // The pending format's suggestion: a direct record is an accepted suggestion's
      [
        'a direct record of a pending suggestion',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, suggestion: (r.over as number[])[0] }))
        ),
      ],
      [
        'a record on its own new words of a decided suggestion',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, untracked: true }))
        ),
      ],
      [
        'a property no format writes',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, format: { properties: [{ localName: 'pStyle' }] } }))
        ),
      ],
      [
        'a field a property does not have',
        withMadeOver(
          FORMAT,
          withRecords((r) => ({ ...r, format: { properties: [{ localName: 'i', val: 1 }] } }))
        ),
      ],
      [
        'a list number that is no number',
        withMadeOver(
          PARAGRAPH,
          withRecords((r) => ({ ...r, paragraphOp: { op: 'setListNumbering', numId: 'x' } }))
        ),
      ],
      [
        "a margin past the engine's bound",
        withMadeOver(
          PAGE,
          withRecords((r) => ({ ...r, sectionOp: { op: 'page', marginTopTwips: 99_999_999 } }))
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
