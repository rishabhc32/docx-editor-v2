// A write made over a pending record with no change record of its own (a direct one, or one on
// its author's own new words) stays when the record's author takes the record back, after a fold
// too: Word keeps one change record, so the checkpoint keeps the room's records (design D14).

import { describe, test } from 'bun:test';
import { openOrderedWordFile } from '../client/open.ts';
import type { OrderedDocument } from '../model/types.ts';
import { addressOf, apply, DIRECT, EDIT, opens } from './change-apply-fixture.ts';
import { checkFold, mediaStore } from './fold-check.ts';
import { docx, p, t } from './word-fixtures.ts';

/** A cell of a fixed width, holding one paragraph. */
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

const TABLE =
  p(t('Before.')) +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('A')}${cell('B')}</w:tr></w:tbl>` +
  p(t('After.'));

/** The first container of `token`, as a table write names it. */
const first = (doc: OrderedDocument, token: string) => [addressOf(opens(doc, token)[0]!)];

/** Each room: its file, the edits, and the suggestion taken back with its author. */
const ROOMS: readonly (readonly [string, string, (doc: OrderedDocument) => void, string])[] = [
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
