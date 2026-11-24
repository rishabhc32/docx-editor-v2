// The shapes of a suggestion on a Word file, one of each mark kind: insertions, deletions, formats,
// paragraph, section, table, row and cell records, moves, links, controls, notes, headers and
// text boxes. Ana makes each one.

import type { OrderedTableOp } from '../edits/tables.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  EDIT,
  HEADER,
  markOf,
  opens,
} from './change-apply-fixture.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, t } from './word-fixtures.ts';

const sections = docx(
  p(t('First section.'), '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>') +
    p(t('Second section.'))
);
const note = docx(bodyOf(MATRIX_ITEMS.find((item) => item.key === 'fnref')!), MATRIX_EXTRAS);
const textBox = docx(bodyOf(MATRIX_ITEMS.find((item) => item.key === 'textbox')!));
export const anas = (doc: OrderedDocument) =>
  [...doc.suggestions.values()]
    .filter((s) => s.status === 'pending' && s.attribution.author === 'Ana')
    .map((s) => s.id);
const deletion = (doc: OrderedDocument) =>
  apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 0, 4), to: at(doc, 0, 8) });
const wordsAdded = (doc: OrderedDocument, paragraph: number, offset: number) => {
  deletion(doc);
  apply(doc, 'ana:2', {
    kind: 'insert',
    ...EDIT,
    into: anas(doc)[0],
    place: at(doc, paragraph, offset),
    text: 'X',
  });
};
const table = (doc: OrderedDocument, on: string, op: OrderedTableOp) =>
  apply(doc, 'ana:1', {
    kind: 'table-property',
    ...EDIT,
    on: [addressOf(opens(doc, on)[0]!)],
    op,
  });

/** Each shape: its label, the Word file, and Ana's suggestion on it. */
export const FOLD_SHAPES = [
  ['a deletion', APPLY_FILE, deletion],
  [
    'a deletion with a date change',
    APPLY_FILE,
    (doc: OrderedDocument) => {
      deletion(doc);
      apply(doc, 'ana:2', {
        kind: 'attribution',
        ...EDIT,
        suggestions: anas(doc),
        author: 'Ana',
        date: '2026-10-30T15:59:00Z',
      });
    },
  ],
  ['a deletion with words added', APPLY_FILE, (doc: OrderedDocument) => wordsAdded(doc, 0, 8)],
  [
    'a deletion with words added in another paragraph',
    APPLY_FILE,
    (doc: OrderedDocument) => wordsAdded(doc, 5, 0),
  ],
  [
    'an insertion',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' }),
  ],
  [
    'a format',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'format',
        ...EDIT,
        from: at(doc, 0, 4),
        to: at(doc, 0, 8),
        properties: [{ localName: 'b' }],
      }),
  ],
  [
    'a paragraph press',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 0)],
        properties: [{ localName: 'jc', attributes: { val: 'center' } }],
      }),
  ],
  [
    'a section press',
    sections,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'section-property',
        ...EDIT,
        sections: [],
        page: { orientation: 'landscape' },
        pageAt: markOf(doc, 0),
      }),
  ],
  [
    'a move',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'move',
        ...EDIT,
        from: at(doc, 0, 4),
        to: at(doc, 0, 8),
        place: at(doc, 5, 0),
      }),
  ],
  [
    'a link target',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'link-target',
        ...EDIT,
        link: addressOf(opens(doc, 'hyperlink')[0]!),
        href: 'https://example.com/new',
      }),
  ],
  [
    'a control value',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'control-value',
        ...EDIT,
        control: addressOf(opens(doc, 'sdt')[0]!),
        value: { kind: 'text', text: 'Nimbus' },
      }),
  ],
  [
    'a new row',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'insert-row',
        ...EDIT,
        row: addressOf(opens(doc, 'tr')[0]!),
        where: 'below',
      }),
  ],
  [
    'a deleted row',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', { kind: 'delete-row', ...EDIT, row: addressOf(opens(doc, 'tr')[0]!) }),
  ],
  [
    'a deleted column',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'delete-column',
        ...EDIT,
        cell: addressOf(opens(doc, 'tc')[0]!),
      }),
  ],
  [
    'a table press',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      table(doc, 'tbl', { op: 'setTableProperties', justification: 'center' }),
  ],
  [
    'a row height',
    APPLY_FILE,
    (doc: OrderedDocument) => table(doc, 'tr', { op: 'setTableRowHeight', heightTwips: 600 }),
  ],
  [
    'a cell fill',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      table(doc, 'tc', { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } }),
  ],
  [
    'a deleted note',
    note,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', { kind: 'delete', ...EDIT, from: at(doc, 1, 6), to: at(doc, 1, 7) }),
  ],
  [
    'a deletion in a header',
    APPLY_FILE,
    (doc: OrderedDocument) =>
      apply(doc, 'ana:1', {
        kind: 'delete',
        part: HEADER,
        mode: 'suggest',
        from: at(doc, 0, 0, HEADER),
        to: at(doc, 0, 7, HEADER),
      }),
  ],
  [
    'a deletion in a text box',
    textBox,
    (doc: OrderedDocument) => {
      const part = [...doc.stories.keys()].find(isTextBoxStoryName)!;
      apply(doc, 'ana:1', {
        kind: 'delete',
        part,
        mode: 'suggest',
        from: at(doc, 0, 0, part),
        to: at(doc, 0, 4, part),
      });
    },
  ],
] as const;
