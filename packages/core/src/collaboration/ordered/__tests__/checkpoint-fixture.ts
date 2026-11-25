// Rooms whose checkpoints hold a fact in every section of the manifest (design D14), and who
// typed each kind of unit: the checkpoint tests and the golden checkpoints use them.

import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { orderedSeenCards } from '../change/apply-decide.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import type { OrderedDocument } from '../model/types.ts';
import { orderedBlobHash } from '../protocol/blobs.ts';
import {
  addressOf,
  apply,
  APPLY_FILE,
  at,
  BODY,
  DIRECT,
  EDIT,
  markOf,
  opens,
  textOf,
} from './change-apply-fixture.ts';
import { mediaStore } from './fold-check.ts';
import { STYLED } from './style-revision-fixture.ts';
import { docx, M, p, RED_PNG, t } from './word-fixtures.ts';

/** The index of the shown body paragraph that starts with `text`. */
export function paragraph(doc: OrderedDocument, text: string): number {
  for (let k = 0; textOf(doc, k) !== '<none>'; k += 1) {
    if (textOf(doc, k).startsWith(text)) {
      return k;
    }
  }

  throw new Error(`no paragraph ${text}`);
}

/**
 * Changes to `APPLY_FILE` that leave a fact of every section past its first paragraph and first
 * cell: a takeover (Pia's press over Quin's), a row height in Ana's new table (an untracked record
 * with its table's cells), Ben's control value (an untracked record of `w:sdtPr`), Olive's words
 * typed inside Ana's insertion (typers), Ben's comment, and Pia's value over Ben's accepted one.
 * Each is found by what it shows.
 */
export function withRoomFacts(doc: OrderedDocument): OrderedDocument {
  // Each takes over the records before it: Rae's holds Quin's by way of Pia's
  for (const [id, who, property] of [
    ['quin:1', 'Quin', { localName: 'jc', attributes: { val: 'center' } }],
    ['pia:1', 'Pia', { localName: 'spacing', attributes: { after: '240' } }],
    ['rae:1', 'Rae', { localName: 'keepNext' }],
  ] as const) {
    const marks = [markOf(doc, paragraph(doc, 'Signed'))];
    apply(doc, id, { kind: 'paragraph-property', ...EDIT, marks, properties: [property] }, who);
  }

  const due = () => paragraph(doc, 'Due');
  const table = { rows: 1, columns: 2, columnWidth: 3000 };
  apply(doc, 'ana:1', { kind: 'insert-table', ...EDIT, place: at(doc, due(), 0), ...table });
  const row = addressOf(opens(doc, 'tr').at(-1)!);
  const height = { op: 'setTableRowHeight', heightTwips: 480 } as const;
  apply(doc, 'ana:2', { kind: 'table-property', ...EDIT, on: [row], op: height });
  apply(doc, 'ana:3', { kind: 'insert', ...EDIT, place: at(doc, due(), 4), text: 'new words' });
  apply(
    doc,
    'olive:1',
    { kind: 'insert', ...DIRECT, place: at(doc, due(), 7), text: 'XY' },
    'Olive'
  );
  const control = addressOf(opens(doc, 'sdt').at(-1)!);
  const value = { kind: 'text', text: 'Beta Ltd' } as const;
  apply(doc, 'ben:1', { kind: 'control-value', ...EDIT, control, value }, 'Ben');
  const [from, to] = [at(doc, due(), 0), at(doc, due(), 3)];
  apply(doc, 'ben:2', { kind: 'comment', part: BODY, from, to, text: 'Which date?' }, 'Ben');

  // Ben's value accepted and Pia's pending: the untracked record names a decided suggestion
  const cards = orderedSeenCards(doc, ['S:ben:1.1']);
  apply(doc, 'olive:9', { kind: 'decide', cards, status: 'accepted' }, 'Olive');
  const again = { kind: 'text', text: 'Gamma Ltd' } as const;
  apply(doc, 'pia:2', { kind: 'control-value', ...EDIT, control, value: again }, 'Pia');
  return doc;
}

/** A room's Word file and the changes made on it. */
export interface CheckpointRoom {
  readonly name: string;
  readonly source: Uint8Array;
  readonly edit: (doc: OrderedDocument) => void;
}

/** The picture the rooms insert, as a change names it. */
const PNG_MEDIA = {
  sha256: orderedBlobHash(RED_PNG),
  contentType: 'image/png',
  size: RED_PNG.byteLength,
} as const;

/** A room's storage: the media of its Word file, and the picture the rooms insert. */
export function roomBlobs(base: OoxmlPackage): (sha256: string) => Uint8Array | undefined {
  const store = mediaStore(base);
  return (sha256) => (sha256 === PNG_MEDIA.sha256 ? RED_PNG : store(sha256));
}

/** The address of the first item drawn from an element of this local name. */
const itemOf = (doc: OrderedDocument, localName: string) =>
  addressOf(
    [...doc.stories.get(BODY)!.tree.items()].find(
      (item) => item.node?.kind !== 'textValue' && item.node?.localName === localName
    )!
  );

/**
 * Units Olive types inside pending insertions, drawn in runs: text, a tab, a line break, a
 * picture and a complex field inside Ana's words, a note reference (and its note), and words in
 * Ana's new cell and new text box (both of its copies).
 */
export const TYPED_RUNS: CheckpointRoom = {
  name: 'typed runs',
  source: APPLY_FILE,
  edit: (doc) => {
    const due = paragraph(doc, 'Due');
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, due, 4), text: 'new words' });

    // Each between two of Ana's letters, from the last, so the places before stay
    const object = {
      type: 'picture',
      media: PNG_MEDIA,
      extent: { cx: 914_400, cy: 457_200 },
    } as const;
    for (const [id, letter, body] of [
      ['olive:1', 12, { kind: 'insert', ...DIRECT, text: 'XY' }],
      ['olive:2', 11, { kind: 'insert-tab', ...DIRECT }],
      ['olive:3', 10, { kind: 'insert-break', ...DIRECT, break: 'line' }],
      ['olive:4', 9, { kind: 'insert-object', ...DIRECT, object }],
      ['olive:5', 8, { kind: 'insert-field', ...DIRECT, field: 'PAGE' }],
      ['olive:6', 6, { kind: 'insert-note', ...EDIT, note: 'footnote', text: ' N' }],
    ] as const) {
      apply(doc, id, { ...body, place: at(doc, due, letter) }, 'Olive');
    }

    // The new cell right of the first, and the text box's story
    const cell = itemOf(doc, 'tc');
    apply(doc, 'ana:2', { kind: 'insert-column', ...EDIT, cell, where: 'right' });
    apply(doc, 'olive:7', { kind: 'insert', ...DIRECT, place: at(doc, 5, 0), text: 'XY' }, 'Olive');
    const box = { type: 'textbox', extent: { cx: 1_270_000 } } as const;
    apply(doc, 'ana:3', { kind: 'insert-object', ...EDIT, place: at(doc, 0, 2), object: box });
    const part = [...doc.stories.keys()].find(isTextBoxStoryName)!;
    const place = at(doc, 0, 0, part);
    apply(doc, 'olive:8', { kind: 'insert', part, mode: 'edit', place, text: 'XY' }, 'Olive');
  },
};

/**
 * Units Olive types drawn bare, or hidden by their run: a simple field's new code, an equation
 * (an element kept whole is drawn as one), and words in hidden text.
 */
export const TYPED_BARE: CheckpointRoom = {
  name: 'typed bare units',
  source: docx(
    p(t('Page ') + `<w:fldSimple w:instr=" PAGE ">${t('1')}</w:fldSimple>`) +
      p(t('Sum ') + `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>`) +
      p(t('Hidden', '<w:vanish/>'))
  ),
  edit: (doc) => {
    const field = itemOf(doc, 'fldSimple');
    apply(doc, 'olive:1', { kind: 'field-code', ...EDIT, field, code: 'NUMPAGES' }, 'Olive');
    const object = itemOf(doc, 'oMath');
    const equation = { kind: 'equation', ...EDIT, object, equation: 0, linear: 'x^2' } as const;
    apply(doc, 'olive:2', equation, 'Olive');
    apply(doc, 'olive:3', { kind: 'insert', ...EDIT, place: at(doc, 2, 2), text: 'XY' }, 'Olive');
  },
};

/** The rooms of `withRoomFacts`, and of STYLED with Ana's insertion accepted (`S1`). */
export const ROOM_FACTS: CheckpointRoom = {
  name: 'room facts',
  source: APPLY_FILE,
  edit: (doc) => {
    withRoomFacts(doc);
  },
};

export const STYLE_RECORDS: CheckpointRoom = {
  name: 'style records',
  source: STYLED,
  edit: (doc) => {
    const cards = orderedSeenCards(doc, ['S1']);
    apply(doc, 'olive:1', { kind: 'decide', cards, status: 'accepted' }, 'Olive');
  },
};

/** A suggestion inside a run: the checkpoint numbers the pieces of the room's run (`oc:run`). */
export const SPLIT_RUN: CheckpointRoom = {
  name: 'split room run',
  source: docx(
    p('<w:r w:rsidR="00A96A35"><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Cloud.</w:t></w:r>')
  ),
  edit: (doc) => {
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 0, 3), text: 'u' });
  },
};

/**
 * Whole paragraphs pasted with their list and style: the checkpoint gives each back to the paste
 * (sections `lists` and `styles`).
 */
export const PASTED_LIST: CheckpointRoom = {
  name: 'pasted list and style',
  source: docx(p(t('Cloud.'))),
  edit: (doc) => {
    const pPr = {
      e: 'w:pPr',
      c: [
        { e: 'w:pStyle', a: { 'w:val': 'PastedClause' } },
        {
          e: 'w:numPr',
          c: [
            { e: 'w:ilvl', a: { 'w:val': '0' } },
            { e: 'w:numId', a: { 'w:val': '#0' } },
          ],
        },
      ],
    };
    const level = {
      e: 'w:lvl',
      a: { 'w:ilvl': '0' },
      c: [
        { e: 'w:numFmt', a: { 'w:val': 'lowerLetter' } },
        { e: 'w:lvlText', a: { 'w:val': '(%1)' } },
      ],
    };
    const style = {
      e: 'w:style',
      a: { 'w:type': 'paragraph', 'w:styleId': 'PastedClause' },
      c: [{ e: 'w:name', a: { 'w:val': 'Pasted Clause' } }],
    };
    apply(doc, 'ana:1', {
      kind: 'insert-fragment',
      ...EDIT,
      place: at(doc, 0, 0),
      blocks: [{ e: 'w:p', c: [pPr, { e: 'w:r', c: [{ e: 'w:t', c: ['Pasted.'] }] }] }],
      whole: true,
      lists: [{ abstract: { e: 'w:abstractNum', c: [level] } }],
      styles: [style],
    });
  },
};

/** A table cell of the given text, 3000 twips wide. */
const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** Cy's insertion in the file, with an attribute past its triple, as Word 2023 writes it. */
const DATED_INSERTION =
  '<w:ins w:id="3" w:author="Cy" w:date="2026-01-01T00:00:00Z" ' +
  'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" ' +
  `w16du:dateUtc="2026-01-01T00:00:00Z">${t(' later')}</w:ins>`;

/**
 * Table records a change record holds (section `changeRecords`: `merge`, `tableCells`, a record
 * that took another's, `replaced`), with Ana's column's lent widths (`untracked`, a cell's
 * record that sets only what it changed, `recordUnchanged`), Ben's and Pia's control values, and
 * a file suggestion's attribute past its triple (`suggestions[].attributes`).
 */
export const TABLE_RECORDS: CheckpointRoom = {
  name: 'table change records',
  source: docx(
    p(t('Fees') + DATED_INSERTION) +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
      '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${cell('A1')}${cell('B1')}</w:tr><w:tr>${cell('A2')}${cell('B2')}</w:tr></w:tbl>` +
      p(
        t('Signed by ') +
          `<w:sdt><w:sdtPr><w:id w:val="5"/><w:text/></w:sdtPr><w:sdtContent>${t('Acme')}` +
          '</w:sdtContent></w:sdt>'
      )
  ),
  edit: (doc) => {
    const tc = (k: number) => addressOf(opens(doc, 'tc')[k]!);
    apply(doc, 'ana:3', { kind: 'insert-column', ...EDIT, cell: tc(1), where: 'right' });
    const fill = { op: 'setTableCellFill', color: { kind: 'hex', value: 'FFCC00' } } as const;
    apply(doc, 'ana:1', { kind: 'table-property', ...EDIT, on: [tc(0)], op: fill });
    const width = {
      op: 'setTableProperties',
      preferredWidth: { type: 'pct', value: 5000 },
    } as const;
    const table = addressOf(opens(doc, 'tbl')[0]!);
    apply(doc, 'ana:2', { kind: 'table-property', ...EDIT, on: [table], op: width });
    const control = addressOf(opens(doc, 'sdt')[0]!);
    for (const [id, who, text] of [
      ['ben:1', 'Ben', 'Beta Ltd'],
      ['pia:1', 'Pia', 'Gamma Ltd'],
    ] as const) {
      apply(
        doc,
        id,
        { kind: 'control-value', ...EDIT, control, value: { kind: 'text', text } },
        who
      );
    }
  },
};
