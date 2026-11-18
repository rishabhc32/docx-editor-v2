// Edits whose new words are the engine's copy of words the view shows, while a party's pending
// move is there (the new-contracts acceptance run xdemo-short, seed 1014). The organiser edited a
// text form field inside a plain text control, into which Bob had moved one character. The edit
// took the whole control as the field's copy, so its words showed twice; the copy held Bob's
// move range markers, which the template load noted in the revision ids (a take-back kept them,
// so a copy that rebased gave other `w:id`s), and noted as the move's own range markers (so the
// save wrote Bob's halves bare, and a reopen added range markers). A new repeating item over a
// pending move took the same path through the template load.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { isHidden } from '../model/document.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord, type OrderedUndoRecord } from '../model/undo-record.ts';
import { BODY, partsOf, textOf } from './change-apply-fixture.ts';
import { snapshot } from './snapshot.ts';
import { docx, p, t } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const textField = (result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Name1"/><w:enabled/>' +
  '<w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r>' +
  '<w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r>' +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t(result)}` +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** The seed's shape: a text form field inside a plain text control, words after it. */
const FIELD_FILE = docx(
  p(t('Fee for 12 months.')) +
    p(
      t('Address: ') +
        sdt(
          '<w:alias w:val="Address"/><w:text w:multiLine="1"/>',
          t('Name: ') + textField('Nimbus Ltd') + t(' Agreed: ') + t(' Term: 1 year.')
        )
    )
);

/** A repeating section of one item. */
const REPEAT_FILE = docx(
  p(t('Fee for 12 months.')) +
    sdt(
      `<w:id w:val="10"/><w15:repeatingSection ${W15}/>`,
      sdt(`<w:id w:val="11"/><w15:repeatingSectionItem ${W15}/>`, p(t('Item one text.')))
    ) +
    p(t('End.'))
);

const at = (doc: OrderedDocument, index: number, offset: number) => {
  const view = orderedViewParagraph(doc, BODY, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, offset);
};

/** One change, as the room applies it: in an undo record, with a minter for its id. */
function applied(
  doc: OrderedDocument,
  id: string,
  body: Record<string, unknown>,
  author: string
): OrderedUndoRecord {
  const change = { v: 1, id, part: BODY, ...body } as unknown as OrderedChange;
  const ctx = {
    person: testPerson(author),
    date: '2026-10-28T09:19:10Z',
    ids: orderedIdMinter(id),
  };
  return withUndoRecord(doc, () => applyOrderedChange(doc, change, ctx)).record;
}

/** Bob's suggestion: "2" of "12" moved to a place of a shown paragraph. */
function bobMoves(doc: OrderedDocument, index: number, offset: number): void {
  applied(
    doc,
    'bob-1:1',
    {
      kind: 'move',
      mode: 'suggest',
      from: at(doc, 0, 9),
      to: at(doc, 0, 10),
      place: at(doc, index, offset),
    },
    'Bob'
  );
}

interface Case {
  readonly file: Uint8Array;
  /** Bob's move, then the organiser's edit; returns the edit. */
  readonly before: (doc: OrderedDocument) => Record<string, unknown>;
}

const CASES: Readonly<Record<string, Case>> = {
  "a text field's deletion inside a control": {
    file: FIELD_FILE,
    before: (doc) => {
      bobMoves(doc, 1, textOf(doc, 1).indexOf(' Term'));
      return {
        kind: 'form-text',
        mode: 'edit',
        field: at(doc, 1, textOf(doc, 1).indexOf('Nimbus')),
        op: { op: 'deleteText', start: 0, end: 3 },
      };
    },
  },
  'a new repeating item over a pending move': {
    file: REPEAT_FILE,
    before: (doc) => {
      bobMoves(doc, 1, 4);
      const section = [...(doc.stories.get(BODY)?.tree.items() ?? [])].find(
        (item) => item.kind === 'open' && item.token === 'sdtBlock' && !isHidden(doc, item)
      );
      if (!section) throw new Error('no repeating section');
      return {
        kind: 'repeating-item',
        mode: 'edit',
        control: { space: section.space, offset: section.offset },
        action: 'add',
      };
    },
  },
};

/** The document after Bob's move and the organiser's edit; `rebase` takes the edit back once. */
function edited(c: Case, rebase: boolean) {
  const opened = openOrderedWordFile(c.file, { file: 1 });
  const doc = opened.document;
  const edit = c.before(doc);
  const record = applied(doc, 'olive-1:1', edit, 'Olive');
  if (rebase) {
    takeBack(doc, record);
    applied(doc, 'olive-1:1', edit, 'Olive');
  }
  return { doc, base: opened.base };
}

const documentXml = (bytes: Uint8Array): string =>
  strFromU8(unzipSync(bytes)['word/document.xml'] ?? new Uint8Array());

describe('an edit that copies words around a pending move', () => {
  for (const [name, c] of Object.entries(CASES)) {
    test(`${name}: applied, taken back and applied again, equals one apply`, () => {
      const once = edited(c, false).doc;
      const again = edited(c, true).doc;
      expect(snapshot(again, false)).toEqual(snapshot(once, false));
      expect(partsOf(again)).toEqual(partsOf(once));
    });

    test(`${name}: the save reopens as the same file, the move with its range markers`, () => {
      const { doc, base } = edited(c, false);
      const saved = orderedPackageBytes(doc, base);
      const reopened = openOrderedWordFile(saved, { file: 2 });
      const xml = documentXml(saved);
      expect(documentXml(orderedPackageBytes(reopened.document, reopened.base))).toBe(xml);
      for (const marker of ['moveFromRangeStart', 'moveToRangeStart'])
        expect(xml.match(new RegExp(`<w:${marker} [^>]*w:author="Bob"`, 'g'))).toHaveLength(1);
    });
  }

  test("a text field's deletion inside a control changes only the field", () => {
    const { doc } = edited(CASES["a text field's deletion inside a control"] as Case, false);
    expect(textOf(doc, 1)).toBe('Address: Name: bus Ltd Agreed: 2 Term: 1 year.');
  });
});
