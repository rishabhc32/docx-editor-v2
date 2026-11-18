// A paragraph after a block the list draws nothing of, with real editors and one ordered room.
// The list leaves out a cell its author took back, and a note whose reference
// shows nowhere, with every item in them. A paragraph drawn alone (the publish of a write, and
// the place an edit reads) took the items of such a block before it as its own: text typed into
// the next cell went into the taken-back cell, where the list shows nothing, and a note after a
// left-out note showed that note's markers. Every editor shows what its list composes.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, MATRIX_EXTRAS, p, R, t, W } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  listTexts,
  listXml,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { caretIn, converged, tableCommand } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

const cell = (body: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${body}</w:tc>`;

/** A table of one row: a cell with `Fee`, and an empty cell. */
const EMPTY_CELL = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell(p(t('Fee')))}${cell('<w:p/>')}</w:tr></w:tbl>` +
    p(t('Signed here.'))
);

const ENDNOTES = '/word/endnotes.xml';
const reference = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${id}"/></w:r>`;
const endnote = (id: number, runs: string) =>
  `<w:endnote w:id="${id}"><w:p>${runs}</w:p></w:endnote>`;

/** Two endnotes; the first holds a bookmark, which a deletion leaves where it is. */
const TWO_NOTES = docx(
  p(t('The fee') + reference(1) + t(' is due.')) + p(t('Paid') + reference(2) + t(' in 30 days.')),
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/endnotes.xml':
        `<w:endnotes xmlns:w="${W}" xmlns:r="${R}">` +
        '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' +
        '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>' +
        endnote(
          1,
          '<w:r><w:endnoteRef/></w:r><w:bookmarkStart w:id="0" w:name="annex"/>' +
            t(' See the annex.') +
            '<w:bookmarkEnd w:id="0"/>'
        ) +
        endnote(2, '<w:r><w:endnoteRef/></w:r>' + t(' Net of tax.')) +
        '</w:endnotes>',
    },
  }
);

/** The endnotes an editor shows, as XML. */
function shownEndnotes(person: Person): string {
  const part = surface(person).session.currentPackage().parts.get(ENDNOTES);
  return part ? serializeOoxmlPart(part) : 'no endnotes';
}

describe('a paragraph after a block the list leaves out', () => {
  test('typing in the cell after a taken-back new cell goes into that cell', async () => {
    const room = localRoom(EMPTY_CELL, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: EMPTY_CELL, review });
    const olive = await join(room, 'Olive', {
      wordFile: EMPTY_CELL,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    // Ana's new column between the two, taken back: its cell stays in the list, hidden
    caretIn(ana, 'Fee');
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    surface(ana).undo();
    await settle();
    converged(room, ana, olive);
    expect(editorTexts(ana)).toEqual(['The fees are below.', 'Fee', '', 'Signed here.']);
    // Ana types in the empty cell, at its paragraph mark
    caretAt(ana, 2, 0);
    surface(ana).insertPlainText('€ 1.200,00 ');
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(listTexts(person)).toEqual([
        'The fees are below.',
        'Fee',
        '€ 1.200,00 ',
        'Signed here.',
      ]);
      expect(editorTexts(person)).toEqual(listTexts(person));
    }
    // The room's document and a save hold Ana's words in the cell that shows, and no third cell
    const inCell =
      /<w:tc><w:tcPr>(?:(?!<\/w:tc>).)*<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t xml:space="preserve">€ 1\.200,00 <\/w:t><\/w:r><\/w:ins><\/w:p><\/w:tc><\/w:tr>/;
    const roomXml = serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part);
    const read = readOoxmlPackage(new Uint8Array(await olive.editor.save()), {});
    if (!read.ok) throw new Error(read.reason);
    const saved = serializeOoxmlPart(read.package.parts.get(BODY)!);
    for (const xml of [roomXml, saved]) {
      expect(xml).toMatch(inCell);
      expect(xml.match(/<w:tc>/g)).toHaveLength(2);
    }
  });

  test('typing in the note after a deleted note shows nothing of the deleted note', async () => {
    const room = localRoom(TWO_NOTES, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: TWO_NOTES, review });
    const olive = await join(room, 'Olive', {
      wordFile: TWO_NOTES,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    const deleted = olive.editor.exec({ type: 'deleteNote', noteKind: 'endnote', noteId: 1 });
    expect(deleted.ok).toBe(true);
    await settle();
    for (const person of [ana, olive])
      expect(shownEndnotes(person)).toBe(listXml(person, ENDNOTES));
    expect(surface(ana).enterNote('endnote:2')).toBe(true);
    surface(ana).insertPlainText('Due. ');
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(shownEndnotes(person)).toBe(listXml(person, ENDNOTES));
      expect(shownEndnotes(person)).not.toContain('annex');
      expect(shownEndnotes(person)).toContain('Due. ');
    }
  });
});
