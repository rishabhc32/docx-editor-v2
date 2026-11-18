// The heading paragraph of a table of contents ("Contents", inside the contents control, before
// the TOC field), in an ordered room: the organiser and a party can type in it, as in the editor
// alone and in Word, and so can a paragraph after the field; the rows the field generates stay
// read-only. Before, the room refused every paragraph of the control (`not-covered`, row 40).

import { afterEach, describe, expect, test } from 'bun:test';
import { applyOrderedChange } from '../../change/apply.ts';
import { orderedIdMinter } from '../../change/ids.ts';
import type { OrderedChange } from '../../change/types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedViewParagraph } from '../../compose/compose.ts';
import { orderedPlaceAt } from '../../edits/places.ts';
import { withUndoRecord } from '../../model/undo-record.ts';
import { BODY, DATE } from '../change-apply-fixture.ts';
import { docx, p, t } from '../word-fixtures.ts';
import { closeAll, join, localRoom, ORGANISER, PARTY, surface } from './editor-harness.ts';
import { act, alone, converged, roomFile, select, storyAsRead, type Act } from './links-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const row = (text: string, opening = '') =>
  p(
    opening + `<w:hyperlink w:anchor="_Toc1">${t(text)}<w:r><w:tab/></w:r>${t('1')}</w:hyperlink>`,
    '<w:pStyle w:val="TOC1"/>'
  );

/** Word's table of contents with its heading paragraph, misspelt, before the field. */
const FILE = docx(
  '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    p(t('Contens'), '<w:pStyle w:val="TOCHeading"/>') +
    row(
      'Terms',
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    ) +
    p('<w:r><w:fldChar w:fldCharType="end"/></w:r>') +
    p(t('Note after.')) +
    '</w:sdtContent></w:sdt>' +
    p(t('Terms'), '<w:pStyle w:val="Heading1"/>')
);

const fixHeading: Act = (e) => (select(e, 0, 5), e.exec({ type: 'insertText', text: 't' }));
const typeAfterField: Act = (e) => (select(e, 3, 0), e.exec({ type: 'insertText', text: 'A ' }));

describe("a table of contents' heading paragraph", () => {
  for (const [who, mode, login] of [
    ['the organiser', 'edit', ORGANISER],
    ['a party', 'suggesting', PARTY],
  ] as const)
    for (const [where, edit] of [
      ['fixes a typo in it', fixHeading],
      ['types after the field, in the same control', typeAfterField],
    ] as const)
      test(`${who} ${where}, as the editor alone does`, async () => {
        const room = localRoom(FILE);
        const typist = await join(room, 'Ana', { mode, login, wordFile: FILE });
        const other = await join(room, 'Olive', {
          mode: 'edit',
          login: ORGANISER,
          wordFile: FILE,
        });
        expect(await act(typist, edit)).toEqual({ ok: true, changed: true });
        converged(room, [typist, other]);
        const editorAlone = await alone(FILE, 'Ana', mode, edit);
        const saved = roomFile(room.room.document, FILE);
        expect(storyAsRead(saved)).toBe(storyAsRead(editorAlone));
      });

  test('the room itself refuses typing in a row, past any editor', () => {
    const { document: doc } = openOrderedWordFile(FILE, { file: 1 });
    const typed = (index: number) => () =>
      withUndoRecord(doc, () =>
        applyOrderedChange(
          doc,
          {
            v: 1,
            id: `olive-1:${index + 1}`,
            kind: 'insert',
            part: BODY,
            mode: 'edit',
            place: orderedPlaceAt(orderedViewParagraph(doc, BODY, index)!, 2),
            text: 'X',
          } as OrderedChange,
          { person: testPerson('Olive'), date: DATE, ids: orderedIdMinter(`olive-1:${index + 1}`) }
        )
      );
    expect(typed(1)).toThrow('a table of contents result is generated');
    expect(typed(0)).not.toThrow();
  });

  test('its rows stay read-only', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    select(olive.editor, 1, 2);
    expect(await act(olive, (e) => e.exec({ type: 'insertText', text: 'X' }))).toMatchObject({
      ok: false,
    });
    expect(surface(olive).state().lastRejection).toBe(
      'the table of contents is generated and read-only'
    );
    converged(room, [olive]);
  });
});
