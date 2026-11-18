// Select All and a format in a document with a table of contents, with two
// real editors and one room.
//
// The editor keeps the table's generated text read-only and formats the rest of the selection
// (issue #1223). Word formats the table's text too, so the part left out is reported through
// `lastRejection` with a reason that names the table, never dropped in silence. The body
// format reaches the room and the other person; the table's text stays as it was.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
} from './editor-harness.ts';

afterEach(closeAll);

const FILE = docx(
  '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('Terms\t1')}</w:p>` +
    p(t('Payment\t2')) +
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    p(t('Body text one.')) +
    p(t('Body text two.'))
);

const SKIPPED =
  'the table of contents is generated and read-only; the rest of the selection was formatted';

const paragraphs = (xml: string) => xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) ?? [];

async function people() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { wordFile: FILE, login: ORGANISER, mode: 'edit' });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, olive, ana };
}

describe('Select All and bold in a document with a table of contents', () => {
  for (const who of ['Olive (organiser)', 'Ana (party)'] as const) {
    test(`${who}: formats the body, reports the table, and keeps its text`, async () => {
      const { room, olive, ana } = await people();
      const person = who.startsWith('Olive') ? olive : ana;
      const before = paragraphs(editorXml(olive));
      surface(person).selectAll();

      expect(person.editor.exec({ type: 'toggleMark', mark: 'bold' })).toEqual({
        ok: true,
        changed: true,
      });
      expect(person.editor.snapshot().lastRejection).toBe(SKIPPED);
      await settle();

      expect(person.refusals).toEqual([]);
      for (const copy of [olive, ana]) {
        const after = paragraphs(editorXml(copy));
        // The table's three paragraphs are as they were; both body paragraphs are bold.
        expect(after.slice(0, 3)).toEqual(before.slice(0, 3));
        expect(after[3]).toContain('<w:b/>');
        expect(after[4]).toContain('<w:b/>');
      }
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(sameAsRoom(ana, room)).toBe(true);
    });
  }
});
