// A bookmark write the room refuses keeps the room's reason, with two real
// editors and one room.
//
// The bookmark lane handed `commit` only whether its write committed, so `exec` answered
// "insertBookmark was refused" and lost the room's code. It now answers with the room's code,
// for the organiser and for a party, and nothing reaches the room or the other person.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
} from './editor-harness.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('Free words.')) +
    '<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>' +
    p(t('Locked words.')) +
    '</w:sdtContent></w:sdt>'
);

async function people() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { wordFile: FILE, login: ORGANISER, mode: 'edit' });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, olive, ana };
}

describe('a bookmark write the room refuses keeps the room reason', () => {
  for (const who of ['Olive (organiser)', 'Ana (party)'] as const) {
    test(`${who}: insertBookmark inside a locked content control`, async () => {
      const { room, olive, ana } = await people();
      const person = who.startsWith('Olive') ? olive : ana;
      const before = editorXml(olive);
      caretAt(person, 1, 0, 6);
      const refusals = person.refusals.length;

      expect(person.editor.exec({ type: 'insertBookmark', name: 'Added' })).toEqual({
        ok: false,
        code: 'unsupported',
        reason: 'ordered-refused',
      });
      await settle();

      expect(person.refusals.length).toBe(refusals + 1);
      expect(person.refusals.at(-1)).toMatchObject({ code: 'ordered-refused' });
      expect(editorXml(olive)).toBe(before);
      expect(editorXml(ana)).toBe(before);
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(sameAsRoom(ana, room)).toBe(true);
    });

    test(`${who}: a bookmark on free text reaches the other person`, async () => {
      const { room, olive, ana } = await people();
      const person = who.startsWith('Olive') ? olive : ana;
      caretAt(person, 0, 0, 4);

      expect(person.editor.exec({ type: 'insertBookmark', name: 'Added' })).toEqual({
        ok: true,
        changed: true,
      });
      await settle();

      expect(editorXml(olive)).toContain('w:name="Added"');
      expect(editorXml(ana)).toContain('w:name="Added"');
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(sameAsRoom(ana, room)).toBe(true);
    });
  }
});
