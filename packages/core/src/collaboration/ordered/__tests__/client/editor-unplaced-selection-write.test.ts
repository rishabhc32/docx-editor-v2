// A write over a range that ends in text the page does not show, with two
// real editors and one room.
//
// A paragraph hidden whole (`w:vanish` on its text and its mark) is in no laid-out order. A
// format, an indent, a deletion or typing over a range that ended in it did nothing and
// reported nothing. Each is now refused by name, for the organiser and for a party, and
// nothing reaches the room or the other person. A range inside the shown text still writes.

import { afterEach, describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../../../contracts/editor.ts';
import { UNPLACED_SELECTION_REFUSAL } from '../../../../editor/surface-range-edit.ts';
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
  p(t('One')) + p(t('Hidden', '<w:vanish/>'), '<w:rPr><w:vanish/></w:rPr>') + p(t('Three'))
);

const WRITES: readonly EditorCommand[] = [
  { type: 'toggleMark', mark: 'bold' },
  { type: 'setIndent', left: 720 },
  { type: 'setAlignment', align: 'center' },
  { type: 'deleteText' },
  { type: 'insertText', text: 'X' },
  { type: 'clearFormatting' },
];

async function people() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { wordFile: FILE, login: ORGANISER, mode: 'edit' });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, olive, ana };
}

describe('a range that ends in hidden text is refused by name', () => {
  for (const who of ['Olive (organiser)', 'Ana (party)'] as const) {
    for (const command of WRITES) {
      test(`${who}: ${command.type}`, async () => {
        const { room, olive, ana } = await people();
        const person = who.startsWith('Olive') ? olive : ana;
        const before = editorXml(olive);
        caretAt(person, 0, 0, 2, 1);

        expect(person.editor.exec(command)).toEqual({
          ok: false,
          code: 'unsupported',
          reason: UNPLACED_SELECTION_REFUSAL,
        });
        await settle();

        expect(person.refusals).toEqual([]);
        expect(editorXml(olive)).toBe(before);
        expect(editorXml(ana)).toBe(before);
        expect(sameAsRoom(olive, room)).toBe(true);
        expect(sameAsRoom(ana, room)).toBe(true);
      });
    }

    test(`${who}: a range inside the shown text still writes`, async () => {
      const { room, olive, ana } = await people();
      const person = who.startsWith('Olive') ? olive : ana;
      caretAt(person, 0, 0, 2, 2);

      expect(person.editor.exec({ type: 'toggleMark', mark: 'bold' })).toEqual({
        ok: true,
        changed: true,
      });
      await settle();

      expect(person.refusals).toEqual([]);
      expect(editorXml(olive)).toBe(editorXml(ana));
      expect(sameAsRoom(olive, room)).toBe(true);
      expect(sameAsRoom(ana, room)).toBe(true);
    });
  }
});
