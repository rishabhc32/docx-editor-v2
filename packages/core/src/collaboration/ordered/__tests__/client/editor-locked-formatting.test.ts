// Formatting inside a content-locked control in an ordered room, with the editor's
// `contentLockPolicy.allowFormattingInLockedContent` (audit R1 row 19). The editor alone lets an
// untracked restyle past the content lock and refuses a tracked one; the room does the same
// when its own policy says so (`OrderedRoomOptions.contentLockPolicy`, sent in `welcome`): the
// organiser's direct format holds on every copy, through undo, redo and a save and reopen; a
// party's format, a suggestion, stays refused (`locked`), and so does every format when the
// room's policy is off.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  join,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  storiesOf,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const chip = (lock: string) =>
  `<w:sdt><w:sdtPr><w:tag w:val="t"/><w:lock w:val="${lock}"/></w:sdtPr>` +
  `<w:sdtContent>${t('TOKEN')}</w:sdtContent></w:sdt>`;
const fileOf = (lock: string) => docx(p(t('head ') + chip(lock) + t(' tail')));
const POLICY = { allowFormattingInLockedContent: true };
const BOLD = { type: 'toggleMark', mark: 'bold' } as const;

/** The chip's runs, as the list composes them. */
const chipXml = (person: Person): string =>
  listXml(person).match(/<w:sdtContent>.*<\/w:sdtContent>/)?.[0] ?? '';

/** A room on `file` with Olive, the organiser in editing, and Ana, a party in suggesting. */
async function people(
  file: Uint8Array,
  room = localRoom(file, { policy: suggestionsOnlyPolicy() })
) {
  const options = { wordFile: file, contentLockPolicy: POLICY };
  const olive = await join(room, 'Olive', { ...options, mode: 'edit', login: ORGANISER });
  const ana = await join(room, 'Ana', options);
  return { room, olive, ana };
}

/** Every copy is the room's document, and every editor shows its copy. */
function converged(room: ReturnType<typeof localRoom>, people: readonly Person[]): void {
  for (const person of people) {
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
}

describe('formatting a locked control in an ordered room', () => {
  for (const lock of ['contentLocked', 'sdtContentLocked']) {
    test(`the organiser bolds a ${lock} chip when the room lets formats pass`, async () => {
      const file = fileOf(lock);
      const room = localRoom(file, { policy: suggestionsOnlyPolicy(), contentLockPolicy: POLICY });
      const { olive, ana } = await people(file, room);
      caretAt(olive, 0, 5, 10);
      expect(olive.editor.exec(BOLD)).toEqual({ ok: true, changed: true });
      await settle();
      expect(olive.refusals).toEqual([]);
      converged(room, [olive, ana]);
      for (const person of [olive, ana]) {
        expect(chipXml(person)).toContain('<w:b/>');
        expect(chipXml(person)).toContain('TOKEN');
        expect(editorXml(person)).toContain(`w:val="${lock}"`);
      }

      expect(olive.handle.session.undo()).toBe(true);
      await settle();
      converged(room, [olive, ana]);
      expect(chipXml(ana)).not.toContain('<w:b/>');
      expect(olive.handle.session.redo()).toBe(true);
      await settle();
      converged(room, [olive, ana]);
      expect(chipXml(ana)).toContain('<w:b/>');

      // Saved and reopened, the file holds the bold chip with its lock
      const base = readOoxmlPackage(file);
      if (!base.ok) throw new Error('no base');
      const reopened = openOrderedWordFile(orderedPackageBytes(room.room.document, base.package), {
        file: 1,
      });
      expect(storiesOf(reopened.document)).toEqual(storiesOf(room.room.document));
    });
  }

  test("a party's format in a locked chip stays refused, as the editor alone refuses a tracked one", async () => {
    const file = fileOf('contentLocked');
    const room = localRoom(file, { policy: suggestionsOnlyPolicy(), contentLockPolicy: POLICY });
    const { olive, ana } = await people(file, room);
    caretAt(ana, 0, 5, 10);
    expect(ana.editor.exec(BOLD).ok).toBe(false);
    await settle();
    expect(ana.refusals.map((refusal) => refusal.reason)).toEqual(['locked']);
    expect(room.room.order).toBe(0);
    converged(room, [olive, ana]);
    expect(chipXml(ana)).not.toContain('<w:b/>');
  });

  test("without the room's policy the organiser's format stays refused (locked)", async () => {
    const { room, olive, ana } = await people(fileOf('contentLocked'));
    caretAt(olive, 0, 5, 10);
    expect(olive.editor.exec(BOLD).ok).toBe(false);
    await settle();
    expect(olive.refusals.map((refusal) => refusal.reason)).toEqual(['locked']);
    expect(room.room.order).toBe(0);
    converged(room, [olive, ana]);
  });
});
