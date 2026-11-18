// Protect Document from the editor in an ordered room (task 9.10): the organiser's toggle is a
// `setting` change, a direct edit replayed on the settings part, that every copy and the saved
// file show; the organiser's undo and redo turn it back and again; a party's toggle is the
// `package` rule (the host's earlier check hashes `word/settings.xml`), because Word records no
// settings change as a suggestion.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const TOGGLE = { type: 'toggleDocumentProtection' } as const;

const enforced = (person: Person): boolean | undefined =>
  person.editor.snapshot().documentProtection?.enforced;

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

/** The settings part of the room's document, saved. */
function savedSettings(room: ReturnType<typeof localRoom>): string {
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error('no base');
  const bytes = orderedPackageBytes(room.room.document, base.package);
  const reopened = readOoxmlPackage(bytes);
  if (!reopened.ok) throw new Error('the saved file does not open');
  // A reopen loads the same stories
  openOrderedWordFile(bytes, { file: 2 });
  const settings = [...reopened.package.parts.values()].find((part) =>
    part.name.endsWith('/settings.xml')
  );
  return settings ? JSON.stringify(settings.root) : '';
}

describe('Protect Document in an ordered room', () => {
  test("the organiser's toggle reaches every copy and the saved file; undo and redo turn it", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana');
    expect(olive.editor.exec(TOGGLE)).toEqual({ ok: true, changed: true });
    await settle();
    expect(kinds(room)).toEqual(['setting']);
    for (const person of [olive, ana]) {
      expect(person.refusals).toEqual([]);
      expect(enforced(person)).toBe(true);
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
    }
    const on = savedSettings(room);
    expect(on).toContain('"documentProtection"');
    expect(on).toContain('"value":"forms"');
    expect(on).toContain('"value":"1"');

    // The organiser's undo is a revert of the setting's direct suggestion
    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['setting', 'revert']);
    for (const person of [olive, ana]) {
      expect(enforced(person)).not.toBe(true);
      expect(sameAsRoom(person, room)).toBe(true);
    }
    expect(savedSettings(room)).not.toContain('"documentProtection"');

    expect(olive.handle.session.redo()).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['setting', 'revert', 'revert']);
    for (const person of [olive, ana]) {
      expect(enforced(person)).toBe(true);
      expect(sameAsRoom(person, room)).toBe(true);
    }
  });

  test("lifting protection again is a second setting, replayed in the room's order", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    olive.editor.exec(TOGGLE);
    await settle();
    olive.editor.exec(TOGGLE);
    await settle();
    expect(kinds(room)).toEqual(['setting', 'setting']);
    expect(olive.refusals).toEqual([]);
    expect(enforced(olive)).toBe(false);
    expect(sameAsRoom(olive, room)).toBe(true);
    const off = savedSettings(room);
    expect(off).toContain('"documentProtection"');
    expect(off).toContain('"value":"0"');
  });

  test("a party's toggle is refused by the package rule, and nothing stays", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
    const ana = await join(room, 'Ana');
    ana.editor.exec(TOGGLE);
    await settle();
    expect(ana.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'package' }),
    ]);
    expect(kinds(room)).toEqual([]);
    for (const person of [olive, ana]) {
      expect(enforced(person)).not.toBe(true);
      expect(sameAsRoom(person, room)).toBe(true);
    }
  });
});
