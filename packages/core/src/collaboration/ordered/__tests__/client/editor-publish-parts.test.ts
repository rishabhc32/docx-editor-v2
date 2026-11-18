// A publish keeps each part that says what the editor already shows.
//
// Once a change adds a list, every publish writes `numbering.xml` again. A new part object with
// the same content told the editor's store that a part outside the main one changed: each
// keystroke laid the whole document out again and built a new numbering index. Unchanged parts
// keep their objects; a part that changes is still published, and a save holds the room's parts.

import { afterEach, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const NUMBERING = '/word/numbering.xml';
const FILE = docx(
  p(t('Plain one.')) +
    p(t('Plain gap.')) +
    p(t('First item.'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
    p(t('Plain two.')) +
    p(t('Plain gap.')) +
    p(t('Plain three.')),
  MATRIX_EXTRAS
);

const shown = (person: Person): OoxmlPackage => surface(person).session.currentPackage();

/** The names of the parts whose objects differ between two packages. */
function changedParts(before: OoxmlPackage, after: OoxmlPackage): string[] {
  return [...after.parts].filter(([name, part]) => before.parts.get(name) !== part).map(([n]) => n);
}

const abstractNums = (pkg: OoxmlPackage): number =>
  serializeOoxmlPart(pkg.parts.get(NUMBERING)!).split('<w:abstractNum ').length - 1;

async function openRoom() {
  const review = reviewTestModule();
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  // A new list: its definition is a resource of the change, written at every publish
  caretAt(olive, 0, 0);
  expect(olive.editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
  await settle();
  // A new footer: a story, its relationship and its content type, also written at every publish
  const made = surface(olive).applyHeaderFooterLifecycle?.({
    op: 'createHeaderFooter',
    sectionIndex: 0,
    kind: 'footer',
    variant: 'default',
  });
  expect(made?.ok).toBe(true);
  await settle();
  expect(room.room.joinSnapshot(0).map((change) => change.kind)).toContain('new-list');
  expect([...shown(ana).parts.keys()].some((name) => /footer/.test(name))).toBe(true);
  return { room, olive, ana };
}

test('typing keeps every part but the main one, and the package lists, in both copies', async () => {
  const { room, olive, ana } = await openRoom();
  for (const typist of [ana, olive]) {
    const before = [shown(olive), shown(ana)];
    caretAt(typist, 3, 0);
    surface(typist).insertPlainText(typist.name[0]!);
    await settle();
    for (const [k, person] of [olive, ana].entries()) {
      expect(changedParts(before[k]!, shown(person))).toEqual([BODY]);
      // A header or footer occurrence is keyed on the main part's relationship list
      expect(shown(person).relationships.get(BODY)).toBe(before[k]!.relationships.get(BODY)!);
      expect(shown(person).contentTypes).toBe(before[k]!.contentTypes);
    }
  }
  for (const person of [olive, ana]) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
});

test('a change to the list definitions still publishes the numbering part', async () => {
  const { room, olive, ana } = await openRoom();
  const before = [shown(olive), shown(ana)];
  caretAt(olive, 5, 0);
  expect(olive.editor.exec({ type: 'toggleList', kind: 'ordered' }).ok).toBe(true);
  await settle();
  for (const [k, person] of [olive, ana].entries()) {
    expect(changedParts(before[k]!, shown(person))).toContain(NUMBERING);
    expect(abstractNums(shown(person))).toBe(abstractNums(before[k]!) + 1);
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
  }
});

test("a save from either copy holds the room's parts and reopens", async () => {
  const { room, olive, ana } = await openRoom();
  for (const typist of [ana, olive]) {
    caretAt(typist, 3, 0);
    surface(typist).insertPlainText(typist.name[0]!);
    await settle();
  }
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error(base.reason);
  const composed = readOoxmlPackage(orderedPackageBytes(room.room.document, base.package));
  if (!composed.ok) throw new Error(composed.reason);
  for (const person of [olive, ana]) {
    const bytes = new Uint8Array(await person.editor.save());
    const saved = readOoxmlPackage(bytes);
    if (!saved.ok) throw new Error(saved.reason);
    for (const name of [BODY, NUMBERING])
      expect(serializeOoxmlPart(saved.package.parts.get(name)!)).toBe(
        serializeOoxmlPart(composed.package.parts.get(name)!)
      );
    expect(openOrderedWordFile(bytes, { file: 2 }).document.stories.has(BODY)).toBe(true);
  }
});
