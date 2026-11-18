// A paste that holds another person's tracked changes, in an ordered room (group 9.10, design
// D18). The editor's copy from a document that does not track changes keeps Cara's insertion,
// deletion and format change, with her name and dates, as Word copies them. In a room the
// server stamps every change with its sender, so a paste never brings Cara's name in: the room
// pastes the fragment's accepted final view, as Word pastes from a document that tracks changes.
// The organiser's paste is direct, and undo and redo turn it back and again; a party's is one
// insertion in the party's name, which a reject takes out.

import { afterEach, describe, expect, test } from 'bun:test';
import { fragmentFromHtml } from '../../../../editor/clipboard-fragment-codec.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, p, packageOf, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedBetween } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

type Room = ReturnType<typeof localRoom>;

const CARA = (minute: string) => `w:author="Cara" w:date="2026-01-02T10:${minute}:00Z"`;

/** Word's sample of paste evidence 1: Cara's insertion, deletion and format change. */
const SOURCE = docx(
  p(
    t('Start ') +
      `<w:ins w:id="1" ${CARA('00')}>${t('inserted ')}</w:ins>` +
      `<w:del w:id="2" ${CARA('01')}><w:r><w:delText xml:space="preserve">deleted </w:delText></w:r></w:del>` +
      `<w:r><w:rPr><w:b/><w:rPrChange w:id="3" ${CARA('02')}><w:rPr/></w:rPrChange></w:rPr>` +
      '<w:t>formatted</w:t></w:r>' +
      t(' end.')
  ) + p(t('Next.'))
);

/** The whole first line, its mark too, as the editor copies it from a file that does not track. */
const copiedWithChanges = () => copiedBetween(0, 1, 0, SOURCE);

const BEFORE = ['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.'];

const kinds = (room: Room): string[] => room.room.joinSnapshot(0).map((change) => change.kind);

const checked = (room: Room, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

/** The room's saved main document part. */
function savedBody(room: Room): string {
  const pkg = readOoxmlPackage(orderedPackageBytes(room.room.document, packageOf(FILE)));
  if (!pkg.ok) throw new Error('the saved file does not open');
  return serializeOoxmlPart(pkg.package.parts.get('/word/document.xml')!);
}

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

describe("a paste that holds another person's tracked changes, in an ordered room", () => {
  test("the copy keeps Cara's changes, as Word copies from a file that does not track", () => {
    const fragment = fragmentFromHtml(copiedWithChanges().html ?? '');
    if (!fragment) throw new Error('no fragment');
    const pkg = readOoxmlPackage(fragment.bytes);
    if (!pkg.ok) throw new Error('the fragment does not open');
    const xml = serializeOoxmlPart(pkg.package.parts.get(pkg.package.mainDocumentPart)!);
    expect(xml).toContain(`<w:ins ${CARA('00')}`);
    expect(xml).toContain(`<w:del ${CARA('01')}`);
    expect(xml).toContain(`<w:rPrChange ${CARA('02')}`);
  });

  test("the organiser's: the accepted final view, direct; no one else's name; undo and redo", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const flavours = copiedWithChanges();
    caretAt(olive, 1, 0);
    expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
    await settle(16);
    expect(kinds(room)).toEqual(['insert-fragment']);
    checked(room, [olive, ana]);
    for (const person of [olive, ana]) {
      expect(editorTexts(person)).toEqual([
        BEFORE[0],
        'Start inserted formatted end.',
        BEFORE[1],
        BEFORE[2],
      ]);
      const xml = editorXml(person);
      expect(xml).not.toContain('Cara');
      expect(xml).not.toContain('<w:ins ');
      expect(xml).not.toContain('<w:del ');
      expect(xml).toMatch(/<w:b\/><\/w:rPr><w:t>formatted<\/w:t>/);
    }
    expect(savedBody(room)).not.toContain('Cara');
    checkSaveAndLoad(room.room.document, 'paste with changes', packageOf(FILE));

    expect(olive.handle.session.undo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(ana)).toEqual(BEFORE);
    expect(olive.handle.session.redo()).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(ana)[1]).toBe('Start inserted formatted end.');
  });

  test("a party's: one insertion in the party's name, with no one else's; a reject takes it out", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    const flavours = copiedWithChanges();
    caretAt(ana, 1, 0);
    expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
    await settle(16);
    expect(kinds(room)).toEqual(['insert-fragment']);
    checked(room, [olive, ana]);
    const xml = editorXml(olive);
    expect(xml).not.toContain('Cara');
    expect(xml).not.toContain('<w:del ');
    expect(xml).not.toContain('rPrChange');
    expect(xml).toMatch(/<w:ins [^>]*w:author="Ana"/);
    expect(xml).toMatch(/<w:b\/><\/w:rPr><w:t>formatted<\/w:t>/);
    expect(editorTexts(olive)[1]).toBe('Start inserted formatted end.');
    expect(savedBody(room)).not.toContain('Cara');
    checkSaveAndLoad(room.room.document, 'party paste with changes', packageOf(FILE));

    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(cards(olive)).toHaveLength(1);
    expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(olive)).toEqual(BEFORE);
  });
});
