// A paste inside a link, an inline content control or another person's pending insertion, in an
// ordered room (group 9.10, design D18), as Word lands it (paste evidence 2): in a link the
// pasted runs go into the link and keep their own formatting; in a rich text control they go
// into the control, each pasted paragraph mark a space; in another person's insertion the paste
// sits beside it, never nested. The organiser's paste is direct; a party's is one insertion in
// the party's name. Every copy equals the room, and the room equals the editor alone.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { grade, normalizedBody } from '../matrix/compare.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { docx, p, packageOf, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedBetween, type Copied } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

type Room = ReturnType<typeof localRoom>;

/** A link, a rich text control and plain text, one paragraph each. */
const FILE = docx(
  p(t('Go ') + '<w:hyperlink w:anchor="target"><w:r><w:t>LinkText</w:t></w:r></w:hyperlink>') +
    p(
      '<w:sdt><w:sdtPr><w:alias w:val="Ctl"/><w:id w:val="424242"/></w:sdtPr><w:sdtContent>' +
        '<w:r><w:t>CtlText</w:t></w:r></w:sdtContent></w:sdt>'
    ) +
    p(t('Base text.'))
);

const BEFORE = ['Go LinkText', 'CtlText', 'Base text.'];

/** Word's source: bold "XX", then "z"; a copy that reaches a line's end takes its mark. */
const SOURCE = docx(p(t('XX', '<w:b/>') + t('y')) + p(t('z') + t('y')));

/** "XX" without its mark, and "XXy¶z" without the last mark. */
const xx = (): Copied => copiedBetween(0, 0, 2, SOURCE);
const twoLines = (): Copied => copiedBetween(0, 1, 1, SOURCE);

const kinds = (room: Room): string[] => room.room.joinSnapshot(0).map((change) => change.kind);

const checked = (room: Room, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

const comparable = (part: OoxmlPart): string => normalizedBody(part, () => null);

/** What the editor alone writes for the same paste at the same place. */
function alone(flavours: Copied, party: boolean, paragraph: number, offset: number): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: party ? 'Ana' : 'Olive',
    mode: party ? 'suggesting' : 'edit',
    modules: [review],
  });
  try {
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const id = s.session.paragraphIds()[paragraph] ?? '';
    s.setSelection({ anchor: { paragraphId: id, offset }, head: { paragraphId: id, offset } });
    expect(s.pasteRich(flavours.text, flavours.html)).toBe(true);
    return comparable(s.session.part());
  } finally {
    editor.destroy();
    container.remove();
  }
}

const agrees = (room: Room, engine: string, party: boolean): string => {
  const g = grade(comparable(composeOrderedPart(room.room.document, BODY).part), engine);
  return g === 'exact' || g === 'same' || (party && g === 'ids') ? 'agrees' : g;
};

async function people(room: Room) {
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: FILE,
  });
  const ana = await join(room, 'Ana', { review, wordFile: FILE });
  return { olive, ana };
}

interface Case {
  readonly name: string;
  readonly copy: () => Copied;
  readonly paragraph: number;
  readonly offset: number;
  readonly texts: readonly string[];
  /** What the paste shows in a copy's body. */
  readonly shows: RegExp;
}

const CASES: readonly Case[] = [
  {
    name: 'a line inside a link: into the link, with its own formatting',
    copy: xx,
    paragraph: 0,
    offset: 7,
    texts: ['Go Link XX Text', BEFORE[1]!, BEFORE[2]!],
    shows: /<w:hyperlink [^>]*>(?:(?!<\/w:hyperlink>).)*<w:b\/>(?:(?!<\/w:hyperlink>).)*XX/,
  },
  {
    name: 'two lines inside a control: into it, the mark a space',
    copy: twoLines,
    paragraph: 1,
    offset: 3,
    texts: [BEFORE[0]!, 'Ctl XXy z Text', BEFORE[2]!],
    shows: /<w:sdtContent>(?:(?!<\/w:sdtContent>).)*XX(?:(?!<\/w:sdtContent>).)*z/,
  },
];

for (const item of CASES)
  describe(`${item.name}, in an ordered room`, () => {
    test("the organiser's: direct, as the editor alone pastes it; undo and redo", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const { olive, ana } = await people(room);
      const flavours = item.copy();
      caretAt(olive, item.paragraph, item.offset);
      expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual([...item.texts]);
      expect(editorXml(ana)).toMatch(item.shows);
      expect(editorXml(ana)).not.toContain('<w:ins ');
      expect(agrees(room, alone(flavours, false, item.paragraph, item.offset), false)).toBe(
        'agrees'
      );
      checkSaveAndLoad(room.room.document, item.name, packageOf(FILE));
      expect(olive.handle.session.undo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual(BEFORE);
      expect(olive.handle.session.redo()).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(ana)).toEqual([...item.texts]);
    });

    test("a party's: one insertion in the party's name, as the editor alone; a reject", async () => {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const { olive, ana } = await people(room);
      const flavours = item.copy();
      caretAt(ana, item.paragraph, item.offset);
      expect(surface(ana).pasteRich(flavours.text, flavours.html)).toBe(true);
      await settle(16);
      expect(kinds(room)).toEqual(['insert-fragment']);
      checked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual([...item.texts]);
      expect(editorXml(olive)).toMatch(item.shows);
      expect(editorXml(olive)).toMatch(/<w:ins [^>]*w:author="Ana"/);
      expect(agrees(room, alone(flavours, true, item.paragraph, item.offset), true)).toBe('agrees');
      checkSaveAndLoad(room.room.document, `party ${item.name}`, packageOf(FILE));
      const [card] = olive.editor
        .getReviewItems()
        .filter((placement) => placement.kind === 'revision')
        .map((placement) => placement.key);
      if (card === undefined) throw new Error('no card');
      expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
      await settle(16);
      checked(room, [olive, ana]);
      expect(editorTexts(olive)).toEqual(BEFORE);
    });
  });

describe("a paste inside another person's pending insertion, in an ordered room", () => {
  test("the organiser's: beside it, never nested, as Word pastes", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const { olive, ana } = await people(room);
    caretAt(ana, 2, 5);
    surface(ana).insertPlainText('really ');
    await settle(16);
    const flavours = xx();
    caretAt(olive, 2, 8);
    expect(surface(olive).pasteRich(flavours.text, flavours.html)).toBe(true);
    await settle(16);
    expect(kinds(room)).toEqual(['insert', 'insert-fragment']);
    checked(room, [olive, ana]);
    // Word's Smart cut and paste puts a space on each side between two words
    expect(editorTexts(olive)[2]).toBe('Base rea XX lly text.');
    // Ana's insertion is split: the bold XX and its spaces are no part of it
    expect(editorXml(olive)).toMatch(
      /<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t>rea<\/w:t><\/w:r><\/w:ins><w:r><w:rPr><w:b\/><\/w:rPr><w:t xml:space="preserve"> <\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>XX/
    );
    checkSaveAndLoad(room.room.document, 'beside an insertion', packageOf(FILE));
  });
});

describe('two lines inside a link, in an ordered room', () => {
  test('are refused by name in the editor, and the room sees no change', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const { olive, ana } = await people(room);
    for (const person of [olive, ana]) {
      const flavours = twoLines();
      caretAt(person, 0, 7);
      expect(surface(person).pasteRich(flavours.text, flavours.html)).toBe(false);
      await settle(16);
      expect(surface(person).state().lastRejection).toContain('fragment-paragraphs-in-link');
    }
    expect(kinds(room)).toEqual([]);
    checked(room, [olive, ana]);
    expect(editorTexts(ana)).toEqual(BEFORE);
  });
});
