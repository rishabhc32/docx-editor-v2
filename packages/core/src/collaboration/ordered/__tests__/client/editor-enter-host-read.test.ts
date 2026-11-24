// Enter, then typing, while the host reads the page geometry on every change, with real editors
// and one ordered room. The room's version of an Enter names the new paragraph by its room
// address, so the editor's caret moves off the engine's own new paragraph (`#new:0`). A host
// read during that change lays the pages out before the session has said where the caret goes;
// the caret then stayed on `#new:0`, which the room does not know, and every later key was
// refused as `out-of-step`, in every story. The room's echo of the person's own Enter changes
// nothing here (one clock for the room and the people), so no later change moved the caret.

import { afterEach, expect, test } from 'bun:test';
import { storyParagraphs } from '../../../../store/package/story-blocks.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { textboxStoriesInPart } from '../../../../store/package/textbox-stories.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, R, t, W } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
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
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const AT = Date.UTC(2026, 9, 28, 9, 0, 0);

const NUMBERED = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>';
const LIST = docx(p(t('Item one.'), NUMBERED) + p(t('After.')), MATRIX_EXTRAS);
const CELL = docx(
  p(t('Before.')) +
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr/>' +
    `${p(t('Cell one.'))}</w:tc></w:tr></w:tbl>` +
    p(t('After.'))
);
const NOTE_REFERENCE =
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';
const NOTED = docx(p(t('The fee') + NOTE_REFERENCE + t(' is due.')), MATRIX_EXTRAS);
const textBox = MATRIX_ITEMS.find((item) => item.key === 'textbox');
if (!textBox) throw new Error('no text box item');
const BOXED = docx(bodyOf(textBox), MATRIX_EXTRAS);
const HEADED = docx(
  p(t('The fee is 100 USD.')) +
    `<w:sectPr xmlns:r="${R}"><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>`,
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}">${p(t('Header text.'))}</w:hdr>`,
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes:
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
  }
);

function caretIn(person: Person, paragraphId: string, offset: number): void {
  const position = { paragraphId, offset };
  surface(person).setSelection({ anchor: position, head: position });
}

/** The caret in the body paragraph that starts with `start`, at `offset` (default: its end). */
function caretAtText(person: Person, start: string, offset?: number): void {
  const texts = editorTexts(person);
  const index = texts.findIndex((text) => text.startsWith(start));
  if (index < 0) throw new Error(`no paragraph ${start}`);
  caretAt(person, index, offset ?? texts[index]!.length);
}

interface Shape {
  readonly name: string;
  readonly file: Uint8Array;
  /** Put the caret where the Enter goes. */
  readonly place: (person: Person) => void;
  readonly enters?: number;
}

const SHAPES: readonly Shape[] = [
  { name: 'a paragraph end', file: FILE, place: (person) => caretAtText(person, 'The fee') },
  { name: 'a paragraph middle', file: FILE, place: (person) => caretAtText(person, 'The fee', 7) },
  { name: 'a paragraph start', file: FILE, place: (person) => caretAtText(person, 'Due', 0) },
  {
    name: 'a paragraph end, twice',
    file: FILE,
    place: (person) => caretAtText(person, 'The fee'),
    enters: 2,
  },
  {
    name: 'a soft break',
    file: FILE,
    place: (person) => {
      caretAtText(person, 'The fee');
      surface(person).insertLineBreak();
    },
  },
  { name: 'a list item', file: LIST, place: (person) => caretAtText(person, 'Item one.') },
  { name: 'a table cell', file: CELL, place: (person) => caretAtText(person, 'Cell one.') },
  {
    name: 'a footnote',
    file: NOTED,
    place: (person) => {
      expect(surface(person).enterNote('footnote:1')).toBe(true);
    },
  },
  {
    name: 'a text box',
    file: BOXED,
    place: (person) => {
      const match = person.editor.findMatches('Box text')[0];
      if (!match?.scope) throw new Error('no text box match');
      expect(surface(person).setActiveScope(match.scope)).toBe(true);
      const [story] = textboxStoriesInPart(surface(person).session.part());
      const paragraph = story ? storyParagraphs(story.root)[0] : undefined;
      if (!paragraph) throw new Error('no text box paragraph');
      caretIn(person, paragraph.id, 'Box text'.length);
    },
  },
  {
    name: 'a header',
    file: HEADED,
    place: (person) => {
      expect(surface(person).enterHeaderFooter({ rId: 'rIdHeader1', kind: 'header' })).toBe(true);
      const ids = surface(person).session.paragraphIdsIn({
        kind: 'headerFooter',
        rId: 'rIdHeader1',
      });
      const first = ids[0];
      if (first === undefined) throw new Error('no header paragraph');
      caretIn(person, first, 'Header text.'.length);
    },
  },
];

for (const shape of SHAPES) {
  for (const who of ['party', 'organiser'] as const) {
    test(`Enter at ${shape.name}, then typing, under a host read (${who})`, async () => {
      const room = localRoom(shape.file, { now: () => AT });
      const options = { wordFile: shape.file, now: () => AT, review: reviewTestModule() };
      const ana = await join(room, 'Ana', options);
      const org = await join(room, 'Org', { ...options, mode: 'edit', login: ORGANISER });
      const person = who === 'party' ? ana : org;
      // As a page placing its review cards does
      person.editor.on('change', () => {
        person.editor.getPageGeometry();
      });
      shape.place(person);
      await settle();
      for (let enter = 0; enter < (shape.enters ?? 1); enter += 1) {
        surface(person).splitParagraph();
      }

      await settle();
      expect(surface(person).state().selection.head.paragraphId).not.toContain('#new:');
      surface(person).insertPlainText('Q');
      await settle();
      expect(person.refusals).toEqual([]);
      for (const one of [ana, org]) {
        expect(sameAsRoom(one, room)).toBe(true);
        expect(screenFollowsList(one)).toBe(true);
      }

      // The saved file reopens with the letter typed after the Enter
      const saved = new Uint8Array(await org.editor.save());
      const again = openOrderedWordFile(saved, { file: 2 }).document;
      const stories = [...again.stories.keys()].map((name) =>
        serializeOoxmlPart(composeOrderedPart(again, name).part)
      );
      expect(stories.some((xml) => xml.includes('>Q<'))).toBe(true);
    });
  }
}
