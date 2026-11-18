// A paste too long for one change goes as several, cut at paragraph edges. A numbered list and a
// style the document does not hold travel once, in the first change; the later changes name
// what it made. Before, each change carried them again, so the list numbers started at 1 again
// in each change and the style was added once per change. Undo and redo keep one list.

import { afterEach, describe, expect, test } from 'bun:test';
import { docx, p, t, W } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  listXml,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedBetween } from './paste-harness.ts';

afterEach(closeAll);

const ITEMS = 120;

/** Twenty runs of alternate bold and italic words: a paragraph of about a hundred nodes. */
const runs = (k: number): string =>
  Array.from({ length: 20 }, (_, j) => t(`w${k}.${j} `, j % 2 === 0 ? '<w:b/>' : '<w:i/>')).join(
    ''
  );

const STYLES =
  `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal">` +
  '<w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="ListItem">' +
  '<w:name w:val="List Item"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/></w:pPr></w:style>' +
  '</w:styles>';

const NUMBERING =
  `<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="3">` +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>' +
  '<w:lvlJc w:val="left"/></w:lvl></w:abstractNum>' +
  '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>';

/** A numbered list in a style, longer than one change holds. */
const SOURCE = docx(
  Array.from({ length: ITEMS }, (_, k) =>
    p(
      runs(k),
      '<w:pStyle w:val="ListItem"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>'
    )
  ).join(''),
  {
    parts: { 'word/styles.xml': STYLES, 'word/numbering.xml': NUMBERING },
    contentTypes:
      '<Override PartName="/word/numbering.xml" ' +
      'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    relationships:
      '<Relationship Id="rIdNumbering" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" ' +
      'Target="numbering.xml"/>',
  }
);

/** The list numbers the person's list names. */
const listNumbers = (person: Person): string[] => [
  ...new Set([...listXml(person).matchAll(/<w:numId w:val="(\d+)"\/>/g)].map((m) => m[1] ?? '')),
];

/** The styles the changes added with the pasted style's id. */
const addedListItemStyles = (person: Person): number =>
  [...person.handle.session.orderedDocument().resources.added.values()].filter(
    (resource) => resource.kind === 'style' && resource.styleId === 'ListItem'
  ).length;

/** Both copies are the room, show one list and the pasted style once. */
function oneList(room: ReturnType<typeof localRoom>, people: readonly Person[]): void {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
    expect(editorTexts(person)).toHaveLength(ITEMS + 3);
    expect(listNumbers(person)).toHaveLength(1);
    expect(addedListItemStyles(person)).toBe(1);
  }
}

describe('a paste longer than one change', () => {
  for (const who of ['organiser', 'party'] as const) {
    test(`the ${who}'s: one list and one style, after undo and redo too`, async () => {
      const room = localRoom();
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER });
      const ana = await join(room, 'Ana');
      const paster = who === 'organiser' ? olive : ana;
      const { text, html } = copiedBetween(0, ITEMS - 1, undefined, SOURCE);
      caretAt(paster, 1, 0);
      expect(surface(paster).pasteRich(text, html)).toBe(true);
      await settle();
      // Cut into two changes
      expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual([
        'insert-fragment',
        'insert-fragment',
      ]);
      oneList(room, [olive, ana]);

      expect(paster.handle.session.undo()).toBe(true);
      await settle();
      expect(editorTexts(olive)).toEqual([
        'The fee is 100 USD.',
        'Due in 30 days.',
        'Signed here.',
      ]);
      expect(paster.handle.session.redo()).toBe(true);
      await settle();
      expect(listNumbers(olive)).toHaveLength(1);
      expect(listNumbers(ana)).toEqual(listNumbers(olive));
      expect(sameAsRoom(olive, room) && sameAsRoom(ana, room)).toBe(true);
    });
  }
});
