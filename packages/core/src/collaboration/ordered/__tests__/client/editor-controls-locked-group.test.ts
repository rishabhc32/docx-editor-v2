// A field in a Word-saved group, with real editors and one ordered room. Word saves every group
// with `contentLocked`, and lets a person fill the controls nested in it while the group's own
// words take no edit (desktop Word for Mac, 2026-11-26). The organiser and a party fill the field,
// live and after a save and reopen; an edit of the group's own words is refused by name
// (`locked`). The engine and the room once passed the group's lock down and refused the field.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const sdt = (props: string, inner: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

/** "A: " + a Word-saved group > "G " + a text field "ab", then " end.": own 3 to 5, field 5 to 7. */
const FILE = docx(
  p(
    t('A: ') +
      sdt(
        '<w:tag w:val="group"/><w:lock w:val="contentLocked"/><w:group/>',
        t('G ') + sdt('<w:tag w:val="field"/><w:text/>', t('ab'))
      ) +
      t(' end.')
  ) + p(t('Signed.'))
);

/** The field's words as the room composes them. */
const fieldOf = (doc: OrderedDocument): string =>
  [
    ...(
      /<w:tag w:val="field"\/>.*?<\/w:sdt>/s.exec(
        serializeOoxmlPart(composeOrderedPart(doc, BODY).part)
      )?.[0] ?? ''
    ).matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g),
  ]
    .map((match) => match[1])
    .join('');

async function seats() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const bob = await join(room, 'Bob', { wordFile: FILE });
  return { room, olive, bob };
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

describe('a field in a Word-saved group in an ordered room', () => {
  for (const who of ['Olive', 'Bob'] as const) {
    test(`${who} fills the field, live and after a save and reopen`, async () => {
      const { room, olive, bob } = await seats();
      const person = who === 'Olive' ? olive : bob;
      caretAt(person, 0, 6);
      surface(person).insertPlainText('Q');
      await settle();
      expect(person.refusals).toEqual([]);
      converged(room, olive, bob);
      expect(editorTexts(olive)[0]).toBe('A: G aQb end.');
      expect(fieldOf(room.room.document)).toBe('aQb');

      const saved = new Uint8Array(await person.editor.save());
      expect(fieldOf(openOrderedWordFile(saved, { file: 2 }).document)).toBe('aQb');
    });

    test(`${who}'s edit of the group's own words is refused by name`, async () => {
      const { room, olive, bob } = await seats();
      const person = who === 'Olive' ? olive : bob;
      caretAt(person, 0, 4);
      surface(person).insertPlainText('Q');
      await settle();

      // A deletion from the own words into the field is refused whole
      caretAt(person, 0, 4, 6);
      surface(person).deleteBackward();
      await settle();
      expect(person.refusals.map((refusal) => refusal.reason)).toEqual(['locked', 'locked']);
      expect(room.room.order).toBe(0);
      converged(room, olive, bob);
    });
  }

  test("a party's words in the field meet the organiser's deletion there, and an undo", async () => {
    const { room, olive, bob } = await seats();
    caretAt(bob, 0, 6);
    surface(bob).insertPlainText('Q');
    caretAt(olive, 0, 5, 6);
    surface(olive).deleteBackward();
    await settle();
    expect([...olive.refusals, ...bob.refusals]).toEqual([]);
    converged(room, olive, bob);
    expect(editorTexts(bob)[0]).toBe(editorTexts(olive)[0]);

    surface(bob).undo();
    await settle();
    converged(room, olive, bob);
    expect(fieldOf(room.room.document)).not.toContain('Q');
  });
});
