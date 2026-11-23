// Typing at an edge of an inline control, with real editors and one ordered room, against the
// editor alone. The keyboard types on in the control the caret is in, before its first
// character and after its last too (the engine's `inside`). The room once put that text beside
// the control: after a prompt was typed over, every key from the second one left the control,
// and the page showed the room's version, not the editor alone's. Each case is checked on every
// copy, in a save and reopen, and after one and two folds.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { checkFold, mediaStore } from '../fold-check.ts';
import { docx, p, t } from '../word-fixtures.ts';
import { aloneShape, shapeOf } from './control-shape.ts';
import {
  BODY,
  caretAt,
  closeAll,
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
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const FACE = '<w:rStyle w:val="PlaceholderText"/>';

/** "A: " and a control holding `content`, with `tail` after it. */
const fileWith = (props: string, content: string, tail: string): Uint8Array =>
  docx(
    p(
      t('A: ') +
        `<w:sdt><w:sdtPr><w:id w:val="1"/>${props}</w:sdtPr><w:sdtContent>${content}` +
        `</w:sdtContent></w:sdt>${tail === '' ? '' : t(tail)}`
    ) + p(t('Signed.'))
  );

/** The controls typing goes into: rich text, plain text, and a rich text control's prompt. */
const CONTROLS = [
  ['rich text', '', t('abc')],
  ['plain text', '<w:text/>', t('abc')],
  ['prompt', '<w:showingPlcHdr/>', t('abc', FACE)],
] as const;

/** Who types: the organiser directly or suggesting, and a party. */
const PEOPLE = [
  ['the organiser, directly', 'Olive', 'edit'],
  ['the organiser, suggesting', 'Olive', 'suggesting'],
  ['a party', 'Ben', 'suggesting'],
] as const;

/** Two keys, as the keyboard sends them. */
const typeTwo = (person: Person) => {
  surface(person).type('d');
  surface(person).type('e');
};

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
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

describe('typing at an edge of an inline control stays in it, as the editor alone types', () => {
  for (const [control, props, content] of CONTROLS) {
    for (const [where, offset] of [
      ['first', 3],
      ['last', 6],
    ] as const) {
      for (const tail of ['', ' end.']) {
        for (const [who, name, mode] of PEOPLE) {
          const label = `${who}, at the ${where} character of a ${control} control${tail ? ', text after it' : ''}`;
          test(label, async () => {
            const file = fileWith(props, content, tail);
            const review = reviewTestModule();
            const room = localRoom(file);
            const olive = await join(room, 'Olive', {
              mode: name === 'Olive' ? mode : 'edit',
              login: ORGANISER,
              wordFile: file,
              review,
            });
            const ben = await join(room, 'Ben', { wordFile: file, review });
            const person = name === 'Olive' ? olive : ben;
            const expected = aloneShape(file, name, mode, [0, offset, offset], typeTwo);
            caretAt(person, 0, offset);
            typeTwo(person);
            await settle();
            converged(room, olive, ben);
            expect(shapeOf(editorXml(olive))).toBe(expected);
            expect(shapeOf(editorXml(ben))).toBe(expected);
            const saved = new Uint8Array(await person.editor.save());
            const reopened = openOrderedWordFile(saved, { file: 2 }).document;
            expect(shapeOf(serializeOoxmlPart(composeOrderedPart(reopened, BODY).part))).toBe(
              expected
            );
            const { base } = openOrderedWordFile(file, { file: 1 });
            await checkFold(room.room.document, base, mediaStore(base), label);
          });
        }
      }
    }
  }
});
