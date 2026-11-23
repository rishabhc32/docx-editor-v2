// A block control's prompt that holds a combo box showing its own prompt, with real editors and
// one ordered room, against the editor alone: "Courts at [Choose a city.] have jurisdiction." as
// the prompt's one paragraph, or as the second of two. The editor alone replaces the whole block
// prompt, its other paragraph and the combo box too, wherever the caret is in it, and Enter does
// the same except strictly inside the combo box (`placeholderControlForInsertion`). The room once
// typed into the combo box, refused Enter right after it, and kept the combo box's paragraph
// when content replaced a prompt of two paragraphs. Each edit is checked on every copy, in a save
// and reopen, and after one and two folds.

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
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const CITY =
  '<w:id w:val="9204"/><w:showingPlcHdr/><w:comboBox><w:listItem w:value=""/>' +
  '<w:listItem w:displayText="Mumbai" w:value="Mumbai"/></w:comboBox>';
const COURTS = p(
  t('Courts at ', FACE) + sdt(CITY, t('Choose a city.', FACE)) + t(' have jurisdiction.', FACE)
);
const venue = (prompt: string) =>
  docx(p(t('Venue:')) + sdt('<w:id w:val="9203"/><w:showingPlcHdr/>', prompt) + p(t('Signed.')));

/** The prompt of one paragraph, and of two, the combo box in the second. */
const ONE = venue(COURTS);
const TWO = venue(p(t('Terms apply.', FACE)) + COURTS);

/** Each file, its prompt's paragraphs, and places in them: from the paragraph's start. */
const FILES = [
  {
    name: 'one paragraph',
    file: ONE,
    unchanged: 'Venue:¶[#9203?Courts at [#9204?Choose a city.] have jurisdiction.¶]Signed.¶',
    places: [
      ['its start', 1, 0],
      ['its middle', 1, 5],
      ['the start of the combo box', 1, 10],
      ['the middle of the combo box', 1, 15],
      ['the end of the combo box', 1, 24],
      ['its end', 1, 43],
    ],
  },
  {
    name: 'two paragraphs',
    file: TWO,
    unchanged:
      'Venue:¶[#9203?Terms apply.¶Courts at [#9204?Choose a city.] have jurisdiction.¶]Signed.¶',
    places: [
      ['the first paragraph, its start', 1, 0],
      ['the first paragraph, its middle', 1, 5],
      ['the first paragraph, its end', 1, 12],
      ['the second paragraph, its start', 2, 0],
      ['the start of the combo box', 2, 10],
      ['the middle of the combo box', 2, 15],
      ['the end of the combo box', 2, 24],
      ['the second paragraph, its end', 2, 43],
    ],
  },
] as const;

/** The edits, as the keyboard and the clipboard make them. */
const EDITS = [['Enter', (person: Person) => surface(person).splitParagraph()]] as const;

/** Who edits: the organiser directly or suggesting, and a party. */
const PEOPLE = [
  ['the organiser, directly', 'Olive', 'edit'],
  ['the organiser, suggesting', 'Olive', 'suggesting'],
  ['a party', 'Ben', 'suggesting'],
] as const;

/** Both copies are the room's, show its list, and were refused only `refused`. */
function converged(
  room: ReturnType<typeof localRoom>,
  people: readonly Person[],
  refused: readonly string[] = []
): void {
  const reasons = people.flatMap((person) => person.refusals.map((report) => report.reason));
  expect(reasons).toEqual([...refused]);
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

/** Every copy, and a save and reopen, reads `expected`. */
async function everywhere(people: readonly Person[], expected: string): Promise<void> {
  for (const person of people) {
    expect({ person: person.name, shape: shapeOf(editorXml(person)) }).toEqual({
      person: person.name,
      shape: expected,
    });
  }

  const saved = new Uint8Array(await people[0]!.editor.save());
  const reopened = openOrderedWordFile(saved, { file: 2 }).document;
  expect(shapeOf(serializeOoxmlPart(composeOrderedPart(reopened, BODY).part))).toBe(expected);
}

/** The organiser, Olive, in `mode`, and Ben, a party, in one room on `file`. */
async function seats(file: Uint8Array, mode: 'edit' | 'suggesting') {
  const review = reviewTestModule();
  const room = localRoom(file);
  const olive = await join(room, 'Olive', { mode, login: ORGANISER, wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  return { room, olive, ben };
}

async function folds(
  room: ReturnType<typeof localRoom>,
  file: Uint8Array,
  label: string
): Promise<void> {
  const { base } = openOrderedWordFile(file, { file: 1 });
  await checkFold(room.room.document, base, mediaStore(base), label);
}

describe('each edit at each place of a block prompt that holds a combo box, as the editor alone', () => {
  for (const { name: file, file: bytes, unchanged, places } of FILES) {
    for (const [who, name, mode] of PEOPLE) {
      for (const [place, line, offset] of places) {
        for (const [edit, act] of EDITS) {
          const label = `${file}, ${who}: ${edit} at ${place}`;
          test(label, async () => {
            const { room, olive, ben } = await seats(bytes, name === 'Olive' ? mode : 'edit');
            const person = name === 'Olive' ? olive : ben;
            const alone = aloneShape(bytes, name, mode, [line, offset, offset], act);

            // Enter strictly inside the combo box: the editor alone leaves the prompt as it is,
            // and the room refuses it by name, as inside any inline prompt
            const refused = edit === 'Enter' && offset === 15;
            if (refused) {
              expect(alone).toBe(unchanged);
            }

            caretAt(person, line, offset);
            act(person);
            await settle();
            converged(room, [olive, ben], refused ? ['break-in-container'] : []);
            await everywhere([olive, ben], alone);
            await folds(room, bytes, label);
          });
        }
      }
    }
  }
});
