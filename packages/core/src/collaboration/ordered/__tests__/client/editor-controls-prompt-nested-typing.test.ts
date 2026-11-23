// A control's prompt that holds a combo box showing its own prompt, with real editors and one
// ordered room, against the editor alone: "Courts at [Choose a city.] have jurisdiction.".
// The editor alone replaces the outer prompt whole, wherever the caret is in it, and drops the
// combo box with it (`clearPlaceholder`). The room once replaced the inner prompt, kept the
// combo box, and put every key from the second one beside the control, so the page and the
// editor alone differed, and a page could type into the combo box or refuse its own copy. Each
// step is checked on every copy, in a save and reopen, and after one and two folds.

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

/** The venue control showing its prompt, with the city combo box showing its own inside it. */
const FILE = docx(
  p(
    t('Venue: ') +
      '<w:sdt><w:sdtPr><w:alias w:val="Venue"/><w:id w:val="9203"/><w:showingPlcHdr/></w:sdtPr>' +
      `<w:sdtContent>${t('Courts at ', FACE)}<w:sdt><w:sdtPr><w:alias w:val="City"/>` +
      '<w:id w:val="9204"/><w:showingPlcHdr/><w:comboBox><w:listItem w:value=""/>' +
      '<w:listItem w:displayText="Mumbai" w:value="Mumbai"/></w:comboBox></w:sdtPr>' +
      `<w:sdtContent>${t('Choose a city.', FACE)}</w:sdtContent></w:sdt>` +
      `${t(' have jurisdiction.', FACE)}</w:sdtContent></w:sdt>`
  ) + p(t('Signed.'))
);

/** The file as a reader sees it. */
const UNCHANGED = 'Venue: [#9203?Courts at [#9204?Choose a city.] have jurisdiction.]¶Signed.¶';

/** Places in the first paragraph: from and to, as offsets. */
const PLACES = [
  ['the start of the outer prompt', 7, 7],
  ['the middle of the outer prompt', 12, 12],
  ['the end of the outer prompt', 50, 50],
  ['the whole outer prompt', 7, 50],
  ['the start of the inner prompt', 17, 17],
  ['the middle of the inner prompt', 22, 22],
  ['the end of the inner prompt', 31, 31],
  ['the whole inner prompt', 17, 31],
] as const;

/** The edits, as the keyboard and the clipboard make them. */
const EDITS = [
  [
    'typing',
    (person: Person) => {
      surface(person).type('Q');
      surface(person).type('R');
    },
  ],
  ['Enter', (person: Person) => surface(person).splitParagraph()],
  ['a paste', (person: Person) => surface(person).pasteRich('Q', '<p>Q</p>')],
  ['Backspace', (person: Person) => surface(person).deleteBackward()],
  ['Delete', (person: Person) => surface(person).deleteForward()],
] as const;

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

/** The organiser, Olive, in `mode`, and Ben, a party, in one room. */
async function seats(mode: 'edit' | 'suggesting') {
  const review = reviewTestModule();
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode, login: ORGANISER, wordFile: FILE, review });
  const ben = await join(room, 'Ben', { wordFile: FILE, review });
  return { room, olive, ben };
}

/** Each key on its own, the room's answer before the next one, as a person types. */
async function typeKeys(person: Person, text: string): Promise<void> {
  for (const key of text) {
    surface(person).type(key);
    await settle();
  }
}

async function folds(room: ReturnType<typeof localRoom>, label: string): Promise<void> {
  const { base } = openOrderedWordFile(FILE, { file: 1 });
  await checkFold(room.room.document, base, mediaStore(base), label);
}

describe('typing over a prompt that holds a control with its own prompt', () => {
  test('the organiser, editing, types over the outer prompt key by key', async () => {
    const { room, olive, ben } = await seats('edit');
    const typed = ' exclusively';
    const expected = aloneShape(FILE, 'Olive', 'edit', [0, 7, 50], (alone) => {
      for (const key of typed) surface(alone).type(key);
    });
    expect(expected).toBe('Venue: [#9203 exclusively]¶Signed.¶');
    caretAt(olive, 0, 7, 50);
    await typeKeys(olive, typed);
    converged(room, [olive, ben]);
    await everywhere([olive, ben], expected);
    await folds(room, 'typed');

    // The organiser's undo brings the prompt back with the combo box, and redo types again
    for (let step = 0; step < 2; step += 1) {
      surface(olive).undo();
      await settle();
    }

    converged(room, [olive, ben]);
    await everywhere([olive, ben], UNCHANGED);
    surface(olive).redo();
    await settle();
    converged(room, [olive, ben]);
    expect(shapeOf(editorXml(ben))).toBe('Venue: [#9203 ]¶Signed.¶');
  });

  test('a party suggests a city over the inner prompt key by key', async () => {
    const { room, olive, ben } = await seats('edit');
    const expected = aloneShape(FILE, 'Ben', 'suggesting', [0, 17, 31], (alone) => {
      for (const key of 'Pune') surface(alone).type(key);
    });
    expect(expected).toBe('Venue: [#9203{+Ben:Pune}]¶Signed.¶');
    caretAt(ben, 0, 17, 31);
    await typeKeys(ben, 'Pune');
    converged(room, [olive, ben]);
    await everywhere([olive, ben], expected);
    await folds(room, 'suggested');

    // A reject shows both prompts again; the same typing accepted keeps the city
    const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    expect(olive.editor.rejectReviewItem(card?.key ?? '').ok).toBe(true);
    await settle();
    converged(room, [olive, ben]);
    await everywhere([olive, ben], UNCHANGED);
    caretAt(ben, 0, 17, 31);
    await typeKeys(ben, 'Pune');
    const [again] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    expect(olive.editor.acceptReviewItem(again?.key ?? '').ok).toBe(true);
    await settle();
    converged(room, [olive, ben]);
    await everywhere([olive, ben], 'Venue: [#9203Pune]¶Signed.¶');
    await folds(room, 'accepted');
  });

  test('the organiser and a party type in the two prompts at the same time', async () => {
    const { room, olive, ben } = await seats('edit');
    caretAt(olive, 0, 50);
    caretAt(ben, 0, 20);

    // Both keys go out before either copy hears of the other: the organiser's comes first
    surface(olive).type('X');
    surface(ben).type('Y');
    await settle();
    converged(room, [olive, ben]);
    await everywhere([olive, ben], 'Venue: {+Ben:Y}[#9203X]¶Signed.¶');
    await folds(room, 'both');
  });
});

describe('each edit at each place of the two prompts, as the editor alone', () => {
  for (const [who, name, mode] of PEOPLE) {
    for (const [place, from, to] of PLACES) {
      for (const [edit, act] of EDITS) {
        const label = `${who}: ${edit} at ${place}`;
        test(label, async () => {
          const { room, olive, ben } = await seats(name === 'Olive' ? mode : 'edit');
          const person = name === 'Olive' ? olive : ben;
          const alone = aloneShape(FILE, name, mode, [0, from, to], act);

          // Enter inside an inline prompt is no split, in the editor alone too: refused by name
          const refused = edit === 'Enter' && from > 7 && to < 50;

          // The organiser's direct Backspace or Delete of a prompt's character: the editor alone
          // deletes it; the room keeps the prompt, as Word changes a prompt only in Design Mode
          const intoPrompt = (edit === 'Backspace' && from > 7) || (edit === 'Delete' && from < 50);
          const keepsPrompt = mode === 'edit' && from === to && intoPrompt;
          const expected = keepsPrompt ? UNCHANGED : alone;
          caretAt(person, 0, from, to);
          act(person);
          await settle();
          converged(room, [olive, ben], refused ? ['break-in-container'] : []);
          await everywhere([olive, ben], expected);
          await folds(room, label);
        });
      }
    }
  }
});
