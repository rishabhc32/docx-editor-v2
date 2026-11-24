// A party's edit of each kind of control, with real editors and one ordered room, against the
// editor alone: typing over a prompt, a pick in a list or a date picker, a check box toggled, and
// typing in the middle of a value, each accepted and rejected live, after a save and reopen, and
// after one and two folds. A control's own record of its value (`w:lastValue`, `w:fullDate`,
// `w14:checked`) is what the editor alone writes. The room once wrote a combo box's typed text
// as its last value while the typing was pending: a fold made that value the file's own, so a
// reject after the fold kept it where the same reject live left none.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { checkpointBytes, folded, mediaStore } from '../fold-check.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const FACE = '<w:rStyle w:val="PlaceholderText"/>';
const GLYPH_FONT = '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/>';

/** "Term: <control> end." with the control's properties and content. */
const line = (props: string, content: string) =>
  docx(
    p(
      t('Term: ') +
        `<w:sdt><w:sdtPr><w:id w:val="5"/>${props}</w:sdtPr><w:sdtContent>${content}` +
        `</w:sdtContent></w:sdt>${t(' end.')}`
    ) + p(t('Signed.'))
  );

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/><w:listItem w:displayText="Net 60" w:value="60"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';
const CHECK_BOX =
  '<w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
  '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>';

/** Type `text` key by key. */
const typing = (text: string) => (person: Person) => {
  for (const key of text) {
    surface(person).type(key);
  }
};

/** Pick `value` in the control's widget. */
const pick = (value: string) => (person: Person) => {
  const id = surface(person).layout().contentControls?.[0]?.id ?? '';
  expect(surface(person).contentControls.setValue(id, value)).toBe(true);
};

/** Each case: its label, the Word file, the caret in the first paragraph, and the party's edit. */
const CASES = [
  [
    "typing over a combo box's prompt",
    line(`<w:showingPlcHdr/><w:comboBox>${ITEMS}</w:comboBox>`, t('Choose an item.', FACE)),
    [6, 21],
    typing('unless'),
  ],
  [
    "an item over a combo box's prompt",
    line(`<w:showingPlcHdr/><w:comboBox>${ITEMS}</w:comboBox>`, t('Choose an item.', FACE)),
    [6, 6],
    pick('60'),
  ],
  [
    "an item over a drop-down list's prompt",
    line(`<w:showingPlcHdr/><w:dropDownList>${ITEMS}</w:dropDownList>`, t('Choose an item.', FACE)),
    [6, 6],
    pick('60'),
  ],
  [
    "a date over a date picker's prompt",
    line(`<w:showingPlcHdr/><w:date>${DATE_FORMAT}</w:date>`, t('Click to enter a date.', FACE)),
    [6, 6],
    pick('2026-03-04'),
  ],
  ['a check box ticked', line(CHECK_BOX, t('☐', GLYPH_FONT)), [6, 6], pick('true')],
  [
    "typing over a plain text control's prompt",
    line('<w:showingPlcHdr/><w:text/>', t('Click here to enter text.', FACE)),
    [6, 31],
    typing('Acme'),
  ],
  [
    "typing over a rich text control's prompt",
    line('<w:showingPlcHdr/>', t('Click here to enter text.', FACE)),
    [6, 31],
    typing('Acme'),
  ],
  [
    "typing in the middle of a combo box's value",
    line(`<w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`, t('Net 30')),
    [9, 9],
    typing('-day'),
  ],
  [
    "typing in the middle of a date picker's date",
    line(`<w:date w:fullDate="2026-01-02T00:00:00Z">${DATE_FORMAT}</w:date>`, t('1/2/2026')),
    [8, 8],
    typing('2'),
  ],
  [
    "typing in the middle of a plain text control's value",
    line('<w:text/>', t('Acme Ltd')),
    [10, 10],
    typing('ish'),
  ],
] as const;

/** The control as a reader of the file sees it: its properties and content, no ids or dates. */
const controlOf = (xml: string): string =>
  (/<w:sdt>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? 'no control')
    .replace(/ w:(id|date|rsidR|rsidRPr|rsidDel)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '')
    .replace(/<w:placeholder>.*?<\/w:placeholder>/g, '')
    .replace(/ xml:space="preserve"|<w:rPr>.*?<\/w:rPr>/g, '')
    .replace(/<\/w:r><w:r>/g, '')
    .replace(/<\/w:t><w:t>/g, '');

type Decision = 'accept' | 'reject';

/** Decide the one card as the organiser. */
async function decide(olive: Person, decision: Decision): Promise<void> {
  const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
  const result =
    decision === 'accept'
      ? olive.editor.acceptReviewItem(card?.key ?? '')
      : olive.editor.rejectReviewItem(card?.key ?? '');
  expect(result.ok).toBe(true);
  await settle();
}

/** What the editor alone writes: after the edit, and after the decision. */
function alone(
  file: Uint8Array,
  [from, to]: readonly [number, number],
  act: (person: Person) => void,
  decision: Decision
): { readonly edited: string; readonly decided: string } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Ben',
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    caretAt(person, 0, from, to);
    act(person);
    surface(person).flushPendingInput();
    const edited = controlOf(editorXml(person));
    const [card] = editor.getReviewItems().filter((item) => item.kind === 'revision');
    const result =
      decision === 'accept'
        ? editor.acceptReviewItem(card?.key ?? '')
        : editor.rejectReviewItem(card?.key ?? '');
    expect(result.ok).toBe(true);
    return { edited, decided: controlOf(editorXml(person)) };
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The organiser, Olive, and Ben, a party, in a room on `file` (a Word file or a checkpoint). */
async function seats(file: Uint8Array, number: number) {
  const review = reviewTestModule();
  const room = localRoom(file, { file: number });
  const olive = await join(room, 'Olive', { login: ORGANISER, wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  return { room, olive, ben };
}

/** The room's control, the same on both copies. */
function roomControl(seated: Awaited<ReturnType<typeof seats>>): string {
  for (const person of [seated.olive, seated.ben]) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, seated.room) }).toEqual({
      person: person.name,
      room: true,
    });
  }

  return controlOf(editorXml(seated.olive));
}

/** The checkpoint of the room's document folded `times` times. */
async function foldedBytes(
  doc: OrderedDocument,
  file: Uint8Array,
  times: number
): Promise<Uint8Array> {
  const { base } = openOrderedWordFile(file, { file: 1 });
  const blobs = mediaStore(base);
  let copy = { document: doc, base };
  for (let fold = 1; fold < times; fold += 1) {
    copy = await folded(copy.document, copy.base, fold + 1, blobs);
  }

  return checkpointBytes(copy.document, copy.base, blobs);
}

describe("a party's edit of a control, decided live and after a fold, as the editor alone", () => {
  for (const [label, file, [from, to], act] of CASES) {
    for (const decision of ['accept', 'reject'] as const) {
      test(`${label}, ${decision}ed`, async () => {
        const expected = alone(file, [from, to], act, decision);
        const seated = await seats(file, 1);
        caretAt(seated.ben, 0, from, to);
        act(seated.ben);
        await settle();
        expect({ label, edited: roomControl(seated) }).toEqual({ label, edited: expected.edited });

        // A save reopened, and the room folded once and twice, each decided by the organiser
        const saved = new Uint8Array(await seated.olive.editor.save());
        const copies = [
          ['saved', saved],
          ['folded once', await foldedBytes(seated.room.room.document, file, 1)],
          ['folded twice', await foldedBytes(seated.room.room.document, file, 2)],
        ] as const;
        await decide(seated.olive, decision);
        expect({ label, live: roomControl(seated) }).toEqual({ label, live: expected.decided });
        for (const [copy, bytes] of copies) {
          const reopened = await seats(bytes, 2);
          expect({ label, copy, edited: roomControl(reopened) }).toEqual({
            label,
            copy,
            edited: expected.edited,
          });
          await decide(reopened.olive, decision);
          const decided = serializeOoxmlPart(
            composeOrderedPart(reopened.room.room.document, BODY).part
          );
          expect({ label, copy, decided: controlOf(decided) }).toEqual({
            label,
            copy,
            decided: expected.decided,
          });
        }
      });
    }
  }
});
