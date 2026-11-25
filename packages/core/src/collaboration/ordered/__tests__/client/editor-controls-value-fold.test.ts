// A party's edit of each kind of control, with real editors and one ordered room, against the
// editor alone: typing over a prompt, a pick in a list or a date picker, a check box toggled, and
// typing in the middle of a value, each accepted and rejected live, after a save and reopen, and
// after one and two folds. A control's own record of its value (`w:lastValue`, `w:fullDate`,
// `w14:checked`) is what the editor alone writes. The room once wrote a combo box's typed text
// as its last value while the typing was pending: a fold made that value the file's own, so a
// reject after the fold kept it where the same reject live left none. A reject of a page field in
// a combo box's typed value once wrote that text as the value live, though a field's code reads as
// no text: the editor alone and a fold keep the value. A reject of typing over a prompt the file
// shows beside a stored value once took that value away in the editor alone and after a fold,
// where the room kept it: the value the file holds stays, and only one the words wrote goes. A date
// over another person's struck, formatted words once copied those words, with their format record,
// into the party's new value: a fold made the copied record that person's, and their format joined
// the party's card. The editor alone inserts only the words the value shows. A date in a block date
// picker once went after the control's paragraph mark, where no paragraph holds it, so no copy
// showed it; it goes in the control's paragraph, as the editor alone writes it. A value picked over
// a prompt beside a stored value once wrote its own record over the stored one, so a reject, and
// the organiser's undo of a direct pick in the room, showed the prompt with no value: the stored
// value now stays while a pick is pending, and an accept reads the new one from the words. A date
// picked over words that state no date, rejected, once kept the picked date's record in the
// editor alone and after a fold, and a date control with no format showed 2026-09-15 in the room.
// A date picked over a date in a format that shows no day, rejected, once kept the picked date in
// the editor alone and after a fold, where the live room gave the file's date back. An accepted
// deletion of a control's words that do not state its stored value once kept that value when the
// prompt showed again: the value goes, as Word reads it back from the prompt. A reject of typing
// over a plain prompt with the organiser's comment on part of the words once showed the prompt
// before the comment's range live, after it once folded: it goes inside the range, as the editor
// alone puts it. A party's format over his own on a prompt kept naming the change it replaced
// once his typing filled the prompt: no copy names it, as a fold reads the file. A reject of typing
// over a prompt formatted after its first fill showed that format live: the prompt the control
// kept at the first fill shows, as the editor alone and a fold show it. A comment on part of a
// plain prompt kept its markers between the prompt's words when a reject showed the prompt again
// live: the prompt goes whole inside the range, as the editor alone writes it. A date in a German
// date picker once showed English names: it shows the names of the control's language, as Word.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedReplacedSuggestions } from '../../compose/replaced.ts';
import { orderedCards } from '../../decide.ts';
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

/** "Term:", a block control with the properties and one paragraph of content, "Signed.". */
const block = (props: string, content: string) =>
  docx(
    p(t('Term:')) +
      `<w:sdt><w:sdtPr><w:id w:val="5"/>${props}</w:sdtPr><w:sdtContent>${p(content)}` +
      `</w:sdtContent></w:sdt>${p(t('Signed.'))}`
  );

const ITEMS =
  '<w:listItem w:displayText="Net 30" w:value="30"/><w:listItem w:displayText="Net 60" w:value="60"/>';
const DATE_FORMAT = '<w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/>';

/** A date picker in Word's German long form, with words that state no date. */
const GERMAN_DATE = line(
  '<w:date><w:dateFormat w:val="dddd, d. MMMM yyyy"/><w:lid w:val="de-DE"/></w:date>',
  t('bald')
);

/** Words of a date that Cy struck and formatted, as a file holds them. */
const STRUCK =
  '<w:del w:id="1" w:author="Cy" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:caps/>' +
  '<w:rPrChange w:id="2" w:author="Cy" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange>' +
  '</w:rPr><w:delText xml:space="preserve">Jan </w:delText></w:r></w:del>';
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

/** A page field after "Term: Net", through the editing API's lane: its code is no text. */
const pageField = (person: Person) => {
  const result = surface(person).applyAutomationOps(() => {
    const paragraphId = surface(person).session.paragraphIds()[0] ?? '';
    return [{ op: 'insertPageField', paragraphId, offset: 9, field: 'PAGE' }];
  });
  expect(result.committed).toBe(true);
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
  [
    "typing over a date picker's prompt with a stored date",
    line(
      `<w:showingPlcHdr/><w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`,
      t('Click to enter a date.', FACE)
    ),
    [6, 28],
    typing('soon'),
  ],
  [
    "a date over a date picker's prompt with a stored date",
    line(
      `<w:showingPlcHdr/><w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`,
      t('Click to enter a date.', FACE)
    ),
    [6, 6],
    pick('2026-03-04'),
  ],
  [
    "typing over a combo box's prompt with a stored value",
    line(
      `<w:showingPlcHdr/><w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`,
      t('Choose an item.', FACE)
    ),
    [6, 21],
    typing('unless'),
  ],
  [
    "an item over a drop-down list's prompt with a stored value",
    line(
      `<w:showingPlcHdr/><w:dropDownList w:lastValue="30">${ITEMS}</w:dropDownList>`,
      t('Choose an item.', FACE)
    ),
    [6, 6],
    pick('60'),
  ],
  [
    "the item a combo box's prompt stores, over the prompt",
    line(
      `<w:showingPlcHdr/><w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`,
      t('Choose an item.', FACE)
    ),
    [6, 6],
    pick('30'),
  ],
  [
    "the date a date picker's prompt stores, over the prompt",
    line(
      `<w:showingPlcHdr/><w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`,
      t('Click to enter a date.', FACE)
    ),
    [6, 6],
    pick('2026-06-30'),
  ],
  [
    'a date in a date picker with no format',
    line('<w:date><w:lid w:val="en-US"/></w:date>', t('Soon')),
    [6, 6],
    pick('2026-09-15'),
  ],
  ['a date in a German date picker', GERMAN_DATE, [6, 6], pick('2026-09-15')],
  [
    'a date over a date in a format that shows no day',
    line(
      '<w:date w:fullDate="2026-06-30T00:00:00Z"><w:dateFormat w:val="MMMM yyyy"/></w:date>',
      t('June 2026')
    ),
    [6, 6],
    pick('2026-09-15'),
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
    "a page field in a combo box's typed value",
    line(`<w:comboBox w:lastValue="60">${ITEMS}</w:comboBox>`, t('Net 60 days')),
    [9, 9],
    pageField,
  ],
  [
    "typing in the middle of a date picker's date",
    line(`<w:date w:fullDate="2026-01-02T00:00:00Z">${DATE_FORMAT}</w:date>`, t('1/2/2026')),
    [8, 8],
    typing('2'),
  ],
  [
    "a date over another person's struck, formatted words",
    line(
      `<w:date w:fullDate="2026-01-02T00:00:00Z">${DATE_FORMAT}</w:date>`,
      STRUCK + t('1/2/2026')
    ),
    [6, 6],
    pick('2026-03-04'),
  ],
  [
    'a date in a block date picker',
    block(`<w:date w:fullDate="2026-01-02T00:00:00Z">${DATE_FORMAT}</w:date>`, t('1/2/2026')),
    [0, 0],
    pick('2026-03-04'),
  ],
  [
    "a date over a block date picker's prompt with a stored date",
    block(
      `<w:showingPlcHdr/><w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`,
      t('Click to enter a date.', FACE)
    ),
    [0, 0],
    pick('2026-03-04'),
  ],
  [
    'a date over a date in a block date picker whose format shows no day',
    block(
      '<w:date w:fullDate="2026-06-30T00:00:00Z"><w:dateFormat w:val="MMMM yyyy"/></w:date>',
      t('June 2026')
    ),
    [0, 0],
    pick('2026-09-15'),
  ],
  [
    'a date over words that state no date in a block date picker',
    block('<w:date><w:lid w:val="en-US"/></w:date>', t('Soon')),
    [0, 0],
    pick('2026-09-15'),
  ],
  [
    "a date over another person's struck, formatted words in a block date picker",
    block(
      `<w:date w:fullDate="2026-01-02T00:00:00Z">${DATE_FORMAT}</w:date>`,
      STRUCK + t('1/2/2026')
    ),
    [0, 0],
    pick('2026-03-04'),
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

/** The party's card, as the page shows it. */
const bensCard = (person: Person) =>
  person.editor.getReviewItems().find((item) => item.kind === 'revision' && item.author === 'Ben');

/** The page's cards, each by its author. */
const pageCards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((item) => item.kind === 'revision')
    .map((item) => item.author)
    .sort();

/** The room's cards, each by the authors of its suggestions. */
const roomCards = (doc: OrderedDocument): string[] =>
  orderedCards(doc)
    .map((card) =>
      [...new Set(card.suggestions.map((id) => doc.suggestions.get(id)?.attribution.author))].join()
    )
    .sort();

/** Decide the party's card as the organiser. */
async function decide(olive: Person, decision: Decision): Promise<void> {
  const card = bensCard(olive);
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
): { readonly edited: string; readonly cards: string[]; readonly decided: string } {
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
    const cards = pageCards(person);
    const card = bensCard(person);
    const result =
      decision === 'accept'
        ? editor.acceptReviewItem(card?.key ?? '')
        : editor.rejectReviewItem(card?.key ?? '');
    expect(result.ok).toBe(true);
    return { edited, cards, decided: controlOf(editorXml(person)) };
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
        expect({ label, cards: roomCards(seated.room.room.document) }).toEqual({
          label,
          cards: expected.cards,
        });

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
          expect({ label, copy, cards: roomCards(reopened.room.room.document) }).toEqual({
            label,
            copy,
            cards: expected.cards,
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

// Word for Mac shows a de-DE date picture with these names (2026-11-25); the cases above check that
// the room, a save and a fold write what the editor alone writes
test("a date in a German date picker shows Word's German names", () => {
  expect(alone(GERMAN_DATE, [6, 6], pick('2026-09-15'), 'accept').decided).toContain(
    '<w:t>Dienstag, 15. September 2026</w:t>'
  );
});

describe('a value picked over a prompt the control shows beside a stored value', () => {
  const stored = line(
    `<w:showingPlcHdr/><w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`,
    t('Choose an item.', FACE)
  );

  test('a reject gives the stored value back and an accept the new one, live and folded', async () => {
    for (const [decision, record] of [
      ['reject', '<w:comboBox w:lastValue="30">'],
      ['accept', '<w:comboBox w:lastValue="60">'],
    ] as const) {
      const seated = await seats(stored, 1);
      pick('60')(seated.ben);
      await settle();
      const copy = await foldedBytes(seated.room.room.document, stored, 1);
      await decide(seated.olive, decision);
      const reopened = await seats(copy, 2);
      await decide(reopened.olive, decision);
      for (const decided of [roomControl(seated), roomControl(reopened)]) {
        expect({ decision, record: decided.includes(record) }).toEqual({ decision, record: true });
        expect(decided.includes('<w:showingPlcHdr/>')).toBe(decision === 'reject');
      }
    }
  });

  test("the organiser's direct pick undone gives the stored value back", async () => {
    const seated = await seats(stored, 1);
    seated.olive.editor.setMode('edit');
    pick('60')(seated.olive);
    await settle();
    surface(seated.olive).undo();
    await settle();
    const control = roomControl(seated);
    expect(control).toContain('<w:showingPlcHdr/><w:comboBox w:lastValue="30">');
    expect(control).toContain('Choose an item.');
  });
});

describe('an accepted deletion of the words of a control with a stored value', () => {
  // Words that do not state the value: an accept shows the prompt again, and the value goes, as
  // Word reads it back from the prompt
  const strike = (person: Person) => {
    surface(person).deleteSelection();
  };

  for (const [label, props] of [
    ['a combo box', `<w:comboBox w:lastValue="30">${ITEMS}</w:comboBox>`],
    ['a date picker', `<w:date w:fullDate="2026-06-30T00:00:00Z">${DATE_FORMAT}</w:date>`],
  ] as const) {
    test(`${label}: the prompt shows with no value, live, folded and alone`, async () => {
      const file = line(props, t('Custom'));
      const expected = alone(file, [6, 12], strike, 'accept');
      expect(expected.decided).toContain('<w:showingPlcHdr/>');
      expect(expected.decided).not.toMatch(/w:lastValue|w:fullDate/);
      const seated = await seats(file, 1);
      caretAt(seated.ben, 0, 6, 12);
      strike(seated.ben);
      await settle();
      const copy = await foldedBytes(seated.room.room.document, file, 1);
      await decide(seated.olive, 'accept');
      const reopened = await seats(copy, 2);
      await decide(reopened.olive, 'accept');
      expect({ live: roomControl(seated), folded: roomControl(reopened) }).toEqual({
        live: expected.decided,
        folded: expected.decided,
      });
    });
  }
});

describe("the organiser's comment on part of a party's typing over a plain prompt", () => {
  // The room writes the comment's markers outside the party's insertion: the editor alone decides
  // the room's own file
  const file = line('<w:showingPlcHdr/><w:text/>', t('Click here to enter text.', FACE));

  for (const decision of ['accept', 'reject'] as const) {
    test(`${decision}ed: the control as the editor alone leaves it, live, saved and folded`, async () => {
      const seated = await seats(file, 1);
      caretAt(seated.ben, 0, 6, 31);
      typing('Acme X')(seated.ben);
      await settle();
      caretAt(seated.olive, 0, 8, 10);
      expect(seated.olive.editor.addComment('Why?').ok).toBe(true);
      await settle();
      const saved = new Uint8Array(await seated.olive.editor.save());
      const expected = alone(saved, [0, 0], () => {}, decision).decided;
      expect(expected).toContain('<w:commentRangeStart/>');
      const copies = [
        saved,
        await foldedBytes(seated.room.room.document, file, 1),
        await foldedBytes(seated.room.room.document, file, 2),
      ];
      await decide(seated.olive, decision);
      const decided = [roomControl(seated)];
      for (const bytes of copies) {
        const reopened = await seats(bytes, 2);
        await decide(reopened.olive, decision);
        decided.push(roomControl(reopened));
      }

      expect(decided).toEqual([expected, ...copies.map(() => expected)]);
    });
  }
});

describe("a comment on part of a plain prompt, then a party's typing over it", () => {
  // The comment's markers stay in the control while the typing fills it; a decision shows the
  // control as the editor alone leaves the room's own file, live and folded
  const file = line('<w:showingPlcHdr/><w:text/>', t('Click here to enter text.', FACE));

  for (const decision of ['accept', 'reject'] as const) {
    test(`${decision}ed: the control as the editor alone leaves it, live and folded`, async () => {
      const seated = await seats(file, 1);
      caretAt(seated.olive, 0, 7, 17);
      expect(seated.olive.editor.addComment('Why?').ok).toBe(true);
      await settle();
      caretAt(seated.ben, 0, 6, 31);
      typing('Acme')(seated.ben);
      await settle();
      const saved = new Uint8Array(await seated.olive.editor.save());
      const expected = alone(saved, [0, 0], () => {}, decision).decided;
      const copies = [saved, await foldedBytes(seated.room.room.document, file, 1)];
      await decide(seated.olive, decision);
      const decided = [roomControl(seated)];
      for (const bytes of copies) {
        const reopened = await seats(bytes, 2);
        await decide(reopened.olive, decision);
        decided.push(roomControl(reopened));
      }

      expect(decided).toEqual([expected, ...copies.map(() => expected)]);
    });
  }
});

describe('a prompt formatted after its first fill, then typed over again', () => {
  // Olive's typing filled the prompt and was rejected; Ben then makes the prompt italic and types
  // over it. The control kept its prompt at the first fill: a reject of Ben's typing shows that
  // prompt, as the editor alone leaves the room's own file, live and folded
  test('the reject shows the prompt the control kept', async () => {
    const file = line(
      `<w:showingPlcHdr/><w:dropDownList>${ITEMS}</w:dropDownList>`,
      t('Choose an item.', FACE)
    );
    const seated = await seats(file, 1);
    caretAt(seated.olive, 0, 20);
    typing('x')(seated.olive);
    await settle();
    const olives = seated.olive.editor
      .getReviewItems()
      .find((item) => item.kind === 'revision' && item.author === 'Olive');
    expect(seated.olive.editor.rejectReviewItem(olives?.key ?? '').ok).toBe(true);
    await settle();
    caretAt(seated.ben, 0, 6, 20);
    seated.ben.editor.exec({ type: 'toggleMark', mark: 'italic' });
    caretAt(seated.ben, 0, 20);
    typing('a')(seated.ben);
    await settle();
    const saved = new Uint8Array(await seated.olive.editor.save());
    const expected = alone(saved, [0, 0], () => {}, 'reject').decided;
    const copy = await foldedBytes(seated.room.room.document, file, 1);
    await decide(seated.olive, 'reject');
    const reopened = await seats(copy, 2);
    await decide(reopened.olive, 'reject');
    expect([roomControl(seated), roomControl(reopened)]).toEqual([expected, expected]);
    expect(editorXml(seated.olive)).not.toContain('<w:i/>');
  });
});

describe("a party's format over his own on a prompt, then his typing over it", () => {
  // Ben's bold takes his italic's record on part of the prompt; his typing then fills the control,
  // so the prompt shows nowhere and no copy names the change the bold replaced, as a fold reads
  // the file
  test('no copy names the change it replaced', async () => {
    const file = line('<w:showingPlcHdr/><w:text/>', t('Click here to enter text.', FACE));
    const seated = await seats(file, 1);
    caretAt(seated.ben, 0, 6, 31);
    seated.ben.editor.exec({ type: 'toggleMark', mark: 'italic' });
    await settle();
    caretAt(seated.ben, 0, 2, 11);
    seated.ben.editor.exec({ type: 'toggleMark', mark: 'bold' });
    await settle();
    const doc = seated.room.room.document;
    const bold = [...doc.suggestions.values()].at(-1)?.id ?? '';
    const replaced = (copy: OrderedDocument) => orderedReplacedSuggestions(copy, bold).length;
    expect(replaced(doc)).toBe(1);
    caretAt(seated.ben, 0, 31);
    typing('Acme')(seated.ben);
    await settle();
    roomControl(seated);
    const { base } = openOrderedWordFile(file, { file: 1 });
    const copies = [
      doc,
      seated.olive.handle.session.orderedDocument(),
      seated.ben.handle.session.orderedDocument(),
      (await folded(doc, base, 2, mediaStore(base))).document,
    ];
    expect(copies.map(replaced)).toEqual([0, 0, 0, 0]);
  });
});
