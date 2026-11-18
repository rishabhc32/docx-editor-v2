// Undo of an edit of one's own pending words, with real editors and one room (audit R1 rows 5
// and 6 and their follow-ups; design D7, D20).
//
// A deletion over one's own new words takes them back: they never existed for anyone else.
// Typing over them, or a format on them, changes one's own suggestion. In the editor alone and
// in Word, an undo of such an edit takes back only that edit: the deleted part of the words
// comes back, a new note's reference and the note with it, the words typed over, the face from
// before the format. The room's undo took back the whole insertion, or left the note gone. Each
// step here is compared with the editor alone, for a party and for the organiser suggesting,
// and the file saved after an undo reopens as the room shows it.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { copiedBetween, type Copied } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { comparedBody, converged } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

type Step = (person: Person) => void | Promise<void>;

const undo: Step = (person) => surface(person).undo();
const redo: Step = (person) => surface(person).redo();

/**
 * Note reference ids in order of first appearance: a room's note made again takes an id past
 * the hidden one's (design, "Note ids"), the editor alone's the lowest free one.
 */
function renumberedNotes(xml: string): string {
  const ids = new Map<string, number>();
  return xml.replace(/(<w:footnoteReference w:id=")(\d+)/g, (_, head: string, id: string) => {
    if (!ids.has(id)) ids.set(id, ids.size + 1);
    return `${head}${ids.get(id)}`;
  });
}

/** What a person's editor shows: the main part as the matrix compares it, and the notes' words. */
function shownBy(person: Person): string {
  const pkg = surface(person).session.currentPackage();
  const notes = [...pkg.parts.keys()].find((name) => name.startsWith('/word/footnotes'));
  const words = notes === undefined ? '' : serializeOoxmlPart(pkg.parts.get(notes)!);
  const noteText = [...words.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map((m) => m[1]);
  const body = renumberedNotes(comparedBody(surface(person).session.part()));
  return `${body}\nnotes: ${noteText.join('|')}`;
}

/** The editor alone's result after each step. */
async function aloneSteps(author: string, steps: readonly Step[]): Promise<string[]> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author,
    mode: 'suggesting',
    modules: [review],
  });
  const me = { name: 'alone', editor, container } as unknown as Person;
  try {
    const shown: string[] = [];
    for (const step of steps) {
      await step(me);
      shown.push(shownBy(me));
    }
    return shown;
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The main part of a person's save, as the matrix compares it. */
async function savedBody(person: Person): Promise<string> {
  const read = readOoxmlPackage(new Uint8Array(await person.editor.save()), {});
  if (!read.ok) throw new Error(read.reason);
  return comparedBody(read.package.parts.get(BODY)!);
}

const WHO = [
  ['a party', 'Ana'],
  ['the organiser suggesting', 'Olive'],
] as const;

/** Run the steps in a room as `author` and compare each with the editor alone; the room. */
async function sameAsAlone(author: string, steps: readonly Step[]) {
  const expected = await aloneSteps(author, steps);
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { review });
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
  const person = author === 'Ana' ? ana : olive;
  person.editor.setMode('suggesting');
  const shown: string[] = [];
  for (const [k, step] of steps.entries()) {
    await step(person);
    await settle();
    converged(room, ana, olive);
    shown.push(shownBy(ana));
    // A save after each undo reopens as the room shows it
    if (step === undo)
      expect({ step: k, saved: await savedBody(olive) }).toEqual({
        step: k,
        saved: comparedBody(surface(olive).session.part()),
      });
  }
  expect(shown).toEqual(expected);
  return room;
}

const type: Step = (person) => {
  caretAt(person, 0, 0);
  surface(person).insertPlainText('abcdef ');
};
const deletePart: Step = (person) => {
  caretAt(person, 0, 2, 4);
  surface(person).deleteSelection();
};

/** Run the steps for each person and check that nothing is left pending at the end. */
function eachPerson(name: string, steps: readonly Step[]): void {
  for (const [who, author] of WHO)
    test(`${who}: ${name}`, async () => {
      const room = await sameAsAlone(author, steps);
      expect(orderedCards(room.room.document)).toEqual([]);
    });
}

describe('undo of a deletion of part of your own new words, as the editor alone (row 6)', () => {
  const backspace: Step = (person) => {
    caretAt(person, 0, 4);
    surface(person).deleteBackward();
  };
  eachPerson('type, delete a part, Backspace, then undo and redo', [
    type,
    deletePart,
    backspace,
    undo,
    undo,
    redo,
    redo,
    undo,
    undo,
    undo,
  ]);
});

describe('undo of typing over part of your own new words, as the editor alone', () => {
  const typeOver: Step = (person) => {
    caretAt(person, 0, 2, 4);
    surface(person).insertPlainText('XY');
  };
  eachPerson('type, type over a part, then undo and redo', [
    type,
    typeOver,
    undo,
    undo,
    redo,
    redo,
    undo,
    undo,
  ]);
});

describe('undo of a format on your own new words, as the editor alone', () => {
  const bold: Step = (person) => {
    caretAt(person, 0, 1, 4);
    surface(person).setRunProperty('b');
  };
  eachPerson('type, bold a part, delete a part, then undo and redo', [
    type,
    bold,
    deletePart,
    undo,
    undo,
    redo,
    redo,
    undo,
    undo,
    undo,
  ]);
});

describe("undo of a deletion over your own new note's reference, as the editor alone (row 5)", () => {
  const note: Step = async (person) => {
    caretAt(person, 0, 7);
    expect(surface(person).insertNote('footnote')).toBe(true);
    await settle();
    surface(person).insertPlainText('A note.');
    await settle();
    surface(person).exitNote();
  };
  // One character on each side of the reference: the file's own words, struck
  const deleteOver: Step = (person) => {
    caretAt(person, 0, 6, 9);
    surface(person).deleteSelection();
  };
  const STEPS = [note, deleteOver, undo, redo, undo];
  for (const [who, author] of WHO)
    test(`${who}: a note, a deletion over its reference, then undo and redo`, async () => {
      await sameAsAlone(author, STEPS);
    });
});

/** A copy's source: bold "XX", then "y", then "z" on the next line. */
const BOLD_LINES = docx(p(t('XX', '<w:b/>') + t('y')) + p(t('z')));

const pasteCopy = (s: ReturnType<typeof surface>, copied: Copied) =>
  s.pasteRich(copied.text, copied.html);

describe('undo of an edit inside your own new words, as the editor alone', () => {
  /** An edit with the caret inside the typed words. */
  const inside =
    (edit: (s: ReturnType<typeof surface>) => unknown): Step =>
    (person) => {
      caretAt(person, 0, 3);
      edit(surface(person));
    };
  const STEPS = (edit: Step) => [type, edit, undo, undo, redo, redo, undo, undo];
  eachPerson('type, Enter inside, then undo and redo', STEPS(inside((s) => s.splitParagraph())));
  eachPerson('type, Tab inside, then undo and redo', STEPS(inside((s) => s.insertTab())));
  eachPerson(
    'type, a line break inside, then undo and redo',
    STEPS(inside((s) => s.insertLineBreak()))
  );
  eachPerson(
    'type, a page break inside, then undo and redo',
    STEPS(inside((s) => s.insertPageBreak()))
  );
  eachPerson(
    'type, paste plain words inside, then undo and redo',
    STEPS(inside((s) => s.pasteRich('XY', null)))
  );
  eachPerson(
    'type, paste two plain lines inside, then undo and redo',
    STEPS(inside((s) => s.pasteRich('XY\nZW', null)))
  );
  eachPerson(
    'type, paste bold words inside, then undo and redo',
    STEPS(inside((s) => pasteCopy(s, copiedBetween(0, 0, 2, BOLD_LINES))))
  );
  eachPerson(
    'type, paste two bold lines inside, then undo and redo',
    STEPS(inside((s) => pasteCopy(s, copiedBetween(0, 1, 1, BOLD_LINES))))
  );
  const note =
    (text: string): Step =>
    async (person) => {
      caretAt(person, 0, 3);
      expect(surface(person).insertNote('footnote')).toBe(true);
      await settle();
      if (text) surface(person).insertPlainText(text);
      await settle();
      surface(person).exitNote();
    };
  eachPerson('type, a note inside, then undo and redo', STEPS(note('')));
  // The redo of words typed in a note made again is refused by name (`hidden-target`)
  eachPerson('type, a note with words inside, then undo', [
    type,
    note('A note.'),
    undo,
    undo,
    undo,
  ]);
});
