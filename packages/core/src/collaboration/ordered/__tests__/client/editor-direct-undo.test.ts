// The organiser's undo and redo of direct edits (task 6.7), with real editors: the organiser
// edits with tracking off and a party suggests in the same room. A typing run is one step, as the
// editor alone groups it (a pause or a caret move ends it); undo takes back typing, a deletion,
// Enter, a join and a format exactly, and redo makes them again. The room once joined every key
// typed at one caret into one step, where the editor alone ends a run after a pause.
// A party's suggestion inside the text the undo takes back or shows again stays as it is, on
// both editors. After every step both copies equal the room and both screens show their lists.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  FILE,
  join,
  localRoom,
  HeldTransport,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  typeKeys,
  type Person,
} from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

function converged(room: OrderedLocalRoom, ...people: Person[]): void {
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
  const [first, ...rest] = people;
  if (first) for (const other of rest) expect(editorTexts(other)).toEqual(editorTexts(first));
}

async function seats() {
  const room = localRoom();
  const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER });
  const ana = await join(room, 'Ana');
  return { room, org, ana };
}

async function undo(person: Person): Promise<void> {
  surface(person).undo();
  await settle();
}

async function redo(person: Person): Promise<void> {
  surface(person).redo();
  await settle();
}

describe("the organiser's undo of direct edits", () => {
  test('a typing run is one step; undo takes it back and redo makes it again', async () => {
    const { room, org, ana } = await seats();
    caretAt(org, 0, 4);
    await typeKeys(org, 'abc');
    expect(editorTexts(ana)[0]).toBe('The abcfee is 100 USD.');
    expect(surface(org).state().canUndo).toBe(true);
    await undo(org);
    expect(editorTexts(org)[0]).toBe('The fee is 100 USD.');
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(org)[0]).toBe('The abcfee is 100 USD.');
    converged(room, org, ana);
    // Undo after redo takes the same run back again
    await undo(org);
    expect(editorTexts(ana)[0]).toBe('The fee is 100 USD.');
    converged(room, org, ana);
    // Direct text never shows as a suggestion
    expect(editorXml(org)).not.toContain('<w:ins ');
  });

  test('typing undoes in the steps the editor alone takes: a pause ends a run', async () => {
    let now = 0;
    const clock = spyOn(performance, 'now').mockImplementation(() => now);
    const typing = async (person: Person) => {
      caretAt(person, 0, 4);
      await typeKeys(person, 'ab');
      now = 1500;
      await typeKeys(person, 'cd');
    };
    try {
      const container = document.createElement('div');
      document.body.append(container);
      const editor = createDocxEditor({ container, document: FILE, author: 'Org', mode: 'edit' });
      const alone = { name: 'alone', editor, container } as unknown as Person;
      await typing(alone);
      const trail = async (person: Person) => {
        const texts: string[] = [];
        for (const step of [undo, undo, redo, redo]) {
          await step(person);
          texts.push(editorTexts(person)[0] ?? '');
        }

        return texts;
      };
      const expected = await trail(alone);
      editor.destroy();
      container.remove();
      expect(expected).toEqual([
        'The abfee is 100 USD.',
        'The fee is 100 USD.',
        'The abfee is 100 USD.',
        'The abcdfee is 100 USD.',
      ]);

      now = 0;
      const { room, org, ana } = await seats();
      await typing(org);
      expect(await trail(org)).toEqual(expected);
      converged(room, org, ana);
    } finally {
      clock.mockRestore();
    }
  });

  test("a party's insertion inside the organiser's text stays when the text goes", async () => {
    const { room, org, ana } = await seats();
    caretAt(org, 0, 4);
    await typeKeys(org, 'abc');
    caretAt(ana, 0, 5);
    surface(ana).insertPlainText('X');
    await settle();
    expect(editorTexts(org)[0]).toBe('The aXbcfee is 100 USD.');
    await undo(org);
    expect(editorTexts(org)[0]).toBe('The Xfee is 100 USD.');
    expect(editorXml(org)).toMatch(/<w:ins w:author="Ana"[^>]*><w:r><w:t>X<\/w:t>/);
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(ana)[0]).toBe('The aXbcfee is 100 USD.');
    converged(room, org, ana);
  });

  test("a deletion over a party's insertion shows it again as the party's", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, 1, 3);
    surface(ana).insertPlainText('XY');
    await settle();
    expect(editorTexts(org)[1]).toBe('DueXY in 30 days.');
    caretAt(org, 1, 0, 9);
    surface(org).deleteSelection();
    await settle();
    expect(editorTexts(ana)[1]).toBe('30 days.');
    converged(room, org, ana);
    await undo(org);
    expect(editorTexts(ana)[1]).toBe('DueXY in 30 days.');
    // Ana's text is her pending insertion again
    expect(editorXml(ana)).toMatch(/<w:ins w:author="Ana"[^>]*><w:r><w:t>XY<\/w:t>/);
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(org)[1]).toBe('30 days.');
    converged(room, org, ana);
  });

  test('a deletion of text a party struck shows it again, still struck by the party', async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, 2, 0, 6);
    surface(ana).deleteSelection();
    await settle();
    expect(editorXml(org)).toContain('<w:delText>Signed</w:delText>');
    caretAt(org, 2, 0, 7);
    surface(org).deleteSelection();
    await settle();
    expect(editorTexts(ana)[2]).toBe('here.');
    await undo(org);
    expect(editorTexts(ana)[2]).toBe('Signed here.');
    expect(editorXml(ana)).toMatch(
      /<w:del w:author="Ana"[^>]*><w:r><w:delText>Signed<\/w:delText>/
    );
    converged(room, org, ana);
  });

  test('Enter joins back and a join splits again; redo makes each again', async () => {
    const { room, org, ana } = await seats();
    caretAt(org, 0, 7);
    surface(org).splitParagraph();
    await settle();
    expect(editorTexts(ana).slice(0, 2)).toEqual(['The fee', ' is 100 USD.']);
    await undo(org);
    expect(editorTexts(ana)).toEqual(['The fee is 100 USD.', 'Due in 30 days.', 'Signed here.']);
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(ana).slice(0, 2)).toEqual(['The fee', ' is 100 USD.']);
    converged(room, org, ana);
    // Backspace at the start of a paragraph joins it; undo splits it again
    caretAt(org, 3, 0);
    surface(org).deleteBackward();
    await settle();
    expect(editorTexts(ana)[2]).toBe('Due in 30 days.Signed here.');
    await undo(org);
    expect(editorTexts(ana).slice(2)).toEqual(['Due in 30 days.', 'Signed here.']);
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(org)[2]).toBe('Due in 30 days.Signed here.');
    converged(room, org, ana);
  });

  test('a format goes off exactly: the property is gone, not turned off', async () => {
    const { room, org, ana } = await seats();
    caretAt(org, 0, 4, 7);
    surface(org).toggleRunProperty('b');
    await settle();
    expect(editorXml(ana)).toContain('<w:b/>');
    await undo(org);
    // The text keeps the runs the format cut, with no run properties at all
    expect(editorTexts(org)[0]).toBe('The fee is 100 USD.');
    expect(editorXml(org)).not.toContain('<w:rPr');
    converged(room, org, ana);
    await redo(org);
    expect(editorXml(ana)).toContain('<w:b/>');
    converged(room, org, ana);
  });

  test("text typed inside a party's insertion goes; the party's text stays", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, 2, 0);
    surface(ana).insertPlainText('XYZ');
    await settle();
    caretAt(org, 2, 1);
    await typeKeys(org, 'q');
    expect(editorTexts(ana)[2]).toBe('XqYZSigned here.');
    await undo(org);
    expect(editorTexts(ana)[2]).toBe('XYZSigned here.');
    converged(room, org, ana);
    await redo(org);
    expect(editorTexts(org)[2]).toBe('XqYZSigned here.');
    converged(room, org, ana);
  });

  test("an undo and a party's typing inside the same text at once end with the room", async () => {
    const { room, org, ana } = await seats();
    caretAt(org, 0, 4);
    await typeKeys(org, 'abc');
    // Both go out before either copy hears of the other
    caretAt(ana, 0, 6);
    surface(ana).insertPlainText('X');
    surface(org).undo();
    await settle();
    expect(editorTexts(org)[0]).toBe('The Xfee is 100 USD.');
    converged(room, org, ana);
    // The room took the undo first: Ana's place was hidden then, so her text went where the
    // view showed it (D6: before the next unit that shows), and the redo shows the run after it
    await redo(org);
    expect(editorTexts(ana)[0]).toBe('The Xabcfee is 100 USD.');
    converged(room, org, ana);
  });

  test("a deletion, a party's typing inside it at once, and the undo end with the room", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, 1, 4);
    caretAt(org, 1, 0, 9);
    surface(org).deleteSelection();
    surface(ana).insertPlainText('Q');
    await settle();
    // The room ordered the deletion first: Ana's place was hidden, so her text went before the
    // next unit that shows (D6), after the deleted text
    expect(editorTexts(ana)[1]).toBe('Q days.');
    converged(room, org, ana);
    await undo(org);
    expect(editorTexts(org)[1]).toBe('Due in 30Q days.');
    expect(editorXml(org)).toMatch(/<w:ins w:author="Ana"[^>]*><w:r><w:t>Q<\/w:t>/);
    converged(room, org, ana);
  });

  test("a redone Enter reaches a party's screen whole while the party's own edit is unconfirmed", async () => {
    // Seed 104 of the random sweep: the party's publish named the paragraph before the Enter,
    // its region took in the Enter's new paragraph and stopped at the next one, which kept the
    // words that had moved up: they showed twice on the party's screen
    const room = localRoom();
    const link = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
    const ana = await join(room, 'Ana', { transport: link });
    const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER });
    caretAt(org, 1, 6);
    surface(org).splitParagraph();
    await settle();
    await undo(org);
    // Ana's next frames wait: the organiser's redo reaches her while her typing is unconfirmed
    link.held = true;
    await redo(org);
    caretAt(ana, 0, 3);
    surface(ana).insertPlainText('Z');
    await settle();
    link.held = false;
    link.release();
    await settle();
    expect(editorTexts(ana)).toEqual([
      'TheZ fee is 100 USD.',
      'Due in',
      ' 30 days.',
      'Signed here.',
    ]);
    converged(room, org, ana);
  });
});
