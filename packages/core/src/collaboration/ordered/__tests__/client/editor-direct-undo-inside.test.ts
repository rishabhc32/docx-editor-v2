// The organiser's undo and redo of direct edits made inside a party's pending suggestion (task
// 9.12): a format inside the party's insertion, a property of the party's new paragraph, the
// party's new row deleted. A direct edit lays its own settled record on those units, so its undo
// turns that record off exactly and leaves the party's suggestion as it was; a decision on the
// party's suggestion gives one result whether it comes before or after the undo. Words typed in
// the party's new cell are part of it, as in a new row: they once had a record of their own in the
// room, which a fold does not keep, so a reader of the labels found them in the party's cell only
// after a fold. Real editors, one room; after every step both copies equal the room and both
// screens show their lists.

import { afterEach, describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../../../contracts/editor.ts';
import { planTableCommand } from '../../../../editor/table-command-plan.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
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

afterEach(closeAll);

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;

/** Two paragraphs and a table of two rows. */
const FILE = docx(
  p(t('The fee is 100 USD.')) +
    p(t('Signed here.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cell('Fee')}${cell('100')}</w:tr><w:tr>${cell('Due')}${cell('30')}</w:tr></w:tbl>` +
    p(t('End.'))
);

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
  if (first) for (const other of rest) expect(editorXml(other)).toBe(editorXml(first));
}

async function seats() {
  const room = localRoom(FILE);
  const org = await join(room, 'Org', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, org, ana };
}

async function press(person: Person, run: (s: ReturnType<typeof surface>) => unknown) {
  run(surface(person));
  await settle();
}

/** Run one of the editor's table commands at the caret. */
function tableCommand(person: Person, command: EditorCommand) {
  const s = surface(person);
  const plan = planTableCommand({
    command,
    part: s.session.part(),
    layout: s.layout(),
    storeRevision: s.session.packageRevision(),
    selection: s.state().selection,
    cellSelection: s.state().cellSelection,
    themeColors: s.session.documentThemeColors(),
    editable: s.session.editable,
    viewing: false,
  });
  return s.applyTableCommandPlan(plan);
}

/** The XML with runs that differ only in where an edit cut them joined again. */
const joined = (xml: string): string => xml.replaceAll('</w:t></w:r><w:r><w:t>', '');

const indexOf = (person: Person, text: string): number => {
  const index = editorTexts(person).findIndex((value) => value === text);
  if (index < 0) throw new Error(`no paragraph "${text}"`);
  return index;
};

/** The authors of the suggestions that hold the room's units of `text`, by their own label. */
const holdersOf = (room: OrderedLocalRoom, text: string): string[] => {
  const doc = room.room.document;
  return [...(doc.stories.get(BODY)?.tree.items() ?? [])]
    .filter((item) => item.text !== '' && text.includes(item.text))
    .map((item) => doc.suggestions.get(item.by ?? '')?.attribution.author ?? 'none');
};

/** Ana types "XYZ" at the start of the second paragraph: her pending insertion. */
async function anaTypes(ana: Person): Promise<void> {
  caretAt(ana, 1, 0);
  surface(ana).insertPlainText('XYZ');
  await settle();
}

describe("the organiser's undo of a direct edit inside a party's suggestion", () => {
  test("a format inside the party's insertion goes, and comes back", async () => {
    const { room, org, ana } = await seats();
    await anaTypes(ana);
    const before = editorXml(org);
    caretAt(org, 1, 0, 2);
    await press(org, (s) => s.toggleRunProperty('b'));
    expect(editorXml(ana)).toMatch(
      /<w:ins w:author="Ana"[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>XY/
    );
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    // Ana's insertion is as she typed it: no bold, no record (the format's cut stays)
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(editorXml(ana)).toContain('<w:b/>');
    converged(room, org, ana);
  });

  test("a property of the party's new paragraph goes, and comes back", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, 0, 7);
    await press(ana, (s) => s.splitParagraph());
    const before = editorXml(org);
    caretAt(org, 0, 2);
    await press(org, (s) => s.setParagraphProperty('jc', { val: 'center' }));
    expect(editorXml(ana)).toContain('<w:jc w:val="center"/>');
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(editorXml(ana)).toContain('<w:jc w:val="center"/>');
    converged(room, org, ana);
  });

  test("the party's new row, deleted, comes back as the party's suggestion", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
    await settle();
    const before = editorXml(org);
    const rows = (person: Person) => editorXml(person).split('<w:tr>').length - 1;
    expect(rows(org)).toBe(3);
    // The organiser's caret in the party's new row: its first cell follows "100"
    caretAt(org, indexOf(org, '100') + 1, 0);
    expect(tableCommand(org, { type: 'deleteRow' }).ok).toBe(true);
    await settle();
    expect(rows(ana)).toBe(2);
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(rows(ana)).toBe(2);
    converged(room, org, ana);
  });

  test("words typed in the party's new cell are part of it, and go with it", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    const before = editorXml(org);

    // The organiser's caret in the party's new cell, right of "Fee"
    caretAt(org, indexOf(org, 'Fee') + 1, 0);
    await press(org, (s) => s.insertPlainText('XY'));
    expect(holdersOf(room, 'XY')).toEqual(['Ana']);
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(holdersOf(room, 'XY')).toEqual(['Ana']);
    converged(room, org, ana);

    // A reject of the party's column takes the words with it
    const column = [...room.room.document.suggestions.values()].find(
      (s) => s.attribution.author === 'Ana'
    );
    expect(org.handle.session.requestDecision(column?.id ?? '', 'rejected')).not.toBeNull();
    await settle();
    expect(editorTexts(ana)).not.toContain('XY');
    converged(room, org, ana);
  });

  test("words pasted in the party's new cell are part of it", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    const before = editorXml(org);
    caretAt(org, indexOf(org, 'Fee') + 1, 0);
    await press(org, (s) => s.pasteRich('XY', '<b>XY</b>'));
    expect(room.room.joinSnapshot(0).at(-1)?.kind).toBe('insert-fragment');
    expect(holdersOf(room, 'XY')).toEqual(['Ana']);
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(holdersOf(room, 'XY')).toEqual(['Ana']);
    converged(room, org, ana);
  });

  test("a footnote made in the party's new cell: its reference is part of it", async () => {
    const { room, org, ana } = await seats();
    caretAt(ana, indexOf(ana, 'Fee'), 0);
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    const before = editorXml(org);
    caretAt(org, indexOf(org, 'Fee') + 1, 0);
    await press(org, (s) => s.insertNote('footnote'));
    expect(holdersOf(room, '\uFFFC')).toEqual(['Ana']);
    converged(room, org, ana);
    await press(org, (s) => s.undo());
    expect(joined(editorXml(org))).toBe(joined(before));
    converged(room, org, ana);
    await press(org, (s) => s.redo());
    expect(holdersOf(room, '\uFFFC')).toEqual(['Ana']);
    converged(room, org, ana);
  });

  test("a decision on the party's suggestion gives one result before or after the undo", async () => {
    const results: string[] = [];
    for (const decisionFirst of [true, false]) {
      const { room, org, ana } = await seats();
      await anaTypes(ana);
      caretAt(org, 1, 1, 3);
      await press(org, (s) => s.toggleRunProperty('i'));
      const suggestion = [...room.room.document.suggestions.values()].find(
        (s) => s.attribution.author === 'Ana'
      )?.id;
      if (!suggestion) throw new Error('no suggestion of Ana');
      const decide = async () => {
        expect(org.handle.session.requestDecision(suggestion, 'accepted')).not.toBeNull();
        await settle();
      };
      if (decisionFirst) await decide();
      await press(org, (s) => s.undo());
      if (!decisionFirst) await decide();
      converged(room, org, ana);
      expect(editorTexts(ana)[1]).toBe('XYZSigned here.');
      results.push(editorXml(ana));
      closeAll();
    }
    expect(results[0]).toBe(results[1]);
    expect(results[0]).not.toContain('<w:i/>');
    expect(results[0]).not.toContain('<w:ins ');
  });

  test("the undo and the party's typing in the same insertion at once end with the room", async () => {
    const { room, org, ana } = await seats();
    await anaTypes(ana);
    caretAt(org, 1, 0, 3);
    await press(org, (s) => s.toggleRunProperty('b'));
    // Both go out before either copy hears of the other
    caretAt(ana, 1, 3);
    surface(ana).insertPlainText('W');
    surface(org).undo();
    await settle();
    expect(editorTexts(org)[1]).toBe('XYZWSigned here.');
    expect(editorXml(org)).not.toContain('<w:b/>');
    converged(room, org, ana);
  });
});
