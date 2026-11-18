// Whole tables and paragraphs removed from the editor in an ordered room (`deleteBlock`, task
// 9.10). A party's Delete Table is a tracked deletion of every row, as the engine writes it,
// in one suggestion the organiser decides; the organiser's is a direct edit that undo and redo
// turn back and again. Every copy equals the room, and the screen follows the list.

import { afterEach, describe, expect, test } from 'bun:test';
import { planTableCommand } from '../../../../editor/table-command-plan.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
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
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${p(t(text))}</w:tc>`;
const row = (a: string, b: string) => `<w:tr>${cell(a)}${cell(b)}</w:tr>`;
const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  `${row('A1', 'B1')}${row('A2', 'B2')}</w:tbl>`;
const FILE = docx(p(t('Before.')) + TABLE + p(t('After.')));

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

function tableCommand(person: Person, command: Readonly<Record<string, unknown>>) {
  const s = surface(person);
  const plan = planTableCommand({
    // The command shapes are the editor's; the test names them by their type
    command: command as never,
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

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

async function open(policy = true) {
  const room = localRoom(FILE, policy ? { policy: suggestionsOnlyPolicy() } : {});
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: FILE,
  });
  const ana = await join(room, 'Ana', { review, wordFile: FILE });
  return { room, olive, ana };
}

describe('Delete Table in an ordered room', () => {
  test("a party's Delete Table is one tracked suggestion over every row; the organiser decides it", async () => {
    const { room, olive, ana } = await open();
    caretAt(ana, 1, 0);
    expect(tableCommand(ana, { type: 'deleteTable' }).ok).toBe(true);
    await settle();
    expect(ana.refusals).toEqual([]);
    expect(kinds(room)).toEqual(['delete-row', 'delete-row']);
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      // Both rows struck under one revision id: one gesture is one decision
      const ids = [...editorXml(person).matchAll(/<w:trPr><w:del [^>]*w:id="(\d+)"/g)].map(
        (m) => m[1]
      );
      expect(ids).toHaveLength(2);
      expect(new Set(ids).size).toBe(1);
    }
    expect(cards(olive)).toHaveLength(1);
    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.acceptReviewItem(card).ok).toBe(true);
    await settle();
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorXml(person)).not.toContain('<w:tbl>');
      expect(editorTexts(person)).toEqual(['Before.', 'After.']);
    }
  });

  test("the organiser's Delete Table is a direct edit; undo and redo turn it back and again", async () => {
    const { room, olive, ana } = await open();
    caretAt(olive, 1, 0);
    expect(tableCommand(olive, { type: 'deleteTable' }).ok).toBe(true);
    await settle();
    expect(olive.refusals).toEqual([]);
    expect(kinds(room)).toEqual(['delete-row', 'delete-row']);
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)).toEqual(['Before.', 'After.']);
    }
    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(screenFollowsList(person)).toBe(true);
      expect(editorTexts(person)).toEqual(['Before.', 'A1', 'B1', 'A2', 'B2', 'After.']);
    }
    expect(olive.handle.session.redo()).toBe(true);
    await settle();
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(editorTexts(person)).toEqual(['Before.', 'After.']);
    }
  });

  test('the saved file reopens as the room holds it, with the suggestion still to decide', async () => {
    const { room, ana } = await open();
    caretAt(ana, 1, 0);
    tableCommand(ana, { type: 'deleteTable' });
    await settle();
    const base = readOoxmlPackage(FILE);
    if (!base.ok) throw new Error('no base');
    const bytes = orderedPackageBytes(room.room.document, base.package);
    expect(readOoxmlPackage(bytes).ok).toBe(true);
    const reopened = openOrderedWordFile(bytes, { file: 2 }).document;
    const pending = [...reopened.suggestions.values()].filter((s) => s.status === 'pending');
    expect(pending.map((s) => s.attribution.author)).toEqual(['Ana']);
  });
});
