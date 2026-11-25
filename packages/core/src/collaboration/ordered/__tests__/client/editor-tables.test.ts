// Table edits with real editors and one room (task 9.7): columns inserted and deleted as a
// party's suggestion or as the organiser's direct edit, decided from the editor, saved and
// reopened, made by two people at once, undone and made again; column drags; new tables where
// the engine adds a paragraph or a pending Enter made one. After every step each copy equals the
// room and each screen shows its list.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  HeldTransport,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import {
  wordCreatedExistingRowHeight,
  wordCreatedNewRowHeight,
} from '../../../../store/__tests__/fixtures/word-created-row-height.ts';
import { docx, p, t } from '../word-fixtures.ts';
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  caretIn,
  caretInSurface,
  commandOn,
  comparedBody,
  converged,
  drag,
  firstRowWidths,
  gridOf,
  MERGED_FILE,
  TABLE_FILE,
  tableCommand,
  tableIdsIn,
  type Surface,
} from './tables-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** The revision cards an editor shows. */
const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

const kinds = (room: OrderedLocalRoom): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

async function seats(file = TABLE_FILE, mode: 'edit' | 'suggesting' = 'edit') {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode,
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

async function decideAll(person: Person, action: 'accept' | 'reject'): Promise<void> {
  expect(
    person.editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' }).ok
  ).toBe(true);
  await settle();
}

describe('columns in an ordered room (task 9.7)', () => {
  test("a party's new column: the neighbour lends half its width; a reject gives it back", async () => {
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Rate');
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert-column']);
    converged(room, ana, olive);
    expect(editorXml(olive)).toContain('<w:cellIns ');
    expect(gridOf(olive)).toEqual([3000, 1500, 1500, 3000]);
    expect(firstRowWidths(olive)).toEqual([3000, 1500, 1500, 3000]);
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([3000, 3000, 3000]);
    expect(firstRowWidths(ana)).toEqual([3000, 3000, 3000]);
    expect(editorXml(ana)).not.toContain('cellIns');
  });

  test("a party's new column, accepted: the widths stay as lent", async () => {
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Fee');
    expect(tableCommand(ana, { type: 'insertColumn', where: 'left' }).ok).toBe(true);
    await settle();
    await decideAll(olive, 'accept');
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([1500, 1500, 3000, 3000]);
    expect(editorXml(olive)).not.toContain('cellIns');
  });

  test("a party's deleted column, accepted: the neighbour takes its width", async () => {
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Rate');
    expect(tableCommand(ana, { type: 'deleteColumn' }).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['delete-column']);
    converged(room, ana, olive);
    expect(editorXml(ana)).toContain('<w:cellDel ');
    expect(gridOf(ana)).toEqual([3000, 3000, 3000]);
    await decideAll(olive, 'accept');
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([6000, 3000]);
    expect(editorTexts(ana)).not.toContain('Rate');
  });

  test("the organiser accepts a party's deleted column beside that party's pending fill", async () => {
    // Word accepts the cell deletion alone: the fill stays pending, with the width from before
    // in its record (`/tmp/word-evidence-tables/~$word-fill-coldel-accept.docx`)
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Fee');
    expect(
      tableCommand(ana, { type: 'setCellFill', color: { kind: 'hex', value: 'FFFF00' } }).ok
    ).toBe(true);
    await settle();
    caretIn(ana, 'Rate');
    expect(tableCommand(ana, { type: 'deleteColumn' }).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const column = cards(olive).find((key) => key.startsWith('revision-structural-'))!;
    expect(olive.editor.acceptReviewItem(column)).toEqual({ ok: true, changed: true });
    await settle();
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([6000, 3000]);
    expect(editorTexts(ana)).not.toContain('Rate');
    expect(cards(olive)).toHaveLength(1);
    const xml = editorXml(olive);
    expect(xml).toContain('w:fill="FFFF00"');
    expect(xml.match(/<w:tcPrChange /g)).toHaveLength(2);
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(editorXml(olive)).not.toContain('w:fill="FFFF00"');
  });

  test('the decisions on the saved file give what they give live', async () => {
    for (const action of ['accept', 'reject'] as const) {
      const live = await seats();
      caretIn(live.ana, 'Rate');
      tableCommand(live.ana, { type: 'insertColumn', where: 'right' });
      await settle();
      caretIn(live.ana, 'Total');
      tableCommand(live.ana, { type: 'deleteColumn' });
      await settle();
      const saved = new Uint8Array(await live.ana.editor.save());
      expect(readOoxmlPackage(saved).ok).toBe(true);
      await decideAll(live.olive, action);
      converged(live.room, live.ana, live.olive);
      const expected = { grid: gridOf(live.olive), texts: editorTexts(live.olive) };
      const reopened = await seats(saved);
      expect(gridOf(reopened.olive)).toEqual([3000, 1500, 1500, 3000]);
      await decideAll(reopened.olive, action);
      converged(reopened.room, reopened.ana, reopened.olive);
      expect({ grid: gridOf(reopened.olive), texts: editorTexts(reopened.olive) }).toEqual(
        expected
      );
      closeAll();
    }
  });

  test('two parties add columns at once; each decision gives back its own width', async () => {
    const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Ben'), ...PARTY }));
    const ben = await join(room, 'Ben', { wordFile: TABLE_FILE, transport: held, review });
    const olive = await join(room, 'Olive', {
      wordFile: TABLE_FILE,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    held.held = true;
    caretIn(ana, 'Fee');
    tableCommand(ana, { type: 'insertColumn', where: 'right' });
    caretIn(ben, 'Total');
    tableCommand(ben, { type: 'insertColumn', where: 'right' });
    await settle();
    held.held = false;
    held.release();
    await settle();
    converged(room, ana, ben, olive);
    expect(gridOf(olive)).toEqual([1500, 1500, 3000, 1500, 1500]);
    // The organiser rejects Ana's column first: Ben's stays, with his lent width
    const [anaCard] = cards(olive);
    if (anaCard === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(anaCard).ok).toBe(true);
    await settle();
    converged(room, ana, ben, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 1500, 1500]);
    await decideAll(olive, 'reject');
    converged(room, ana, ben, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000]);
  });

  // A reject of a party's column gives its new cell's width to the cell before it, as the
  // engine's reject does, from the table before the decision: the cell the column took width
  // from at the width it lent. The organiser's suggested column took half of the party's cell;
  // a direct one took its own width
  const besideNewCell: ['suggesting' | 'edit', number[]][] = [
    ['suggesting', [2250, 750, 3000, 3000]],
    ['edit', [3000, 1500, 3000, 3000]],
  ];
  for (const [mode, widths] of besideNewCell) {
    test(`a party's column, then the organiser's ${mode} column beside its new cell`, async () => {
      const { room, ana, olive } = await seats(TABLE_FILE, mode);
      caretIn(ana, 'Fee');
      expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
      await settle();
      caretIn(olive, '');
      expect(tableCommand(olive, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
      await settle();
      const [anaCard] = cards(olive);
      expect(olive.editor.rejectReviewItem(anaCard!).ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect({ grid: gridOf(olive), row: firstRowWidths(olive) }).toEqual({
        grid: widths,
        row: widths,
      });
    });
  }

  // The organiser deletes the column a party's new column took width from: the party's cell
  // gives its width to the cell before it
  test("a party's column, then the organiser deletes the column it took width from", async () => {
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Rate');
    expect(tableCommand(ana, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    caretIn(olive, 'Rate');
    expect(tableCommand(olive, { type: 'deleteColumn' }).ok).toBe(true);
    await settle();
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect({ grid: gridOf(olive), row: firstRowWidths(olive) }).toEqual({
      grid: [4500, 3000],
      row: [4500, 3000],
    });
  });

  test('a party deletes a column while another types in it', async () => {
    const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Ben'), ...PARTY }));
    const ben = await join(room, 'Ben', { wordFile: TABLE_FILE, transport: held, review });
    const olive = await join(room, 'Olive', {
      wordFile: TABLE_FILE,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    held.held = true;
    caretIn(ana, 'Rate');
    tableCommand(ana, { type: 'deleteColumn' });
    caretIn(ben, 'Rate', 4);
    surface(ben).insertPlainText(' per year');
    await settle();
    held.held = false;
    held.release();
    await settle();
    converged(room, ana, ben, olive);
    expect(editorTexts(olive)).toContain('Rate per year');
    await decideAll(olive, 'accept');
    converged(room, ana, ben, olive);
    expect(editorTexts(olive).some((text) => text.startsWith('Rate'))).toBe(false);
    expect(gridOf(olive)).toEqual([6000, 3000]);
  });

  test('a party takes back their own column: the widths come back', async () => {
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Rate');
    tableCommand(ana, { type: 'insertColumn', where: 'left' });
    await settle();
    expect(gridOf(ana)).toEqual([1500, 1500, 3000, 3000]);
    surface(ana).undo();
    await settle();
    expect(kinds(room)).toEqual(['insert-column', 'retract']);
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([3000, 3000, 3000]);
    surface(ana).redo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([1500, 1500, 3000, 3000]);
  });

  test('a party deletes their own new column: it is withdrawn, with no card live or reopened', async () => {
    // The engine withdraws the author's own pending cells and gives the lent widths back; the
    // room showed a card for the column's width records until a save and reopen
    const { room, ana, olive } = await seats();
    const original = editorXml(olive);
    caretIn(ana, 'Rate');
    tableCommand(ana, { type: 'insertColumn', where: 'right' });
    await settle();
    caretAt(ana, editorTexts(ana).indexOf('Rate') + 1, 0);
    expect(tableCommand(ana, { type: 'deleteColumn' }).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000]);
    expect(orderedCards(room.room.document)).toEqual([]);
    expect(cards(olive)).toEqual([]);
    expect(editorXml(olive)).toBe(original);
    const reopened = await seats(new Uint8Array(await ana.editor.save()));
    expect(orderedCards(reopened.room.room.document)).toEqual([]);
    expect(cards(reopened.olive)).toEqual([]);
  });

  test("the organiser's direct column: as the editor alone does it, undone and made again", async () => {
    const { room, ana, olive } = await seats();
    caretIn(olive, 'Rate');
    expect(tableCommand(olive, { type: 'insertColumn', where: 'right' }).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    // Tracking off: the new column takes its reference column's width, and the table grows
    expect(gridOf(olive)).toEqual([3000, 3000, 3000, 3000]);
    expect(editorXml(olive)).not.toContain('cellIns');
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000]);
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000, 3000]);
    caretIn(olive, 'Fee');
    expect(tableCommand(olive, { type: 'deleteColumn' }).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    // The column goes with its width: the table shrinks
    expect(gridOf(olive)).toEqual([3000, 3000, 3000]);
    expect(editorTexts(olive)).not.toContain('Fee');
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000, 3000]);
    expect(editorTexts(olive)).toContain('Fee');
  });
});

describe('column widths and the right edge in an ordered room (task 9.7)', () => {
  test("a party drags a column edge: Word's width records; a reject restores them", async () => {
    const { room, ana, olive } = await seats();
    expect(drag(ana, 0, [4000, 2000]).ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['table-property']);
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([4000, 2000, 3000]);
    expect(editorXml(olive)).toContain('<w:tblGridChange ');
    expect(editorXml(olive)).toContain('<w:tcPrChange ');
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 3000]);
    expect(editorXml(olive)).not.toContain('Change ');
  });

  test('a drag decided on the saved file gives what it gives live', async () => {
    for (const action of ['accept', 'reject'] as const) {
      const live = await seats();
      drag(live.ana, 1, [4000, 2000]);
      await settle();
      const saved = new Uint8Array(await live.ana.editor.save());
      expect(readOoxmlPackage(saved).ok).toBe(true);
      await decideAll(live.olive, action);
      converged(live.room, live.ana, live.olive);
      const expected = gridOf(live.olive);
      // The grid record (`CT_Markup`) is one card with the cell records of its drag
      const reopened = await seats(saved);
      expect(cards(reopened.olive)).toHaveLength(1);
      await decideAll(reopened.olive, action);
      converged(reopened.room, reopened.ana, reopened.olive);
      expect(gridOf(reopened.olive)).toEqual(expected);
      expect(editorXml(reopened.olive)).not.toContain('Change ');
      closeAll();
    }
  });

  test("a party drags the table's right edge; accepted, the widths stay", async () => {
    const { room, ana, olive } = await seats();
    expect(drag(ana, 2, [4500, 10500]).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 4500]);
    expect(editorXml(olive)).toContain('<w:tblW w:type="dxa" w:w="10500"/>');
    await decideAll(olive, 'accept');
    converged(room, ana, olive);
    expect(gridOf(olive)).toEqual([3000, 3000, 4500]);
    expect(editorXml(olive)).not.toContain('Change ');
  });

  test("the organiser's direct drag: no records; undone and made again", async () => {
    const { room, ana, olive } = await seats();
    expect(drag(olive, 1, [1000, 5000]).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([3000, 1000, 5000]);
    expect(editorXml(ana)).not.toContain('Change ');
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([3000, 3000, 3000]);
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(gridOf(ana)).toEqual([3000, 1000, 5000]);
  });

  test("the organiser's drag over a party's pending cell fill shows at once, decided as alone", async () => {
    // Word keeps one record, the cell's properties before the fill, so a reject of the fill
    // takes the organiser's width of that cell away too: in the editor alone and in the room
    const dragFirstEdge = (s: Surface): void => {
      const ids = tableIdsIn(s.session.part().root);
      const target = {
        sourceRevision: s.session.packageRevision(),
        tableId: ids.table,
        leftGridColumnId: ids.grid[0],
        rightGridColumnId: ids.grid[1],
        isHeaderRepeat: false,
      };
      const command = {
        type: 'commitTableColumnDividerResize',
        target,
        leftWidthTwips: 4000,
        rightWidthTwips: 2000,
      };
      expect(commandOn(s, command).ok).toBe(true);
    };
    for (const action of ['accept', 'reject'] as const) {
      const { room, ana, olive } = await seats();
      caretIn(ana, 'Fee');
      const fill = { type: 'setCellFill', color: { kind: 'hex', value: 'FFFF00' } };
      expect(tableCommand(ana, fill).ok).toBe(true);
      await settle();
      const filled = new Uint8Array(await ana.editor.save());
      dragFirstEdge(surface(olive));
      await settle();
      converged(room, ana, olive);
      expect(firstRowWidths(ana)).toEqual([4000, 2000, 3000]);
      expect(editorXml(ana)).toContain('w:fill="FFFF00"');
      await decideAll(olive, action);
      converged(room, ana, olive);
      expect(comparedBody(surface(olive).session.part())).toBe(
        alone(
          { author: 'Olive', mode: 'edit', review, file: filled, decide: action },
          dragFirstEdge
        )
      );
      closeAll();
    }
  });

  test("a second party's drag on rows that hold another party's pending widths is refused", async () => {
    // The engine refuses a change to a row that holds another author's pending property
    // change; the host's earlier check refuses the same edit for a party (`others_suggestions`)
    const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Ben'), ...PARTY }));
    const ben = await join(room, 'Ben', { wordFile: TABLE_FILE, transport: held, review });
    held.held = true;
    drag(ana, 0, [4000, 2000]);
    drag(ben, 0, [2500, 3500]);
    await settle();
    held.held = false;
    held.release();
    await settle();
    // Found when the copy applies it again, and refused by the room too: one report for the drag
    expect(ben.refusals).toEqual([
      expect.objectContaining({ reason: 'pending-refused', detail: 'engine-refused' }),
    ]);
    ben.refusals.length = 0;
    converged(room, ana, ben);
    expect(gridOf(ana)).toEqual([4000, 2000, 3000]);
    expect(cards(ana)).toHaveLength(1);
  });

  test('a drag on an edge another person moved by adding a column is refused by name', async () => {
    const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Ben'), ...PARTY }));
    const ben = await join(room, 'Ben', { wordFile: TABLE_FILE, transport: held, review });
    held.held = true;
    caretIn(ana, 'Fee');
    tableCommand(ana, { type: 'insertColumn', where: 'right' });
    drag(ben, 0, [4000, 2000]);
    await settle();
    held.held = false;
    held.release();
    await settle();
    // Ben's edge is now between Fee and Ana's column, whose widths do not add up to his
    expect(ben.refusals).toEqual([
      expect.objectContaining({ reason: 'pending-refused', detail: 'engine-refused' }),
    ]);
    ben.refusals.length = 0;
    converged(room, ana, ben);
    expect(gridOf(ben)).toEqual([1500, 1500, 3000, 3000]);
  });
});

/**
 * A party's steps in the room give what the editor alone gives, pending and decided, and
 * decided on the saved file. The room settles between steps, as the publish does between a
 * person's key presses.
 */
async function asAlone(
  steps: readonly ((s: Surface) => void)[],
  file: Uint8Array = TABLE_FILE
): Promise<void> {
  for (const decide of [undefined, 'accept', 'reject'] as const) {
    const { room, ana, olive } = await seats(file);
    for (const step of steps) {
      step(surface(ana));
      await settle();
    }
    converged(room, ana, olive);
    const saved = decide === undefined ? null : new Uint8Array(await ana.editor.save());
    if (decide !== undefined) await decideAll(olive, decide);
    converged(room, ana, olive);
    const expected = alone({ author: 'Ana', mode: 'suggesting', review, decide, file }, (s) => {
      for (const step of steps) step(s);
    });
    expect({ decide, body: comparedBody(surface(olive).session.part()) }).toEqual({
      decide,
      body: expected,
    });
    if (saved && decide !== undefined) {
      // The same decision on the saved file
      const reopened = await seats(saved);
      await decideAll(reopened.olive, decide);
      converged(reopened.room, reopened.ana, reopened.olive);
      expect({ decide, reopened: comparedBody(surface(reopened.olive).session.part()) }).toEqual({
        decide,
        reopened: expected,
      });
    }
    closeAll();
  }
}

describe('rows in an ordered room (task 9.7)', () => {
  test("a party's new row is Word's form; another person's typing in it joins its card", async () => {
    // As Word writes a row it inserts with Track Changes on (a file Word saved:
    // `e2e/fixtures/list-pagination-break.docx`): the row and each cell's paragraph mark
    const { room, ana } = await seats();
    const olive = await join(room, 'Olive', { wordFile: TABLE_FILE, login: ORGANISER, review });
    caretIn(ana, 'Fee');
    expect(tableCommand(ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const xml = editorXml(olive);
    expect(xml.match(/<w:trPr><w:ins /g)).toHaveLength(1);
    expect(xml.match(/<w:pPr><w:rPr><w:ins /g)).toHaveLength(3);
    expect(xml).not.toContain('<w:cellIns ');
    // The organiser suggests text in it: the host's check reads it inside the row, as the engine
    // does
    caretAt(olive, editorTexts(olive).indexOf('Total') + 1, 0);
    surface(olive).insertPlainText('x');
    await settle();
    converged(room, ana, olive);
    expect(orderedCards(room.room.document).map((card) => card.suggestions.length)).toEqual([2]);
    expect(cards(olive)).toHaveLength(1);
    const saved = new Uint8Array(await olive.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    await decideAll(olive, 'reject');
    converged(room, ana, olive);
    expect(editorTexts(olive)).not.toContain('x');
    expect(editorXml(olive)).not.toContain('<w:ins ');
    // The saved file reopens to the same card, decided the same way
    const reopened = await seats(saved);
    expect(orderedCards(reopened.room.room.document)).toHaveLength(1);
    await decideAll(reopened.olive, 'reject');
    converged(reopened.room, reopened.ana, reopened.olive);
    expect(editorTexts(reopened.olive)).toEqual(editorTexts(olive));
  });
});

describe('a row height Word saved, in an ordered room (task 9.7)', () => {
  for (const [name, fixture] of [
    ['with no row properties', wordCreatedNewRowHeight],
    ['over a height', wordCreatedExistingRowHeight],
  ] as const)
    test(`${name}: one card, decided from the editor as the editor alone decides it`, async () => {
      // Word records it on the table and on each cell of the row, each record with its own id
      const file = docx(fixture.replace(/ \/>/g, '/>') + p(t('After.')));
      for (const action of ['accept', 'reject'] as const) {
        const { room, ana, olive } = await seats(file);
        expect(cards(olive)).toHaveLength(1);
        expect(orderedCards(room.room.document)).toHaveLength(1);
        await decideAll(olive, action);
        converged(room, ana, olive);
        const expected = alone(
          { author: 'Olive', mode: 'edit', review, file, decide: action },
          () => {}
        );
        expect(comparedBody(surface(olive).session.part())).toBe(expected);
        expect(cards(olive)).toEqual([]);
        closeAll();
      }
    });
});

describe('new tables in an ordered room (task 9.7)', () => {
  test("a party's new table right after a table: the paragraph the engine adds between", async () => {
    await asAlone([
      (s) => {
        caretInSurface(s, 'Signed here.');
        expect(s.insertTable(1, 2)).toBe(true);
      },
    ]);
  });

  test("a party's new table is Word's form: inserted rows and cell marks, no cell markers", async () => {
    // As Word writes a table it inserts with Track Changes on (a file Word saved:
    // `e2e/fixtures/list-pagination-break.docx`)
    const { room, ana, olive } = await seats();
    caretIn(ana, 'Signed here.');
    expect(surface(ana).insertTable(2, 2)).toBe(true);
    await settle();
    converged(room, ana, olive);
    const table =
      editorXml(olive).match(/<w:tbl>(?:(?!<w:tbl>).)*?<w:trPr><w:ins .*?<\/w:tbl>/)?.[0] ?? '';
    expect(table.match(/<w:trPr><w:ins /g)).toHaveLength(2);
    expect(table.match(/<w:pPr><w:rPr><w:ins /g)).toHaveLength(4);
    expect(table).not.toContain('<w:cellIns ');
  });

  test("a party's new table in the empty paragraph a pending Enter made", async () => {
    await asAlone([
      (s) => {
        caretInSurface(s, 'Signed here.');
        s.splitParagraph();
      },
      (s) => {
        caretInSurface(s, '');
        expect(s.insertTable(2, 2)).toBe(true);
      },
    ]);
  });

  test('two new tables before one paragraph: each new paragraph has its own w14:paraId, kept through a save and reopen', async () => {
    // The engine builds a new table on a part that holds only the paragraph it goes before, and
    // seeds the cells' ids from that paragraph: with another paragraph between the two tables,
    // the second table took the first one's ids (acceptance s10 seed 1002)
    const file = docx(p(t('One.')) + p(t('Two.')));
    const { room, ana, olive } = await seats(file);
    caretIn(ana, 'Two.');
    expect(surface(ana).insertTable(2, 2)).toBe(true);
    await settle();
    caretIn(olive, 'Two.');
    surface(olive).splitParagraph();
    await settle();
    caretIn(olive, 'Two.');
    expect(surface(olive).insertTable(2, 2)).toBe(true);
    await settle();
    converged(room, ana, olive);
    const base = readOoxmlPackage(file);
    if (!base.ok) throw new Error('no base');
    const bytes = orderedPackageBytes(room.room.document, base.package);
    const saved = readOoxmlPackage(bytes);
    if (!saved.ok) throw new Error('the saved file does not open');
    const main = saved.package.parts.get(BODY);
    if (!main) throw new Error('no main part');
    const ids = [...serializeOoxmlPart(main).matchAll(/<w:p [^>]*w14:paraId="([0-9A-F]{8})"/g)];
    expect(ids.length).toBeGreaterThan(8);
    expect(new Set(ids.map((m) => m[1])).size).toBe(ids.length);
    const reopened = openOrderedWordFile(bytes, { file: 2 }).document;
    expect(serializeOoxmlPart(composeOrderedPart(reopened, BODY).part)).toBe(
      serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part)
    );
  });

  test("the organiser's new table right after a table, undone", async () => {
    const { room, ana, olive } = await seats();
    const before = editorXml(olive);
    caretIn(olive, 'Signed here.');
    expect(surface(olive).insertTable(1, 1)).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(comparedBody(surface(olive).session.part())).toBe(
      alone({ author: 'Olive', mode: 'edit', review }, (s) => {
        caretInSurface(s, 'Signed here.');
        s.insertTable(1, 1);
      })
    );
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(editorXml(olive)).toBe(before);
  });
});

describe('columns in a table with merged cells, as Word makes them (task 9.7)', () => {
  /** Put the caret in a cell and run one of the editor's column commands, in a room or alone. */
  const column =
    (text: string, command: Readonly<Record<string, unknown>>) =>
    (s: Surface): void => {
      caretInSurface(s, text);
      expect(commandOn(s, command).ok).toBe(true);
    };

  /** The organiser's direct command gives what the editor alone gives, and undo takes it back. */
  async function directAsAlone(step: (s: Surface) => void): Promise<void> {
    const { room, ana, olive } = await seats(MERGED_FILE);
    const before = editorXml(olive);
    step(surface(olive));
    await settle();
    converged(room, ana, olive);
    expect(comparedBody(surface(olive).session.part())).toBe(
      alone({ author: 'Olive', mode: 'edit', review, file: MERGED_FILE }, step)
    );
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(editorXml(olive)).toBe(before);
  }

  test("the organiser inserts a column beside the merged cells: Word's cells, undone", async () => {
    await directAsAlone(column('5', { type: 'insertColumn', where: 'right' }));
  });

  test('the organiser deletes the column under the spanning cell: it goes whole, undone', async () => {
    await directAsAlone(column('5', { type: 'deleteColumn' }));
  });

  test("a party's column beside the merged cells: pending, decided, and decided after a reopen", async () => {
    await asAlone([column('5', { type: 'insertColumn', where: 'right' })], MERGED_FILE);
  });

  test("a party's deletion of the column of the vertical merge", async () => {
    await asAlone([column('105', { type: 'deleteColumn' })], MERGED_FILE);
  });

  test("a party's drag of the edge beside the vertical merge", async () => {
    await asAlone(
      [
        (s) => {
          const ids = tableIdsIn(s.session.part().root);
          const sourceRevision = s.session.packageRevision();
          const target = {
            sourceRevision,
            tableId: ids.table,
            leftGridColumnId: ids.grid[1],
            rightGridColumnId: ids.grid[2],
            isHeaderRepeat: false,
          };
          const command = {
            type: 'commitTableColumnDividerResize',
            target,
            leftWidthTwips: 4000,
            rightWidthTwips: 2240,
          };
          expect(commandOn(s, command).ok).toBe(true);
        },
      ],
      MERGED_FILE
    );
  });
});
