// A cell's one change record with real editors and one room (backlog 34). A second person's
// fill that the cell already has changes nothing, also when it was pressed before the first
// person's fill reached them, so the first person's change keeps its author. A person's fill on
// a row where their own height is pending, with no `w:trPr` to hold it, joins the cells' one
// record, so Reject All takes the height too, in the room and in the editor alone. So does one
// fill press over a column of two rows that each hold their own pending height. Each decision
// gives Word's result: a reject gives back each row and cell as it was before that change.

import { afterEach, describe, expect, setSystemTime, test } from 'bun:test';
import { selectCellRectangle } from '../../../../editor/__tests__/paginated-surface-fixtures.ts';
import { planTableRowHeightResize } from '../../../../editor/table-command-plan.ts';
import {
  findTableOccurrence,
  tableRowOccurrenceTargetFrom,
} from '../../../../layout/table-interaction-targets.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
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
import { reviewTestModule } from './review-module.ts';
import {
  alone,
  caretIn,
  caretInSurface,
  commandOn,
  comparedBody,
  converged,
  TABLE_FILE,
  tableCommand,
  tableIdsIn,
  type Surface,
} from './tables-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(() => {
  closeAll();
  setSystemTime();
});

const review = reviewTestModule();
const FILL = { type: 'setCellFill', color: { kind: 'hex', value: 'FFCC00' } };
const RECORDS = /<w:(trPrChange|tcPrChange|tblPrChange|tblGridChange|tblPrExChange)\b/;

/** Ana, Ben and the organiser Olive in one room; Ben's frames from the room wait while `held`. */
async function seats() {
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
  return { room, ana, ben, held, olive };
}

/** The authors of the revision cards an editor shows. */
const authors = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [item.author ?? ''] : []));

/** Drag the divider under table row `index` to `heightTwips`, as the editor's pointer does. */
function dragRow(s: Surface, index: number, heightTwips: number): void {
  const layout = s.layout();
  const { table, rows } = tableIdsIn(s.session.part().root);
  const ref = findTableOccurrence(layout, table, rows[index]!, false);
  if (!ref) throw new Error('no row occurrence');
  const state = s.state();
  const plan = planTableRowHeightResize(
    {
      part: s.session.part(),
      layout,
      storeRevision: s.session.packageRevision(),
      selection: state.selection,
      cellSelection: state.cellSelection,
      themeColors: s.session.documentThemeColors(),
      editable: true,
      viewing: false,
    },
    tableRowOccurrenceTargetFrom(layout.revision, ref),
    heightTwips
  );
  expect(s.applyTableCommandPlan(plan).ok).toBe(true);
}

async function decideAll(person: Person, action: 'accept' | 'reject'): Promise<void> {
  expect(
    person.editor.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' }).ok
  ).toBe(true);
  await settle();
}

/** The editor's table as decided: the fills and the heights it shows, and no record left. */
function decided(person: Person): { fills: number; heights: number } {
  const xml = editorXml(person);
  expect(xml).not.toMatch(RECORDS);
  return {
    fills: xml.split('w:fill="FFCC00"').length - 1,
    heights: xml.split('<w:trHeight').length - 1,
  };
}

/** Ana fills the first cell; Ben gives it the same fill, after it reached him or before. */
async function sameFill(at: 'after' | 'before') {
  const seated = await seats();
  const { ana, ben, held } = seated;
  held.held = at === 'before';
  caretIn(ana, 'Fee');
  expect(tableCommand(ana, FILL).ok).toBe(true);
  await settle();
  caretIn(ben, 'Fee');
  expect(tableCommand(ben, FILL).ok).toBe(true);
  await settle();
  held.held = false;
  held.release();
  await settle();
  return seated;
}

type Step = (s: Surface) => void;
const height =
  (row: number): Step =>
  (s) =>
    dragRow(s, row, 600);
const fill: Step = (s) => expect(commandOn(s, FILL).ok).toBe(true);
/** One press over the first column of both rows. */
const columnFill: Step = (s) => {
  selectCellRectangle(s, { row: 0, column: 0 }, { row: 1, column: 0 });
  expect(commandOn(s, FILL).ok).toBe(true);
};

/**
 * Ana's own heights (the rows have no `w:trPr`) and her fill, at the caret's first cell or over
 * the first column, with the rows each holds.
 */
const OWN_STEPS: Readonly<Record<string, { steps: readonly Step[]; rows: number }>> = {
  'height, then fill': { steps: [height(0), fill], rows: 1 },
  'fill, then height': { steps: [fill, height(0)], rows: 1 },
  'two heights, then a column fill': { steps: [height(0), height(1), columnFill], rows: 2 },
  'a column fill, then two heights': { steps: [columnFill, height(0), height(1)], rows: 2 },
};

/** Run the steps at the first cell, each a minute after the one before. */
function runSteps(s: Surface, steps: readonly Step[], minute: number): void {
  setSystemTime(new Date(Date.UTC(2026, 10, 3, 9, minute)));
  caretInSurface(s, 'Fee');
  for (const step of steps) step(s);
}

describe("a second person's fill that the cell already has (backlog 34)", () => {
  for (const at of ['after', 'before'] as const)
    test(`pressed ${at} the first fill reached them: the first party's change keeps its author`, async () => {
      const { room, ana, ben, olive } = await sameFill(at);
      converged(room, ana, ben, olive);
      expect([...new Set(authors(olive))]).toEqual(['Ana']);
      const [card] = olive.editor.getReviewItems();
      expect(olive.editor.rejectReviewItem(card!.key).ok).toBe(true);
      await settle();
      converged(room, ana, ben, olive);
      expect(decided(olive)).toEqual({ fills: 0, heights: 0 });
    });

  for (const action of ['accept', 'reject'] as const)
    test(`pressed before the first fill reached them: ${action} all gives Word's result`, async () => {
      const { room, ana, ben, olive } = await sameFill('before');
      await decideAll(olive, action);
      converged(room, ana, ben, olive);
      expect(decided(olive)).toEqual({ fills: action === 'accept' ? 1 : 0, heights: 0 });
    });
});

describe("a party's own heights and fill on rows with no properties (backlog 34)", () => {
  for (const [order, { steps, rows }] of Object.entries(OWN_STEPS))
    for (const action of ['accept', 'reject'] as const)
      test(`${order}: ${action} all decides every one, as in the editor alone`, async () => {
        const { room, ana, ben, olive } = await seats();
        for (const [minute, step] of steps.entries()) {
          runSteps(surface(ana), [step], minute);
          await settle();
        }
        converged(room, ana, ben, olive);
        await decideAll(olive, action);
        converged(room, ana, ben, olive);
        const kept = action === 'accept' ? rows : 0;
        expect(decided(olive)).toEqual({ fills: kept, heights: kept });
        const editorAlone = alone(
          { author: 'Ana', mode: 'suggesting', review, decide: action },
          (s) => {
            for (const [minute, step] of steps.entries()) runSteps(s, [step], minute);
          }
        );
        expect(comparedBody(surface(olive).session.part())).toBe(editorAlone);
      });
});
