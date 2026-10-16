/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Two participants suggesting table formatting on one row at the same time.
//
// Word answers a row's property records as one decision, so suggesting refuses a change to a
// row that holds another author's pending change. Two people can still change one row before
// either sees the other's record. The row then belongs to the change with the earlier
// `w:date`, then the lesser author, then the lesser revision id, on every replica:
//
// - the author of the later change takes it back when the earlier one arrives, while it is
//   still their last step, and reads why;
// - after a later local edit, the row shows as its owner's one decision, records last in their
//   containers, and no order of accepting and rejecting loses a cell width;
// - offline, a table formatting suggestion waits for the room, because it would arrive too late
//   to be taken back.

import { afterEach, expect, setSystemTime, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { ROW_TAKEN_REASON } from '../../../../core/src/editor/surface-table-row-race.ts';
import {
  HELD_ROW_REASON,
  OFFLINE_TABLE_REASON,
} from '../../../../core/src/editor/surface-tracked-table-refusal.ts';
import { planTableRowHeightResize } from '../../../../core/src/editor/table-command-plan.ts';
import {
  findTableOccurrence,
  tableColumnDividerResizeTargetFrom,
  tableRowOccurrenceTargetFrom,
  type TableOccurrenceRef,
} from '../../../../core/src/layout/table-interaction-targets.ts';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { mainPart, packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('table-row-races', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
type ExecResult = ReturnType<Editor['exec']>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
/** Close every editor and replica, so the next race meets in a room of its own. */
function closeAll(): void {
  setSystemTime();
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
}
afterEach(closeAll);

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
// Paragraphs in order: before, a1, a2, a3, b1, b2, b3, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}${cell('a3')}</w:tr>` +
  `<w:tr>${cell('b1')}${cell('b2')}${cell('b3')}</w:tr></w:tbl>` +
  p('after');
const A1 = 1;
const A3 = 3;
const B3 = 6;
const AFTER = 7;

const LANDED: ExecResult = { ok: true, changed: true };
const OFFLINE: ExecResult = { ok: false, code: 'unsupported', reason: OFFLINE_TABLE_REASON };
const EARLY = '2026-10-16T09:00:00Z';
const LATE = '2026-10-16T09:01:00Z';

function mount(document: Uint8Array, author: string, peer?: Peer): Editor {
  peer?.detach();
  const container = window.document.createElement('div');
  window.document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    modules: [
      reviewModule(),
      ...(peer ? [collaborationModule({ session: peer.room.session })] : []),
    ],
  });
  editors.push(editor);
  return editor;
}

function caret(editor: Editor, paragraph: number, offset = 0): void {
  const position = { paragraphId: editor.surface!.session.paragraphIds()[paragraph]!, offset };
  editor.surface!.setSelection({ anchor: position, head: position });
}

function suggest(editor: Editor): void {
  expect(editor.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
}

const xmlOf = (editor: Editor) =>
  serializeOoxmlPart(mainPart(editor.surface!.session.currentPackage()));
/** The authors of the formatting cards, in document order. */
const authorsOf = (editor: Editor) =>
  editor
    .getReviewItems()
    .flatMap((item) =>
      item.kind === 'revision' && item.revisionKind === 'format' ? [item.author] : []
    );
/** Whether the `!` typed at the end of the last paragraph shows. */
const typed = (editor: Editor) => /<w:t[^>]*>!<\/w:t>/.test(xmlOf(editor));
const rejectionOf = (editor: Editor) => editor.snapshot().lastRejection ?? null;

/** One row's occurrence of the table, as a pointer gesture targets it. */
function rowOf(editor: Editor, index: number): TableOccurrenceRef {
  const layout = editor.surface!.layout();
  const table = layout.pages[0]!.fragments.find((fragment) => fragment.kind === 'table');
  if (table?.kind !== 'table') throw new Error('no table on the first page');
  const ref = findTableOccurrence(layout, table.tableId, table.rows[index]!.id, false);
  if (!ref) throw new Error('row occurrence missing');
  return ref;
}

/** A table formatting gesture, and the markup its change leaves in the document. */
interface Gesture {
  readonly name: string;
  readonly run: (editor: Editor) => ExecResult;
  readonly mark: string;
}

const fill = (paragraph: number, color: string): Gesture => ({
  name: `fill of ${['', 'a1', 'a2', 'a3'][paragraph]}`,
  run: (editor) => {
    caret(editor, paragraph);
    return editor.exec({ type: 'setCellFill', color: { kind: 'hex', value: color } });
  },
  mark: `w:fill="${color}"`,
});

const ROW_HEIGHT: Gesture = {
  name: 'height of the first row',
  run: (editor) => {
    const surface = editor.surface!;
    const layout = surface.layout();
    const state = surface.state();
    const plan = planTableRowHeightResize(
      {
        part: surface.session.part(),
        layout,
        storeRevision: surface.session.packageRevision(),
        selection: state.selection,
        cellSelection: state.cellSelection,
        themeColors: surface.session.documentThemeColors(),
        editable: true,
        viewing: false,
      },
      tableRowOccurrenceTargetFrom(layout.revision, rowOf(editor, 0)),
      900
    );
    return surface.applyTableCommandPlan(plan);
  },
  mark: 'w:val="900"',
};

const RESIZE: Gesture = {
  name: 'column resize',
  run: (editor) => {
    const ref = rowOf(editor, 0);
    const [left, right] = ref.row.cells.map((entry) => entry.gridColumnId!);
    const target = tableColumnDividerResizeTargetFrom(
      editor.surface!.layout().revision,
      ref,
      left!,
      right!
    );
    return editor.exec({
      type: 'commitTableColumnDividerResize',
      target,
      leftWidthTwips: 4000,
      rightWidthTwips: 2000,
    });
  },
  mark: 'w:w="4000"',
};

async function peers() {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice.room.document, 'Alice', pair.alice);
  const bob = mount(pair.bob.room.document, 'Bob', pair.bob);
  suggest(alice);
  suggest(bob);
  const flush = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
  };
  const converge = () => {
    flush();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
    expect(saveReopenDigest(alice.surface!.session.currentPackage())).toEqual(
      saveReopenDigest(bob.surface!.session.currentPackage())
    );
  };
  return { pair, alice, bob, flush, converge };
}

/** Run `gesture` on `editor` as if at `time`. */
function at(time: string, editor: Editor, gesture: Gesture): void {
  setSystemTime(new Date(time));
  try {
    expect(gesture.run(editor)).toEqual(LANDED);
  } finally {
    setSystemTime();
  }
}

/** Let the editors act on the updates they just received. */
const settled = () => new Promise<void>((resolve) => queueMicrotask(resolve));

/**
 * Alice's and Bob's gestures, made at once, with `earlier`'s first by the clock. With
 * `editAfter`, the later author also types before the two meet.
 */
async function race(alices: Gesture, bobs: Gesture, earlier: 'Alice' | 'Bob', editAfter = false) {
  const session = await peers();
  const { pair, alice, bob, flush } = session;
  const later = earlier === 'Alice' ? bob : alice;
  pair.pause();
  at(earlier === 'Alice' ? EARLY : LATE, alice, alices);
  at(earlier === 'Alice' ? LATE : EARLY, bob, bobs);
  if (editAfter) {
    caret(later, AFTER, 5);
    later.surface!.type('!');
  }
  flush();
  pair.resume();
  await settled();
  session.converge();
  const [winner, loser] = earlier === 'Alice' ? [alice, bob] : [bob, alice];
  const [kept, lost] = earlier === 'Alice' ? [alices, bobs] : [bobs, alices];
  return { ...session, winner, loser, kept, lost };
}

const RACES: readonly (readonly [string, Gesture, Gesture])[] = [
  ['cell against cell', fill(A1, '0000FF'), fill(A3, 'FF0000')],
  ['cell against row height', fill(A1, '0000FF'), ROW_HEIGHT],
  ['cell against column resize', fill(A3, '0000FF'), RESIZE],
];

for (const [name, alices, bobs] of RACES) {
  for (const earlier of ['Alice', 'Bob'] as const) {
    test(`${name}: ${earlier}'s earlier change keeps the row, the later one is taken back`, async () => {
      const { alice, bob, winner, loser, kept, lost, converge } = await race(alices, bobs, earlier);
      for (const editor of [alice, bob]) {
        expect(authorsOf(editor)).toEqual([earlier]);
        expect(xmlOf(editor)).toContain(kept.mark);
        expect(xmlOf(editor)).not.toContain(lost.mark);
      }
      expect(rejectionOf(loser)).toBe(ROW_TAKEN_REASON);
      expect(rejectionOf(winner)).toBeNull();

      const reopened = mount(new Uint8Array(await winner.save()), 'Reader');
      expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
        saveReopenDigest(loser.surface!.session.currentPackage())
      );
      expect(authorsOf(reopened)).toEqual([earlier]);

      // Neither Redo nor Undo brings the change that was taken back again.
      expect(loser.exec({ type: 'redo' }).ok).toBe(false);
      loser.exec({ type: 'undo' });
      converge();
      for (const editor of [alice, bob]) {
        expect(authorsOf(editor)).toEqual([earlier]);
        expect(xmlOf(editor)).not.toContain(lost.mark);
      }
    });
  }
}

test('changes to different rows at once keep two cards, and neither is taken back', async () => {
  const { alice, bob } = await race(fill(A1, '0000FF'), fill(B3, 'FF0000'), 'Bob');
  for (const editor of [alice, bob]) {
    expect(authorsOf(editor)).toEqual(['Alice', 'Bob']);
    expect(xmlOf(editor)).toContain('w:fill="0000FF"');
    expect(xmlOf(editor)).toContain('w:fill="FF0000"');
    expect(rejectionOf(editor)).toBeNull();
  }
});

test('a change to the row of a pending change that arrived first is still refused', async () => {
  const { alice, bob, converge } = await peers();
  expect(fill(A1, '0000FF').run(alice)).toEqual(LANDED);
  converge();
  expect(fill(A3, 'FF0000').run(bob)).toEqual({
    ok: false,
    code: 'unsupported',
    reason: HELD_ROW_REASON,
  });
  expect(authorsOf(bob)).toEqual(['Alice']);
});

// AFTER A LATER LOCAL EDIT the losing change cannot be taken back alone. The row shows as one
// decision, so every order of answering it keeps each cell's width.

/** `w:tcW` widths that are not twentieths of a point: a reset cell reads `auto`. */
const autoWidths = (editor: Editor) => xmlOf(editor).match(/<w:tcW[^>]*w:type="auto"/g) ?? [];

/** Each row's records, as the distinct attributions they show. */
const rowOwners = (editor: Editor) =>
  (xmlOf(editor).match(/<w:tr>.*?<\/w:tr>/g) ?? []).map(
    (row) =>
      new Set(
        [...row.matchAll(/<w:(?:tcPrChange|trPrChange|tblPrExChange) ([^>]*)>/g)].map(
          ([, attributes]) => attributes
        )
      ).size
  );

/** Whether each record is the last child of its container. */
const recordsLast = (editor: Editor) =>
  !/<\/w:(tcPrChange|trPrChange|tblPrExChange)>(?!<\/w:(tcPr|trPr|tblPrEx)>)/.test(xmlOf(editor));

type Answer = 'accept' | 'reject';
const ANSWERS: readonly Answer[] = ['accept', 'reject'];

/** The formatting card of `author`. */
const cardOf = (editor: Editor, author: string) =>
  editor
    .getReviewItems()
    .find(
      (item) => item.kind === 'revision' && item.revisionKind === 'format' && item.author === author
    );

const C_RACES: readonly (readonly [string, Gesture, Gesture])[] = [
  // Each cell of the first row held both authors' records, merged cell by cell.
  ['a column resize and a fill', RESIZE, fill(A3, 'FF0000')],
  // One card named Bob and held Alice's row record, and hers had none.
  ['a row height and a fill', ROW_HEIGHT, fill(A3, 'FF0000')],
  // Every cell showed Bob's records, and a1 showed `w:shd` after its `w:tcPrChange`.
  ['two fills', fill(A1, '0000FF'), fill(A3, 'FF0000')],
];

for (const [name, alices, bobs] of C_RACES) {
  for (const earlier of ['Alice', 'Bob'] as const) {
    test(`${name}, with an edit after the later one: one owner per row, no width lost (${earlier} first)`, async () => {
      const { alice, bob, loser } = await race(alices, bobs, earlier, true);
      const cards = authorsOf(alice);
      expect(cards[0]).toBe(earlier);
      expect(cards.length).toBeLessThanOrEqual(2);
      expect(rejectionOf(loser)).toBeNull();
      for (const editor of [alice, bob]) {
        expect(authorsOf(editor)).toEqual(cards);
        expect(recordsLast(editor)).toBe(true);
        for (const owners of rowOwners(editor)) expect(owners).toBeLessThanOrEqual(1);
        expect(typed(editor)).toBe(true);
      }

      // Every order of answering the cards, each either way, each on replicas of its own.
      const orders = cards.length === 1 ? [cards] : [cards, [...cards].reverse()];
      const answers =
        cards.length === 1
          ? ANSWERS.map((answer) => [answer])
          : ANSWERS.flatMap((first) => ANSWERS.map((second) => [first, second]));
      for (const order of orders) {
        for (const chosen of answers) {
          closeAll();
          const run = await race(alices, bobs, earlier, true);
          for (const [index, author] of order.entries()) {
            const card = cardOf(run.alice, author);
            expect(card).toBeDefined();
            const answered =
              chosen[index] === 'accept'
                ? run.bob.acceptReviewItem(card!.key)
                : run.bob.rejectReviewItem(card!.key);
            expect(answered.ok).toBe(true);
            run.converge();
            expect(autoWidths(run.alice)).toEqual([]);
          }
          expect(authorsOf(run.alice)).toEqual([]);
          expect(xmlOf(run.alice)).not.toMatch(/<w:(tcPrChange|trPrChange|tblPrExChange)\b/);
        }
      }
    });
  }
}

test('offline, a table formatting suggestion waits for the room, and text does not', async () => {
  const { pair, alice, bob, converge } = await peers();
  const original = xmlOf(alice);
  pair.pause();
  pair.alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test');
  caret(alice, A1);
  for (const command of [
    { type: 'setCellFill', color: { kind: 'hex', value: '0000FF' } },
    { type: 'setTableCellVerticalAlignment', alignment: 'center' },
    { type: 'toggleHeaderRow' },
    { type: 'setTableProperties', justification: 'center' },
  ] as const) {
    expect(alice.exec(command)).toEqual(OFFLINE);
  }
  expect(ROW_HEIGHT.run(alice)).toEqual(OFFLINE);
  expect(RESIZE.run(alice)).toEqual(OFFLINE);
  expect(xmlOf(alice)).toBe(original);
  caret(alice, AFTER, 5);
  alice.surface!.type('!');

  pair.alice.room.session.setTransportStatus('ready');
  pair.resume();
  converge();
  expect(typed(bob)).toBe(true);
  expect(fill(A1, '0000FF').run(alice)).toEqual(LANDED);
  converge();
  expect(authorsOf(bob)).toEqual(['Alice']);
  expect(xmlOf(bob)).toContain('w:fill="0000FF"');
});

test('offline, table formatting outside suggesting still applies', async () => {
  const { pair, alice, bob, converge } = await peers();
  expect(alice.exec({ type: 'setEditingMode', mode: 'editing' }).ok).toBe(true);
  pair.pause();
  pair.alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test');
  expect(fill(A1, '0000FF').run(alice)).toEqual(LANDED);
  pair.alice.room.session.setTransportStatus('ready');
  pair.resume();
  converge();
  expect(xmlOf(bob)).toContain('w:fill="0000FF"');
  expect(authorsOf(bob)).toEqual([]);
});
