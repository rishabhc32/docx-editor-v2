/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A proposed row height and a proposed column resize, between two participants.
//
// Both are ordinary tracked markup — `w:trPrChange`, `w:tblGridChange`, `w:tblPrChange` and the
// cells' `w:tcPrChange` — so they travel as any other edit does. Each participant sees the same
// card and can answer it, undo and redo replicate, concurrent edits elsewhere keep theirs, and a
// returning participant and a reopened save read the same proposal. A change to a row that holds
// the other participant's pending change is refused until that change is answered.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { HELD_ROW_REASON } from '../../../../core/src/editor/surface-tracked-table-refusal.ts';
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
const harness = createPeerHarness('tracked-table-properties', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
// Paragraphs in order: before, a1, a2, b1, b2, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}</w:tr><w:tr>${cell('b1')}${cell('b2')}</w:tr></w:tbl>` +
  p('after');
const A1 = 1;
const B2 = 4;
// The same with a third column, whose cell a resize of the first two leaves alone.
const THREE_COLUMNS = BODY.replace('</w:tblGrid>', '<w:gridCol w:w="3000"/></w:tblGrid>')
  .replace('</w:tr><w:tr>', `${cell('a3')}</w:tr><w:tr>`)
  .replace('</w:tr></w:tbl>', `${cell('b3')}</w:tr></w:tbl>`);
const A3 = 3;

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
const cards = (editor: Editor) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.revisionKind : item.kind));
const RECORDS = /<w:(trPrChange|tcPrChange|tblPrChange|tblGridChange)\b/;

/** One row's occurrence of the table, as a pointer gesture targets it. */
function rowOf(editor: Editor, index: number): TableOccurrenceRef {
  const layout = editor.surface!.layout();
  const table = layout.pages[0]!.fragments.find((fragment) => fragment.kind === 'table');
  if (table?.kind !== 'table') throw new Error('no table on the first page');
  const ref = findTableOccurrence(layout, table.tableId, table.rows[index]!.id, false);
  if (!ref) throw new Error('row occurrence missing');
  return ref;
}

type ExecResult = ReturnType<Editor['exec']>;
const LANDED: ExecResult = { ok: true, changed: true };

/** Drag the divider under row `index` to `heightTwips`. */
function dragRowHeight(editor: Editor, index: number, heightTwips: number): ExecResult {
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
    tableRowOccurrenceTargetFrom(layout.revision, rowOf(editor, index)),
    heightTwips
  );
  return surface.applyTableCommandPlan(plan);
}

function rowHeight(editor: Editor, index: number, heightTwips: number): void {
  expect(dragRowHeight(editor, index, heightTwips)).toEqual(LANDED);
}

/** Drag the divider between the two columns. */
function dragColumnDivider(editor: Editor): ExecResult {
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
}

function resizeColumns(editor: Editor): void {
  expect(dragColumnDivider(editor)).toEqual(LANDED);
}

async function peers(body = BODY) {
  const pair = await harness.pair(zipDocument(body));
  const alice = mount(pair.alice.room.document, 'Alice', pair.alice);
  const bob = mount(pair.bob.room.document, 'Bob', pair.bob);
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

test('a proposed row height replicates, undoes, redoes, reconnects and reopens', async () => {
  const { pair, alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  rowHeight(alice, 0, 900);
  converge();
  expect(cards(bob)).toEqual(['format']);
  expect(xmlOf(bob)).toContain('<w:trPrChange w:author="Alice"');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(cards(bob)).toEqual([]);
  expect(xmlOf(bob)).toBe(original);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(cards(bob)).toEqual(['format']);

  const rejoined = await harness.join(pair.alice, 'returning');
  const returning = mount(rejoined.room.document, 'Alice', rejoined);
  expect(cards(returning)).toEqual(['format']);
  expect(xmlOf(returning)).toBe(xmlOf(bob));

  const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(cards(reopened)).toEqual(['format']);
});

for (const action of ['accept', 'reject'] as const) {
  for (const [name, gesture] of [
    ['row height', (editor: Editor) => rowHeight(editor, 0, 900)],
    ['column resize', resizeColumns],
  ] as const) {
    test(`the other participant can ${action} a proposed ${name}`, async () => {
      const { alice, bob, converge } = await peers();
      const original = xmlOf(bob);
      suggest(alice);
      gesture(alice);
      converge();
      expect(cards(bob)).toEqual(['format']);
      const [card] = bob.getReviewItems();
      const resolved =
        action === 'accept' ? bob.acceptReviewItem(card!.key) : bob.rejectReviewItem(card!.key);
      expect(resolved.ok).toBe(true);
      converge();
      expect(cards(alice)).toEqual([]);
      expect(xmlOf(alice)).not.toMatch(RECORDS);
      if (action === 'accept') return;
      // A row with no `w:trPr` keeps the empty one its record needed; nothing else differs.
      expect(xmlOf(alice).replace('<w:tr><w:trPr/>', '<w:tr>')).toBe(original);
    });
  }
}

test("each participant's row height on a different row is its own card", async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  suggest(bob);
  pair.pause();
  rowHeight(alice, 0, 900);
  rowHeight(bob, 1, 1200);
  flush();
  pair.resume();
  converge();
  const authors = alice.getReviewItems().map((item) => item.kind === 'revision' && item.author);
  expect(authors).toEqual(['Alice', 'Bob']);
  // Rejecting Bob's leaves Alice's proposal and row exactly as they were.
  expect(alice.rejectReviewItem(alice.getReviewItems()[1]!.key).ok).toBe(true);
  converge();
  expect(cards(bob)).toEqual(['format']);
  expect(xmlOf(bob)).toContain('<w:trHeight w:hRule="exact" w:val="900"/>');
  expect(xmlOf(bob)).not.toContain('w:val="1200"');
});

test('typing in the row while its height is proposed keeps the text and the card', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  rowHeight(alice, 0, 900);
  caret(bob, A1, 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).toContain('a1!');
  expect(cards(alice)).toEqual(['format']);
  expect(alice.rejectReviewItem(alice.getReviewItems()[0]!.key).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toContain('a1!');
  expect(xmlOf(bob)).not.toMatch(RECORDS);
});

test('deleting the row while its height is proposed leaves no stray record', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  rowHeight(alice, 0, 900);
  expect(cards(alice)).toEqual(['format']);
  caret(bob, A1);
  expect(bob.exec({ type: 'deleteRow' }).ok).toBe(true);
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).not.toContain('a1');
  expect(xmlOf(alice)).not.toMatch(RECORDS);
  expect(cards(alice)).toEqual([]);
});

test("a proposed resize moves the other participant's view of the caret with the cells", async () => {
  const { pair, alice, bob, converge } = await peers();
  suggest(alice);
  caret(alice, B2, 1);
  resizeColumns(alice);
  converge();
  expect(cards(bob)).toEqual(['format']);
  applyAwarenessUpdate(
    pair.bob.awareness,
    encodeAwarenessUpdate(pair.alice.awareness, [pair.alice.ydoc.clientID]),
    'test'
  );
  const remote = pair.bob.room.session.remoteSelections()[0]!;
  const remoteCaret = caretAt(bob.surface!.layout(), {
    paragraphId: remote.head.nodeId,
    offset: remote.head.offset,
  });
  expect(remoteCaret).toEqual(
    caretAt(alice.surface!.layout(), alice.surface!.state().selection.head)
  );
});

// A ROW HOLDS ONE AUTHOR'S PENDING DECISION AT A TIME. Bob's proposed fill in a3 records on each
// cell of the first row. Alice's resize used to take some of those records over, and rejecting
// hers and then his reset those cells to `tcW auto`. Her change to the row is refused instead.
const HELD: ExecResult = { ok: false, code: 'unsupported', reason: HELD_ROW_REASON };
const authorsOf = (editor: Editor) =>
  editor.getReviewItems().map((item) => item.kind === 'revision' && item.author);
/** Every cell's `w:tcW`, as `type:width`. */
const widths = (editor: Editor): string[] =>
  [...xmlOf(editor).matchAll(/<w:tcW w:type="(\w+)" w:w="(\d+)"\/>/g)].map(
    ([, type, width]) => `${type}:${width}`
  );

async function withBobsFill() {
  const session = await peers(THREE_COLUMNS);
  suggest(session.alice);
  suggest(session.bob);
  caret(session.bob, A3);
  expect(
    session.bob.exec({ type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } })
  ).toEqual(LANDED);
  session.converge();
  return session;
}

test("a change to the row of the other participant's pending fill is refused", async () => {
  const { alice, bob, converge } = await withBobsFill();
  const held = xmlOf(bob);
  expect(dragColumnDivider(alice)).toEqual(HELD);
  expect(dragRowHeight(alice, 0, 900)).toEqual(HELD);
  converge();
  expect(xmlOf(alice)).toBe(held);
  expect(authorsOf(bob)).toEqual(['Bob']);
  // A row the fill does not hold takes the change.
  rowHeight(alice, 1, 900);
  converge();
  expect(authorsOf(bob)).toEqual(['Bob', 'Alice']);
});

for (const action of ['accept', 'reject'] as const) {
  test(`once the fill is answered (${action}), the resize lands and no width is lost`, async () => {
    const { alice, bob, converge } = await withBobsFill();
    expect(dragColumnDivider(alice)).toEqual(HELD);
    const [fill] = alice.getReviewItems();
    const answered =
      action === 'accept' ? alice.acceptReviewItem(fill!.key) : alice.rejectReviewItem(fill!.key);
    expect(answered.ok).toBe(true);
    converge();
    resizeColumns(alice);
    converge();
    expect(authorsOf(bob)).toEqual(['Alice']);
    expect(bob.rejectReviewItem(bob.getReviewItems()[0]!.key).ok).toBe(true);
    converge();
    expect(cards(alice)).toEqual([]);
    expect(widths(alice)).toEqual(Array(6).fill('dxa:3000'));
    expect(xmlOf(alice).includes('w:fill="FF0000"')).toBe(action === 'accept');
  });
}
