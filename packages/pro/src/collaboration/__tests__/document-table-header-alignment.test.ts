/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Repeat Header Rows and table alignment, between two participants.
//
// Both write ordinary table properties — `w:trPr/w:tblHeader` and `w:tblPr/w:jc` — and, in
// suggesting mode, ordinary tracked markup: `w:trPrChange`, `w:tblPrChange` and the cells'
// `w:tcPrChange`. So they travel as any other edit does. Each participant sees the same card
// and can answer it, undo and redo replicate, concurrent edits converge, and a returning
// participant and a reopened save read the same document.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { mainPart, packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('table-header-alignment', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
// Paragraphs in order: before, a1, a2, b1, b2, c1, c2, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="2880" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="1440"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}</w:tr><w:tr>${cell('b1')}${cell('b2')}</w:tr>` +
  `<w:tr>${cell('c1')}${cell('c2')}</w:tr></w:tbl>` +
  p('after');
const A1 = 1;
const B1 = 3;
const C2 = 6;

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
const authors = (editor: Editor) =>
  editor.getReviewItems().map((item) => (item.kind === 'revision' ? item.author : item.kind));
const RECORDS = /<w:(trPrChange|tcPrChange|tblPrChange|tblGridChange)\b/;
const HEADER = '<w:tblHeader w:val="1"/>';
const CENTERED = '<w:jc w:val="center"/>';
/** Rows that carry `w:tblHeader`, in order; a record's copy of the old row does not count. */
const headerRows = (editor: Editor) =>
  (xmlOf(editor).match(/<w:tr>.*?<\/w:tr>/g) ?? []).map((row) =>
    row.replace(/<w:trPrChange\b.*?<\/w:trPrChange>/, '').includes(HEADER)
  );

/** Repeat Header Rows with the caret in the first cell of row `paragraph` sits in. */
function toggleHeader(editor: Editor, paragraph: number): void {
  caret(editor, paragraph);
  expect(editor.exec({ type: 'toggleHeaderRow' })).toEqual({ ok: true, changed: true });
}

function centerTable(editor: Editor, paragraph = A1): void {
  caret(editor, paragraph);
  expect(editor.exec({ type: 'setTableProperties', justification: 'center' })).toEqual({
    ok: true,
    changed: true,
  });
}

async function peers() {
  const pair = await harness.pair(zipDocument(BODY));
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

test('a proposed header row replicates, undoes, redoes, reconnects and reopens', async () => {
  const { pair, alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  toggleHeader(alice, B1);
  converge();
  expect(cards(bob)).toEqual(['format']);
  expect(headerRows(bob)).toEqual([true, true, false]);
  expect(bob.isActive({ type: 'toggleHeaderRow' })).toBe(false);
  caret(bob, A1);
  expect(bob.isActive({ type: 'toggleHeaderRow' })).toBe(true);

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

test('a proposed alignment replicates, undoes, redoes and reopens', async () => {
  const { alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  centerTable(alice);
  converge();
  expect(cards(bob)).toEqual(['format']);
  expect(xmlOf(bob)).toContain(CENTERED);
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toBe(original);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(cards(bob)).toEqual(['format']);
  const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(cards(reopened)).toEqual(['format']);
});

for (const action of ['accept', 'reject'] as const) {
  for (const [name, gesture, applied] of [
    ['header row', (editor: Editor) => toggleHeader(editor, B1), HEADER],
    ['alignment', (editor: Editor) => centerTable(editor), CENTERED],
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
      if (action === 'accept') {
        expect(xmlOf(alice)).toContain(applied);
        return;
      }
      // A row with no `w:trPr` keeps the empty one its record needed; nothing else differs.
      expect(xmlOf(alice).replaceAll('<w:tr><w:trPr/>', '<w:tr>')).toBe(original);
    });
  }
}

test('a header row and an alignment made at once both land, untracked', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  pair.pause();
  toggleHeader(alice, A1);
  centerTable(bob, C2);
  flush();
  pair.resume();
  converge();
  expect(headerRows(alice)).toEqual([true, false, false]);
  expect(xmlOf(alice)).toContain(CENTERED);
  expect(xmlOf(alice)).not.toMatch(RECORDS);
});

for (const proposer of ['header row', 'alignment'] as const) {
  test(`a ${proposer} proposed while the other participant edits the other keeps both`, async () => {
    const { pair, alice, bob, flush, converge } = await peers();
    const original = xmlOf(bob);
    suggest(alice);
    pair.pause();
    if (proposer === 'header row') {
      toggleHeader(alice, A1);
      centerTable(bob, C2);
    } else {
      centerTable(alice, C2);
      toggleHeader(bob, A1);
    }
    flush();
    pair.resume();
    converge();
    expect(authors(bob)).toEqual(['Alice']);
    expect(headerRows(bob)).toEqual([true, false, false]);
    expect(xmlOf(bob)).toContain(CENTERED);
    expect(bob.rejectReviewItem(bob.getReviewItems()[0]!.key).ok).toBe(true);
    converge();
    // Rejecting the proposal takes back only what it proposed; the untracked edit stays.
    expect(xmlOf(alice)).not.toMatch(RECORDS);
    expect(headerRows(alice)).toEqual([proposer !== 'header row', false, false]);
    expect(xmlOf(alice).includes(CENTERED)).toBe(proposer !== 'alignment');
    expect(xmlOf(alice)).not.toBe(original);
  });
}

test('deleting the row while its header flag is proposed leaves no stray record', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  toggleHeader(alice, A1);
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

test('typing in a header row while it is proposed keeps the text and the card', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  toggleHeader(alice, B1);
  caret(bob, B1, 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).toContain('b1!');
  expect(cards(alice)).toEqual(['format']);
  expect(alice.rejectReviewItem(alice.getReviewItems()[0]!.key).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toContain('b1!');
  expect(xmlOf(bob)).not.toMatch(RECORDS);
});

test("a proposed alignment moves the other participant's view of the caret with the table", async () => {
  const { pair, alice, bob, converge } = await peers();
  suggest(alice);
  caret(alice, C2, 1);
  const before = caretAt(bob.surface!.layout(), {
    paragraphId: bob.surface!.session.paragraphIds()[C2]!,
    offset: 1,
  });
  centerTable(alice, C2);
  caret(alice, C2, 1);
  converge();
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
  expect(remoteCaret!.x).toBeGreaterThan(before!.x + 100);
});
