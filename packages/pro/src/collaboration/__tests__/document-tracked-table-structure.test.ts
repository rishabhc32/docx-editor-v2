/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Proposed columns, a proposed table under a range deletion, and the review of a proposed table
// its author typed into, between two participants.
//
// All of it is ordinary tracked markup — `w:cellDel` and `w:cellIns` on a column's cells, `w:del`
// on every row of a table — so it travels as any other edit does. Each participant sees the same
// card and can answer it, undo and redo replicate, concurrent edits converge, and a returning
// participant and a reopened save read the same proposal.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { mainPart, packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('tracked-table-structure', { offlineEditing: true });
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
// Paragraphs in order: before, a1, a2, a3, b1, b2, b3, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
  '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}${cell('a3')}</w:tr>` +
  `<w:tr>${cell('b1')}${cell('b2')}${cell('b3')}</w:tr></w:tbl>` +
  p('after');
const A2 = 2;
const B1 = 4;
const B2 = 5;
const AFTER = 7;

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
const cardOf = (editor: Editor, kind: string) =>
  editor.getReviewItems().find((item) => item.kind === 'revision' && item.revisionKind === kind)!;
const markers = (editor: Editor, name: 'cellDel' | 'cellIns') =>
  xmlOf(editor).match(new RegExp(`<w:${name} `, 'g'))?.length ?? 0;

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
  };
  return { pair, alice, bob, flush, converge };
}

const COLUMN_COMMANDS = [
  { name: 'deletion', command: { type: 'deleteColumn' }, marker: 'cellDel' },
  { name: 'insertion', command: { type: 'insertColumn', where: 'right' }, marker: 'cellIns' },
] as const;

for (const { name, command, marker } of COLUMN_COMMANDS) {
  test(`a proposed column ${name} replicates, undoes, redoes, reconnects and reopens`, async () => {
    const { pair, alice, bob, converge } = await peers();
    const original = xmlOf(bob);
    suggest(alice);
    caret(alice, A2);
    expect(alice.exec(command)).toEqual({ ok: true, changed: true });
    converge();
    expect(cards(bob)).toEqual(['structural']);
    expect(markers(bob, marker)).toBe(2);

    expect(alice.exec({ type: 'undo' }).ok).toBe(true);
    converge();
    expect(cards(bob)).toEqual([]);
    expect(xmlOf(bob)).toBe(original);
    expect(alice.exec({ type: 'redo' }).ok).toBe(true);
    converge();
    expect(cards(bob)).toEqual(['structural']);

    const rejoined = await harness.join(pair.alice, 'returning');
    const returning = mount(rejoined.room.document, 'Alice', rejoined);
    expect(cards(returning)).toEqual(['structural']);
    expect(xmlOf(returning)).toBe(xmlOf(bob));

    const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
    expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
      saveReopenDigest(bob.surface!.session.currentPackage())
    );
    expect(cards(reopened)).toEqual(['structural']);
  });

  for (const action of ['accept', 'reject'] as const) {
    test(`the other participant can ${action} a proposed column ${name}`, async () => {
      const { alice, bob, converge } = await peers();
      const original = xmlOf(bob);
      suggest(alice);
      caret(alice, A2);
      expect(alice.exec(command).ok).toBe(true);
      converge();
      const key = cardOf(bob, 'structural').key;
      const resolved = action === 'accept' ? bob.acceptReviewItem(key) : bob.rejectReviewItem(key);
      expect(resolved.ok).toBe(true);
      converge();
      expect(cards(alice)).toEqual([]);
      expect(xmlOf(alice)).not.toMatch(/<w:(cellDel|cellIns)\b/);
      if (action === 'reject') expect(xmlOf(alice)).toBe(original);
      else expect(xmlOf(alice).includes('a2')).toBe(name === 'insertion');
    });
  }
}

test('typing in a column while its deletion is proposed keeps the text, and Accept takes it', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  caret(alice, A2);
  expect(alice.exec({ type: 'deleteColumn' }).ok).toBe(true);
  caret(bob, B2, 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).toContain('b2!');
  expect(markers(alice, 'cellDel')).toBe(2);
  expect(cards(alice)).toEqual(['structural']);
  expect(alice.acceptReviewItem(cardOf(alice, 'structural').key).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).not.toContain('b2');
});

test("another participant's suggestion in the column holds the deletion until it is answered", async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  suggest(bob);
  pair.pause();
  caret(alice, A2);
  expect(alice.exec({ type: 'deleteColumn' }).ok).toBe(true);
  caret(bob, B2, 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  expect(cards(alice).sort()).toEqual(['insert', 'structural']);
  const refused = alice.acceptReviewItem(cardOf(alice, 'structural').key);
  expect(refused.ok).toBe(false);
  expect(refused.ok ? '' : refused.reason).toContain('another pending change');
  // Bob answers his own suggestion; the deletion can then be accepted.
  expect(bob.rejectReviewItem(cardOf(bob, 'insert').key).ok).toBe(true);
  converge();
  expect(alice.acceptReviewItem(cardOf(alice, 'structural').key).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).not.toMatch(/b2|cellDel/);
});

test('deleting a row while the column deletion is proposed converges', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  caret(alice, A2);
  expect(alice.exec({ type: 'deleteColumn' }).ok).toBe(true);
  caret(bob, B1);
  expect(bob.exec({ type: 'deleteRow' }).ok).toBe(true);
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).not.toContain('b1');
  expect(markers(alice, 'cellDel')).toBe(1);
  expect(cards(bob)).toEqual(['structural']);
  expect(bob.acceptReviewItem(cardOf(bob, 'structural').key).ok).toBe(true);
  converge();
  expect(xmlOf(alice)).not.toMatch(/a2|cellDel/);
});

test('a range deletion over a table replicates as a proposal, and rejecting it restores the table', async () => {
  const { alice, bob, converge } = await peers();
  suggest(alice);
  const ids = alice.surface!.session.paragraphIds();
  alice.surface!.setSelection({
    anchor: { paragraphId: ids[0]!, offset: 2 },
    head: { paragraphId: ids[AFTER]!, offset: 2 },
  });
  alice.surface!.deleteForward();
  converge();
  expect(cards(bob)).toEqual(['delete', 'structural', 'delete']);
  expect(xmlOf(bob)).toContain('<w:tbl>');
  for (const card of bob.getReviewItems()) expect(bob.rejectReviewItem(card.key).ok).toBe(true);
  converge();
  expect(cards(alice)).toEqual([]);
  expect(xmlOf(alice)).not.toMatch(/<w:(del|cellDel)\b/);
  expect(xmlOf(alice)).toContain('a1');
});

test("rejecting a proposed table takes its author's typing with it for both participants", async () => {
  const { alice, bob, converge } = await peers(p('before'));
  const original = xmlOf(bob);
  suggest(alice);
  caret(alice, 0, 6);
  expect(alice.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true);
  // The first cell is the first paragraph: the table goes in before the anchor paragraph.
  caret(alice, 0);
  alice.surface!.type('hi');
  converge();
  expect(cards(bob)).toEqual(['structural', 'insert']);
  expect(bob.rejectReviewItem(cardOf(bob, 'structural').key)).toEqual({ ok: true, changed: true });
  converge();
  expect(cards(alice)).toEqual([]);
  expect(xmlOf(alice)).toBe(original);
});
