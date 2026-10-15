/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A proposed table deletion and a proposed cell format change, between two participants.
//
// Both are ordinary tracked markup — `w:del` on every row, `w:tcPrChange` on the cells of a
// row — so they travel as any other edit does. Each participant sees the same card and can
// answer it, undo and redo replicate, concurrent edits keep their text, and a returning
// participant and a reopened save read the same proposal.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { mainPart, packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('tracked-table-edits', { offlineEditing: true });
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
const FILL = { type: 'setCellFill', color: { kind: 'hex', value: 'FF0000' } } as const;

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
  };
  return { pair, alice, bob, flush, converge };
}

test('a proposed table deletion replicates, undoes, redoes, reconnects and reopens', async () => {
  const { pair, alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  caret(alice, A1);
  expect(alice.exec({ type: 'deleteTable' })).toEqual({ ok: true, changed: true });
  converge();
  expect(cards(bob)).toEqual(['structural']);
  expect(xmlOf(bob).match(/<w:del w:author="Alice"/g)).toHaveLength(2);

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
  test(`the other participant can ${action} a proposed table deletion`, async () => {
    const { alice, bob, converge } = await peers();
    const original = xmlOf(bob);
    suggest(alice);
    caret(alice, A1);
    expect(alice.exec({ type: 'deleteTable' }).ok).toBe(true);
    converge();
    const [card] = bob.getReviewItems();
    const resolved =
      action === 'accept' ? bob.acceptReviewItem(card!.key) : bob.rejectReviewItem(card!.key);
    expect(resolved.ok).toBe(true);
    converge();
    expect(cards(alice)).toEqual([]);
    if (action === 'accept') expect(xmlOf(alice)).not.toContain('<w:tbl>');
    // Rejecting a row deletion leaves its emptied `w:trPr`; nothing else differs.
    else expect(xmlOf(alice).replaceAll('<w:trPr/>', '')).toBe(original);
  });

  test(`the other participant can ${action} a proposed cell format change`, async () => {
    const { alice, bob, converge } = await peers();
    const original = xmlOf(bob);
    suggest(alice);
    caret(alice, A1);
    expect(alice.exec(FILL).ok).toBe(true);
    converge();
    expect(cards(bob)).toEqual(['format']);
    const [card] = bob.getReviewItems();
    const resolved =
      action === 'accept' ? bob.acceptReviewItem(card!.key) : bob.rejectReviewItem(card!.key);
    expect(resolved.ok).toBe(true);
    converge();
    expect(cards(alice)).toEqual([]);
    expect(xmlOf(alice)).not.toContain('<w:tcPrChange');
    if (action === 'accept') expect(xmlOf(alice)).toContain('w:fill="FF0000"');
    else expect(xmlOf(alice)).toBe(original);
  });
}

test('a cell format proposal undoes and redoes for both participants', async () => {
  const { alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  caret(alice, A1);
  expect(alice.exec(FILL).ok).toBe(true);
  converge();
  const proposed = xmlOf(bob);
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toBe(original);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toBe(proposed);
  expect(cards(bob)).toEqual(['format']);
});

test('typing in a cell while the table deletion is proposed keeps the text, struck with it', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  caret(alice, A1);
  expect(alice.exec({ type: 'deleteTable' }).ok).toBe(true);
  caret(bob, B2, 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).toContain('b2!');
  expect(xmlOf(alice).match(/<w:del w:author="Alice"/g)).toHaveLength(2);
  expect(cards(alice)).toEqual(['structural']);
  // Accepting takes the table, and the words typed into it, together.
  expect(alice.acceptReviewItem(alice.getReviewItems()[0]!.key).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).not.toContain('<w:tbl>');
});

test('deleting the table while a cell format change is proposed leaves no stray record', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  pair.pause();
  caret(alice, A1);
  expect(alice.exec(FILL).ok).toBe(true);
  expect(cards(alice)).toEqual(['format']);
  caret(bob, B2);
  expect(bob.exec({ type: 'deleteTable' }).ok).toBe(true);
  flush();
  pair.resume();
  converge();
  expect(xmlOf(alice)).not.toContain('<w:tbl>');
  expect(xmlOf(alice)).not.toContain('<w:tcPrChange');
  expect(cards(alice)).toEqual([]);
});

test("deleting a table withdraws its author's proposed row for both participants", async () => {
  const { alice, bob, converge } = await peers();
  const original = xmlOf(bob);
  suggest(alice);
  caret(alice, A1);
  expect(alice.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  converge();
  const proposed = xmlOf(bob);
  caret(alice, A1);
  expect(alice.exec({ type: 'deleteTable' })).toEqual({ ok: true, changed: true });
  converge();
  expect(cards(bob)).toEqual(['structural']);
  expect(xmlOf(bob).match(/<w:tr>/g)).toHaveLength(2);
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xmlOf(bob)).toBe(proposed);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xmlOf(bob).match(/<w:tr>/g)).toHaveLength(2);
  expect(bob.rejectReviewItem(bob.getReviewItems()[0]!.key).ok).toBe(true);
  converge();
  expect(cards(alice)).toEqual([]);
  expect(xmlOf(alice).replaceAll('<w:trPr/>', '')).toBe(original);
});

test('typing in a proposed row while its author deletes the table converges', async () => {
  const { pair, alice, bob, flush, converge } = await peers();
  suggest(alice);
  caret(alice, A1);
  expect(alice.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  converge();
  pair.pause();
  caret(alice, A1);
  expect(alice.exec({ type: 'deleteTable' }).ok).toBe(true);
  caret(bob, A1 + 2);
  bob.surface!.type('!');
  flush();
  pair.resume();
  converge();
  // The withdrawal removes the row and what was typed into it, as rejecting the row would.
  expect(xmlOf(alice)).not.toContain('!');
  expect(xmlOf(alice).match(/<w:tr>/g)).toHaveLength(2);
  expect(cards(alice)).toEqual(['structural']);
});
