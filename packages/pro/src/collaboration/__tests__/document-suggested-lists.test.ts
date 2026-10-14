/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A bullet list one participant suggests replicates as a suggestion.
//
// The press writes the list definition to `numbering.xml` untracked and records the paragraph's
// old `w:pPr` in a `w:pPrChange`. The peer must see the list, the record and one review card;
// rejecting the card there must converge, and so must a race with the peer's own deletion.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor, runToolbarCommand } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('suggested-lists', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

// No styles part, so the press writes `w:numPr` alone: no List Paragraph write records it.
const BODY = '<w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:p><w:r><w:t>beta</w:t></w:r></w:p>';

function mount(document: Uint8Array, author: string, peer?: Peer, mode?: 'suggesting') {
  peer?.detach();
  const container = window.document.createElement('div');
  window.document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    ...(mode ? { mode } : {}),
    modules: [
      reviewModule(),
      ...(peer ? [collaborationModule({ session: peer.room.session })] : []),
    ],
  });
  editors.push(editor);
  return editor;
}

async function peers() {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice.room.document, 'Alice', pair.alice, 'suggesting');
  const bob = mount(pair.bob.room.document, 'Bob', pair.bob);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    const left = alice.surface!.session.currentPackage();
    const right = bob.surface!.session.currentPackage();
    expect(packageFingerprint(right)).toBe(packageFingerprint(left));
    expect(saveReopenDigest(right)).toEqual(saveReopenDigest(left));
  };
  return { pair, alice, bob, converge };
}

function caret(editor: Editor, offset: number): void {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset },
    head: { paragraphId, offset },
  });
}

const xml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());
const revisions = (editor: Editor) =>
  editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [`${item.revisionKind}:${item.author}`] : []));
/** Whether a paragraph paints a list marker, which needs the replicated definition. */
const bulleted = (editor: Editor) =>
  editor
    .surface!.layout()
    .pages[0]!.fragments.some((fragment) => fragment.kind === 'paragraph' && !!fragment.marker);

test('a suggested bullet list converges, reopens, and a peer rejection converges', async () => {
  const { alice, bob, converge } = await peers();
  caret(alice, 2);
  expect(runToolbarCommand(alice, 'list.bullet')).toMatchObject({ ok: true, changed: true });
  converge();

  expect(xml(bob)).toContain('<w:numPr>');
  expect(xml(bob)).toMatch(/<w:pPrChange [^>]*w:author="Alice"/);
  expect(bulleted(bob)).toBe(true);
  expect(revisions(bob)).toEqual(['format:Alice']);

  const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(revisions(reopened)).toEqual(['format:Alice']);

  const card = bob.getReviewItems().find((item) => item.kind === 'revision')!;
  expect(bob.rejectReviewItem(card.key).ok).toBe(true);
  converge();
  for (const editor of [alice, bob]) {
    expect(xml(editor)).not.toContain('<w:numPr>');
    expect(xml(editor)).not.toContain('pPrChange');
    expect(bulleted(editor)).toBe(false);
    expect(revisions(editor)).toEqual([]);
  }
});

test('a suggested bullet list racing the peer deleting its text converges and undoes', async () => {
  const { pair, alice, bob, converge } = await peers();
  pair.pause();
  caret(alice, 2);
  expect(runToolbarCommand(alice, 'list.bullet')).toMatchObject({ ok: true, changed: true });
  const paragraphId = bob.surface!.session.paragraphIds()[0]!;
  bob.surface!.setSelection({
    anchor: { paragraphId, offset: 0 },
    head: { paragraphId, offset: 3 },
  });
  bob.surface!.deleteSelection();
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  expect(alice.surface!.session.bodyText()).toContain('ha');
  expect(alice.surface!.session.bodyText()).not.toContain('alpha');
  expect(revisions(bob)).toEqual(['format:Alice']);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toContain('pPrChange');
  expect(xml(bob)).not.toContain('<w:numPr>');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(revisions(bob)).toEqual(['format:Alice']);
});
