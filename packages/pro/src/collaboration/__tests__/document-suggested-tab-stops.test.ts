/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Tab stops one participant suggests replicate as a suggestion.
//
// The paragraph dialog records the paragraph's old `w:tabs` in a `w:pPrChange`. The peer must
// see the new stops, the record and one review card; rejecting there must converge, and so must
// a race with the peer typing in the same paragraph, through undo and redo.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('suggested-tab-stops', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const BODY =
  '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>' +
  '<w:r><w:t>alpha</w:t></w:r></w:p>';
const STOPS = [{ positionTwips: 2880, alignment: 'right' as const }];

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
const RECORD =
  /<w:pPrChange [^>]*w:author="Alice"[^>]*><w:pPr><w:tabs><w:tab w:pos="720" w:val="left"\/>/;

test('suggested tab stops converge, reopen, and a peer rejection converges', async () => {
  const { alice, bob, converge } = await peers();
  caret(alice, 2);
  expect(alice.exec({ type: 'setParagraphFormat', tabStops: STOPS })).toMatchObject({ ok: true });
  converge();

  expect(xml(bob)).toContain('<w:tab w:pos="2880" w:val="right"/>');
  expect(xml(bob)).toMatch(RECORD);
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
    expect(xml(editor)).not.toContain('pPrChange');
    expect(xml(editor)).not.toContain('w:pos="2880"');
    expect(xml(editor)).toContain('<w:tab w:pos="720" w:val="left"/>');
    expect(revisions(editor)).toEqual([]);
  }
});

test('suggested tab stops racing the peer typing in the paragraph converge and undo', async () => {
  const { pair, alice, bob, converge } = await peers();
  pair.pause();
  caret(alice, 2);
  expect(alice.exec({ type: 'setParagraphFormat', tabStops: STOPS })).toMatchObject({ ok: true });
  caret(bob, 5);
  bob.surface!.type(' beta');
  bob.surface!.flushPendingInput();
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  expect(alice.surface!.session.bodyText()).toBe('alpha beta');
  expect(xml(bob)).toMatch(RECORD);
  expect(revisions(bob)).toEqual(['format:Alice']);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toContain('pPrChange');
  expect(alice.surface!.session.bodyText()).toBe('alpha beta');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).toMatch(RECORD);
});
