/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A format suggestion taken back with a second press leaves no suggestion on either replica,
// beside a concurrent edit, through undo and redo, and after save and reopen.
import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('toggle-suggestion');
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});
function mount(peer: Peer, author: string): Editor {
  peer.detach();
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: peer.room.document,
    author,
    mode: 'suggesting',
    modules: [reviewModule(), collaborationModule({ session: peer.room.session })],
  });
  editors.push(editor);
  return editor;
}
function select(editor: Editor, start: number, end = start) {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}
const xml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());

test('Underline pressed twice leaves no suggestion on either replica', async () => {
  const pair = await harness.pair(zipDocument('<w:p><w:r><w:t>Anchor text</w:t></w:r></w:p>'));
  const alice = mount(pair.alice, 'Alice');
  const bob = mount(pair.bob, 'Bob');
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };

  select(alice, 0, 6);
  alice.surface!.toggleRunProperty('u', { val: 'single' });
  converge();
  expect(xml(bob)).toContain('rPrChange');

  // The second press lands while Bob types elsewhere in the paragraph.
  pair.pause();
  alice.surface!.toggleRunProperty('u', { val: 'single' });
  select(bob, 11);
  bob.surface!.type('!');
  pair.resume();
  converge();
  for (const editor of [alice, bob]) {
    expect(xml(editor)).not.toContain('rPrChange');
    expect(xml(editor)).not.toContain('w:u');
    expect(xml(editor)).toContain('w:author="Bob"');
  }

  // Undo brings the suggestion back on both replicas, and redo takes it back again.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).toContain('rPrChange');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toContain('rPrChange');

  const saved = new Uint8Array(await bob.save());
  const reopened = createDocxEditor({ document: saved, modules: [reviewModule()] });
  editors.push(reopened);
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  reopened.attach(container);
  expect(xml(reopened)).not.toContain('rPrChange');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(alice.surface!.session.currentPackage())
  );
});
