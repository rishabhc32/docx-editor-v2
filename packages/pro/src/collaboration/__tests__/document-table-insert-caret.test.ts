/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A new table, row or column takes the next keystroke in a shared room, where created
// paragraphs are renamed before the caret is placed (see `committed-caret.ts`).

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('table-insert-caret', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

function mount(options: Parameters<typeof createDocxEditor>[0]) {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({ ...options, container });
  editors.push(editor);
  return editor;
}

function caret(editor: Editor, index: number, offset: number) {
  const position = { paragraphId: editor.surface!.session.paragraphIds()[index]!, offset };
  editor.surface!.setSelection({ anchor: position, head: position });
}

function texts(editor: Editor): string[] {
  return editor.query({ type: 'paragraphs' }).map(({ text }) => text);
}

test('typing after a new table, row and column lands in the new cell on both peers', async () => {
  const pair = await harness.pair(zipDocument('<w:p><w:r><w:t>hello</w:t></w:r></w:p>'));
  const [alice, bob] = [pair.alice, pair.bob].map((peer) => {
    peer.detach();
    return mount({
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
    });
  }) as [Editor, Editor];
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };

  // A 1x2 table before "hello", with Bob typing in "hello" while it is in flight.
  pair.pause();
  caret(alice, 0, 'hello'.length);
  expect(alice.exec({ type: 'insertTable', rows: 1, cols: 2 }).ok).toBe(true);
  alice.surface!.type('A');
  caret(bob, 0, 'hello'.length);
  bob.surface!.type('!');
  expect(alice.surface!.state().lastRejection).toBeNull();
  pair.resume();
  converge();
  expect(texts(bob)).toEqual(['A', '', 'hello!']);

  // A new row lands the caret in its first cell, and a new column in the new cell.
  expect(alice.exec({ type: 'insertRow', where: 'below' }).ok).toBe(true);
  alice.surface!.type('R');
  expect(alice.surface!.state().lastRejection).toBeNull();
  caret(alice, 0, 1);
  expect(alice.exec({ type: 'insertColumn', where: 'right' }).ok).toBe(true);
  alice.surface!.type('C');
  expect(alice.surface!.state().lastRejection).toBeNull();
  converge();
  expect(texts(bob)).toEqual(['A', 'C', '', 'R', '', '', 'hello!']);

  // History travels as single steps, and a returning peer and a reopened file agree.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(texts(bob)).toEqual(['A', '', '', 'R', '', '', 'hello!']);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  const rejoined = await harness.join(pair.alice, 'returning');
  rejoined.detach();
  const returning = mount({
    document: rejoined.room.document,
    modules: [collaborationModule({ session: rejoined.room.session })],
  });
  expect(texts(returning)).toEqual(texts(alice));
  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
});
