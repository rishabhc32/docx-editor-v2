/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `headerFooterEditing: false` is a local UI option. The peer that has it cannot open a header,
// but header edits from a peer without it still arrive, converge, and survive save and reopen,
// through concurrent body typing, deletion, undo, redo, and a session remount.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule } from '../collaboration-module.ts';
import { createPeerHarness, R, W, zipDocument, type Peer } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('header-footer-editing', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const HEADER = { kind: 'headerFooter', rId: 'rId7' } as const;
const EDIT_HEADER = { type: 'editHeaderFooter', position: 'header' } as const;

function headerDoc(): Uint8Array {
  return zipDocument(
    '<w:p><w:r><w:t>body</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/></w:sectPr>',
    {
      documentRels:
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId7" Type="${R}/header" Target="header1.xml"/></Relationships>`,
      overrides:
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
      extraXml: {
        'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
      },
    }
  );
}

function mount(options: Parameters<typeof createDocxEditor>[0]) {
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({ ...options, container });
  editors.push(editor);
  return editor;
}

function headerText(editor: Editor): string | null {
  return editor.surface!.session.storyText(HEADER);
}

function syncPresence(from: Peer, to: Peer) {
  const ids = [...from.awareness.getStates().keys()];
  applyAwarenessUpdate(to.awareness, encodeAwarenessUpdate(from.awareness, ids), 'relay');
}

/** Alice's headers are read-only; Bob keeps Word's header editing. */
async function peers() {
  const pair = await harness.pair(headerDoc());
  const open = (peer: Peer, headerFooterEditing: boolean) => {
    peer.detach();
    return mount({
      document: peer.room.document,
      modules: [collaborationModule({ session: peer.room.session })],
      ...(headerFooterEditing ? {} : { headerFooterEditing }),
    });
  };
  const alice = open(pair.alice, false);
  const bob = open(pair.bob, true);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  return { pair, alice, bob, converge };
}

test('a header edit from the other peer reaches the read-only peer and survives reopen', async () => {
  const { alice, bob, converge } = await peers();
  expect(alice.exec(EDIT_HEADER)).toMatchObject({ ok: false, code: 'locked' });
  expect(bob.exec(EDIT_HEADER).ok).toBe(true);
  bob.surface!.type('X');
  converge();
  expect(headerText(alice)).toBe('XHeader');
  expect(alice.getActiveScope()).toEqual({ kind: 'body' });

  const reopened = mount({ document: new Uint8Array(await alice.save()) });
  expect(headerText(reopened)).toBe('XHeader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(alice.surface!.session.currentPackage())
  );
});

test('body typing on the read-only peer and a header deletion on the other converge', async () => {
  const { pair, alice, bob, converge } = await peers();
  pair.pause();
  const body = alice.surface!.session.paragraphIds()[0]!;
  alice.surface!.setSelection({
    anchor: { paragraphId: body, offset: 4 },
    head: { paragraphId: body, offset: 4 },
  });
  alice.surface!.type('Y');
  expect(bob.exec(EDIT_HEADER).ok).toBe(true);
  const header = bob.surface!.session.paragraphIdsIn(HEADER)[0]!;
  bob.surface!.setSelection({
    anchor: { paragraphId: header, offset: 0 },
    head: { paragraphId: header, offset: 3 },
  });
  bob.surface!.deleteSelection();
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  expect(alice.surface!.session.bodyText()).toBe('bodyY');
  expect(headerText(alice)).toBe('der');
});

test("undo and redo of the other peer's header edit converge on the read-only peer", async () => {
  const { alice, bob, converge } = await peers();
  expect(bob.exec(EDIT_HEADER).ok).toBe(true);
  bob.surface!.type('X');
  converge();
  expect(bob.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(headerText(alice)).toBe('Header');
  expect(bob.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(headerText(alice)).toBe('XHeader');
});

test('the read-only peer sees a header caret, and stays read-only after a remount', async () => {
  const { pair, alice, bob, converge } = await peers();
  expect(bob.exec(EDIT_HEADER).ok).toBe(true);
  const header = bob.surface!.session.paragraphIdsIn(HEADER)[0]!;
  syncPresence(pair.bob, pair.alice);
  expect(pair.alice.room.session.remoteSelections()).toMatchObject([
    { actorId: 'bob', anchor: { nodeId: header, offset: 0 }, head: { nodeId: header, offset: 0 } },
  ]);

  bob.surface!.type('X');
  converge();
  alice.destroy();
  const next = await harness.remount(pair.alice);
  next.detach();
  const reconnected = mount({
    document: next.room.document,
    headerFooterEditing: false,
    modules: [collaborationModule({ session: next.room.session })],
  });
  expect(headerText(reconnected)).toBe('XHeader');
  expect(reconnected.exec(EDIT_HEADER)).toMatchObject({ ok: false, code: 'locked' });
  expect(reconnected.getActiveScope()).toEqual({ kind: 'body' });
});
