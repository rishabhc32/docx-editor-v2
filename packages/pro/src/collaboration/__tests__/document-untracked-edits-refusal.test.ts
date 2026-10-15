/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `reviewModule({ untrackedEdits: 'refuse' })` over a collaboration session. The refusal is a
// local decision: a refused edit commits and sends nothing, a peer's direct edit still applies
// on the refusing replica, and the suggestions the refusing peer can make converge, undo,
// survive a reconnect and reopen.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('untracked-edits-refusal', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const BODY =
  '<w:p><w:r><w:t>Alpha clause</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Beta clause text</w:t></w:r></w:p>' +
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>';

/** Alice suggests and refuses what she cannot record; Bob edits directly. */
function mount(peer: Peer, author: string, suggesting: boolean): Editor {
  peer.detach();
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: peer.room.document,
    author,
    mode: suggesting ? 'suggesting' : 'edit',
    modules: [
      reviewModule(suggesting ? { untrackedEdits: 'refuse' } : {}),
      collaborationModule({ session: peer.room.session }),
    ],
  });
  editors.push(editor);
  return editor;
}

function select(editor: Editor, paragraph: number, start: number, end = start) {
  const paragraphId = editor.surface!.session.paragraphIds()[paragraph]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

const xml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());
const fingerprint = (editor: Editor) =>
  packageFingerprint(editor.surface!.session.currentPackage());
const leftMargin = (editor: Editor) => /<w:pgMar[^>]*w:left="(\d+)"/.exec(xml(editor))?.[1];

test('a refused edit sends nothing while peer edits and suggestions converge', async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice, 'Alice', true);
  const bob = mount(pair.bob, 'Bob', false);
  const sync = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
  };
  const converge = () => {
    sync();
    expect(fingerprint(alice)).toBe(fingerprint(bob));
  };

  // Refused locally: nothing is committed, so nothing reaches the room.
  const before = fingerprint(bob);
  expect(alice.exec({ type: 'setPageSetup', marginLeft: 720 }).ok).toBe(false);
  converge();
  expect(fingerprint(bob)).toBe(before);
  expect(fingerprint(alice)).toBe(before);

  // A peer's direct edit is remote on Alice's replica and is never refused there.
  expect(bob.exec({ type: 'setPageSetup', marginLeft: 900 }).ok).toBe(true);
  converge();
  expect(leftMargin(alice)).toBe('900');

  // Concurrent: Alice suggests and tries a refused edit while Bob types and deletes text in
  // the paragraph she is suggesting in, all while the link is down.
  pair.pause();
  select(alice, 1, 4);
  alice.surface!.type(' new');
  expect(alice.exec({ type: 'setPageSetup', orientation: 'landscape' }).ok).toBe(false);
  select(bob, 1, 11, 16);
  bob.surface!.deleteBackward();
  select(bob, 0, 0);
  bob.surface!.type('B ');
  sync();
  pair.resume();
  converge();
  expect(xml(bob)).toContain('w:author="Alice"');
  expect(xml(bob)).not.toContain('w:orient="landscape"');
  expect(bob.surface!.session.bodyText()).toContain('Beta new clause');
  expect(bob.surface!.session.bodyText()).toContain('B Alpha clause');

  // Undo and redo of the suggestion converge on both replicas.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).not.toContain(' new');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).toContain('Beta new clause');

  // A replica that rejoins sees the same document, and refuses for itself.
  const rejoined = mount(await harness.join(pair.alice, 'rejoined'), 'Carol', true);
  expect(fingerprint(rejoined)).toBe(fingerprint(alice));
  expect(rejoined.exec({ type: 'setPageSetup', marginLeft: 720 }).ok).toBe(false);
  expect(fingerprint(rejoined)).toBe(fingerprint(alice));

  // Save and reopen keep the suggestion and the peer's direct edit.
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const reopened = createDocxEditor({
    container,
    document: new Uint8Array(await alice.save()),
    modules: [reviewModule()],
  });
  editors.push(reopened);
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(alice.surface!.session.currentPackage())
  );
  expect(reopened.getTrackedChanges().some((change) => change.author === 'Alice')).toBe(true);
  expect(leftMargin(reopened)).toBe('900');
});

test('the refusing peer removes its own pending link, and the peers converge', async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice, 'Alice', true);
  const bob = mount(pair.bob, 'Bob', false);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(fingerprint(alice)).toBe(fingerprint(bob));
  };
  select(alice, 0, 0, 'Alpha'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://a.example/' }).ok).toBe(true);
  converge();

  // Alice takes her pending link off again while Bob types in the same paragraph.
  pair.pause();
  select(alice, 0, 'AlphaAl'.length);
  const unlink = alice.exec({ type: 'removeHyperlink' });
  expect(unlink.ok ? null : unlink.reason).toBeNull();
  select(bob, 0, 'AlphaAlpha clause'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(xml(bob)).not.toMatch(/w:hyperlink|w:rStyle/);
  expect(bob.surface!.session.bodyText()).toContain('AlphaAlpha clause!');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).toContain('w:hyperlink');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toContain('w:hyperlink');

  // Bob rejects: the words are back as they were, with no link.
  expect(bob.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
  converge();
  expect(alice.surface!.session.bodyText()).toContain('Alpha clause!');
  expect(alice.getTrackedChanges()).toEqual([]);
});

test('the refusing peer links words in a paragraph made in the room, and re-aims the link', async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice, 'Alice', true);
  const bob = mount(pair.bob, 'Bob', false);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(fingerprint(alice)).toBe(fingerprint(bob));
  };
  // Bob adds a paragraph, so its node ids are made in the room and name no part.
  select(bob, 1, 'Beta clause text'.length);
  bob.surface!.splitParagraph();
  bob.surface!.type('Gamma words');
  converge();

  select(alice, 2, 0, 'Gamma'.length);
  const link = alice.exec({ type: 'insertHyperlink', href: 'https://a.example/' });
  expect(link.ok ? null : link.reason).toBeNull();
  select(alice, 2, 'GammaGa'.length);
  const reaim = alice.exec({ type: 'insertHyperlink', href: 'https://b.example/' });
  expect(reaim.ok ? null : reaim.reason).toBeNull();
  converge();
  expect(xml(bob)).toContain('w:hyperlink');
  expect(bob.getTrackedChanges().every((change) => change.author === 'Alice')).toBe(true);
});

test('the refusing peer removes and re-aims links another peer made in the room', async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice, 'Alice', true);
  const bob = mount(pair.bob, 'Bob', false);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(fingerprint(alice)).toBe(fingerprint(bob));
  };
  // Bob links words in the file's own paragraphs; the links' node ids are made in the room.
  select(bob, 0, 0, 'Alpha'.length);
  expect(bob.exec({ type: 'insertHyperlink', href: 'https://a.example/' }).ok).toBe(true);
  select(bob, 1, 0, 'Beta'.length);
  expect(bob.exec({ type: 'insertHyperlink', href: 'https://b.example/' }).ok).toBe(true);
  converge();
  const linked = xml(bob);

  select(alice, 0, 'Al'.length);
  expect(alice.can({ type: 'removeHyperlink' }).ok).toBe(true);
  const unlink = alice.exec({ type: 'removeHyperlink' });
  expect(unlink.ok ? null : unlink.reason).toBeNull();
  select(alice, 1, 'Be'.length);
  const reaim = alice.exec({ type: 'insertHyperlink', href: 'https://c.example/' });
  expect(reaim.ok ? null : reaim.reason).toBeNull();
  converge();
  expect(bob.getTrackedChanges().every((change) => change.author === 'Alice')).toBe(true);

  // Bob rejects both proposals: his links are back as he made them.
  expect(bob.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
  converge();
  expect(alice.surface!.session.bodyText()).toBe(bob.surface!.session.bodyText());
  expect(xml(alice).match(/<w:hyperlink /g)).toHaveLength(linked.match(/<w:hyperlink /g)!.length);
});
