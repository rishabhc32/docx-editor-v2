/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A first link's `Hyperlink` style, a suggested new target or new words, and a suggested
// removal, across two editors, also of a link the other author proposed.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  collectReviewItems,
  listBookmarks,
  paragraphTextOf,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { packageFingerprint, saveReopenDigest } from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('hyperlink-commands', { offlineEditing: true });
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

function select(editor: Editor, start: number, end = start) {
  const paragraphId = editor.surface!.session.paragraphIds()[0]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
}

/** Each link in the first paragraph as [live text, href, tooltip]. */
function links(editor: Editor) {
  select(editor, 0);
  return editor
    .surface!.hyperlinks.linksInCaretParagraph()
    .map((link) => [link.text, link.href, link.tooltip]);
}

/** The character style of each run in the body, in order. */
function runStyles(editor: Editor) {
  return [
    ...serializeOoxmlPart(editor.surface!.session.part()).matchAll(/<w:rStyle w:val="([^"]+)"/g),
  ].map((match) => match[1]);
}

/** The text the document's one bookmark holds. */
function bookmarkedText(editor: Editor) {
  const part = editor.surface!.session.part();
  const { start, end } = listBookmarks(part)[0]!;
  return paragraphTextOf(part, start.paragraphId)?.slice(start.offset, end.offset);
}

async function peers(body = '<w:p><w:r><w:t xml:space="preserve">See Docs now</w:t></w:r></w:p>') {
  // No styles part: the first link has to bring Word's definition with it.
  const pair = await harness.pair(zipDocument(body));
  const [alice, bob] = (
    [
      [pair.alice, 'Alice'],
      [pair.bob, 'Bob'],
    ] as const
  ).map(([peer, author]) => {
    peer.detach();
    return mount({
      document: peer.room.document,
      author,
      modules: [reviewModule(), collaborationModule({ session: peer.room.session })],
    });
  }) as [Editor, Editor];
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  return { pair, alice, bob, converge };
}

test('a first link and a suggested new target converge, undo, redo, reconnect and reopen', async () => {
  const { pair, alice, bob, converge } = await peers();

  // The first link, with a tooltip: the style definition travels with it.
  select(alice, 'See '.length, 'See Docs'.length);
  expect(
    alice.exec({ type: 'insertHyperlink', href: 'https://example.com', tooltip: 'Tip' }).ok
  ).toBe(true);
  converge();
  expect(bob.surface!.session.documentStyles().map((style) => style.styleId)).toContain(
    'Hyperlink'
  );
  expect(links(bob)).toEqual([['Docs', 'https://example.com', 'Tip']]);

  // A new target while suggesting, with Bob typing in the same paragraph meanwhile.
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  select(bob, 'See Docs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([
    ['', 'https://example.com', 'Tip'],
    ['Docs', 'https://example.org', 'Tip'],
  ]);
  expect(bob.surface!.session.bodyText()).toBe('See DocsDocs now!');

  // Undo and redo of the proposal reach the peer as one step each.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', 'Tip']]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toHaveLength(2);

  // A returning peer reads the proposal, and the reopened file keeps it.
  const rejoined = await harness.join(pair.alice, 'returning');
  rejoined.detach();
  const returning = mount({
    document: rejoined.room.document,
    modules: [reviewModule(), collaborationModule({ session: rejoined.room.session })],
  });
  expect(links(returning)).toEqual(links(bob));
  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );

  // Bob rejects: the old link comes back whole on both screens.
  expect(bob.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.com', 'Tip']]);
});

test('deleting the link while a new target is suggested converges', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  // Bob removes the old link's words outright while the proposal is in flight.
  select(bob, 'See '.length, 'See Docs'.length);
  bob.surface!.deleteBackward();
  pair.resume();
  converge();
  // Alice's proposed copy survives; only the words Bob saw are gone.
  expect(bob.surface!.session.bodyText()).toBe('See Docs now');
  expect(links(bob).filter(([text]) => text !== '')).toEqual([
    ['Docs', 'https://example.org', undefined],
  ]);
});

test('suggested new words for a link keep its style, converge, undo, redo, reopen and accept', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();

  // Alice proposes new words for the link while Bob types in the same paragraph.
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(
    alice.exec({ type: 'insertHyperlink', href: 'https://example.com', text: 'Guide' }).ok
  ).toBe(true);
  select(bob, 'See Docs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  // The new words follow the struck ones in a link of their own, in the link's style.
  expect(links(bob)).toEqual([
    ['', 'https://example.com', undefined],
    ['Guide', 'https://example.com', undefined],
  ]);
  expect(bob.surface!.session.bodyText()).toBe('See DocsGuide now!');
  expect(runStyles(bob)).toEqual(['Hyperlink', 'Hyperlink']);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toHaveLength(2);

  // A returning peer reads the proposal, and the reopened file keeps it.
  const rejoined = await harness.join(pair.alice, 'returning');
  rejoined.detach();
  const returning = mount({
    document: rejoined.room.document,
    modules: [reviewModule(), collaborationModule({ session: rejoined.room.session })],
  });
  expect(links(returning)).toEqual(links(bob));
  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );

  // Bob accepts: one link holds the new words, still in the style, on both screens.
  expect(bob.surface!.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Guide', 'https://example.com', undefined]]);
  expect(alice.surface!.session.bodyText()).toBe('See Guide now!');
  expect(runStyles(alice)).toEqual(['Hyperlink']);
});

test('deleting the link while new words for it are suggested converges', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(
    alice.exec({ type: 'insertHyperlink', href: 'https://example.com', text: 'Guide' }).ok
  ).toBe(true);
  // Bob removes the old link's words outright while the proposal is in flight.
  select(bob, 'See '.length, 'See Docs'.length);
  bob.surface!.deleteBackward();
  pair.resume();
  converge();
  // Alice's proposed words survive in their link; only the words Bob saw are gone.
  expect(bob.surface!.session.bodyText()).toBe('See Guide now');
  expect(links(bob).filter(([text]) => text !== '')).toEqual([
    ['Guide', 'https://example.com', undefined],
  ]);
});

test('a suggested link removal converges, undoes, redoes, reopens, rejects and accepts', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();

  // Alice proposes the removal while Bob types in the same paragraph.
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  select(bob, 'See Docs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  // The old link keeps only its struck words; the plain copy follows it, without the style.
  expect(links(bob)).toEqual([['', 'https://example.com', undefined]]);
  expect(bob.surface!.session.bodyText()).toBe('See DocsDocs now!');
  expect(runStyles(bob)).toEqual(['Hyperlink']);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['', 'https://example.com', undefined]]);

  // A returning peer reads the proposal, and the reopened file keeps it.
  const rejoined = await harness.join(pair.alice, 'returning');
  rejoined.detach();
  const returning = mount({
    document: rejoined.room.document,
    modules: [reviewModule(), collaborationModule({ session: rejoined.room.session })],
  });
  expect(links(returning)).toEqual(links(bob));
  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );

  // Bob rejects: the link comes back on both screens.
  expect(bob.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.com', undefined]]);

  // Proposed again and accepted: the words stay, with no link.
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([]);
  expect(alice.surface!.session.bodyText()).toBe('See Docs now!');
  expect(runStyles(alice)).toEqual([]);
});

test('an unlink while editing takes the style off on both screens, and undo gives it back', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();

  // Alice unlinks while Bob types into the link's paragraph.
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  select(bob, 'See Docs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([]);
  expect(runStyles(bob)).toEqual([]);
  expect(bob.surface!.session.bodyText()).toBe('See Docs now!');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(runStyles(bob)).toEqual(['Hyperlink']);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(runStyles(bob)).toEqual([]);

  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
});

test('deleting the link while its removal is suggested converges', async () => {
  const { pair, alice, bob, converge } = await peers();
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  // Bob removes the link's words outright while the proposal is in flight.
  select(bob, 'See '.length, 'See Docs'.length);
  bob.surface!.deleteBackward();
  pair.resume();
  converge();
  // Alice's plain copy survives; no linked words are left.
  expect(bob.surface!.session.bodyText()).toBe('See Docs now');
  expect(links(bob).filter(([text]) => text !== '')).toEqual([]);
});

test('a suggested new target inside a bookmark converges, reopens, rejects and accepts', async () => {
  // Word's cross-reference bookmark around the link: the proposal must stay inside it.
  const { pair, alice, bob, converge } = await peers(
    '<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r><w:bookmarkStart w:id="7" w:name="_Ref1"/>' +
      '<w:hyperlink w:anchor="old"><w:r><w:t>Docs</w:t></w:r></w:hyperlink>' +
      '<w:bookmarkEnd w:id="7"/><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:p>'
  );

  // Alice proposes a new target while Bob types in the same paragraph.
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  pair.pause();
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  select(bob, 'See Docs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([
    ['', '#old', undefined],
    ['Docs', 'https://example.org', undefined],
  ]);
  // The struck words and the copy are both in the bookmark.
  expect(bookmarkedText(bob)).toBe('DocsDocs');

  // The reopened file keeps the proposal and the bookmark.
  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(bookmarkedText(reopened)).toBe('DocsDocs');

  // Bob rejects: the old link comes back whole, in the bookmark, on both screens.
  expect(bob.surface!.session.applyTreeOps([{ op: 'rejectAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', '#old', undefined]]);
  expect(bookmarkedText(alice)).toBe('Docs');

  // Proposed again and accepted: the new link takes the old one's place in the bookmark.
  select(alice, 'See Do'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.org', undefined]]);
  expect(bookmarkedText(alice)).toBe('Docs');
  expect(alice.surface!.session.bodyText()).toBe('See Docs now!');
});

/** Both peers suggesting, with Bob's proposed link over "Docs". */
async function proposedLinkPeers() {
  const room = await peers();
  const { alice, bob, converge } = room;
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  expect(bob.setEditingMode('suggesting').ok).toBe(true);
  select(bob, 'See '.length, 'See Docs'.length);
  expect(bob.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.com', undefined]]);
  return room;
}

/** Reject the changes `author` proposed, and no other. */
function rejectChangesBy(editor: Editor, author: string) {
  const ops = collectReviewItems({ storyPart: editor.surface!.session.part() }).flatMap((item) =>
    item.kind === 'revision' && item.address.author === author
      ? item.addresses.map((revision) => ({ op: 'rejectRevision' as const, revision }))
      : []
  );
  expect(ops.length).toBeGreaterThan(0);
  expect(editor.surface!.session.applyTreeOps(ops).committed).toBe(true);
}

test('a new target for the other author’s proposed link converges, undoes, redoes, reopens, and is decided alone', async () => {
  const { pair, alice, bob, converge } = await proposedLinkPeers();

  // Alice proposes a new target for Bob's link while Bob types in the same paragraph.
  pair.pause();
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  select(bob, 'See DocsDocs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([
    ['', 'https://example.com', undefined],
    ['Docs', 'https://example.org', undefined],
  ]);
  expect(bob.surface!.session.bodyText()).toBe('See DocsDocsDocs now!');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toHaveLength(2);

  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );

  // Bob rejects Alice's change alone: his proposed link stands on both screens.
  rejectChangesBy(bob, 'Alice');
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.surface!.session.bodyText()).toBe('See DocsDocs now!');

  // Proposed again and all accepted: the new target remains.
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.org' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.org', undefined]]);
  expect(alice.surface!.session.bodyText()).toBe('See Docs now!');
});

test('removing the other author’s proposed link converges, undoes, redoes, reopens, and is decided alone', async () => {
  const { pair, alice, bob, converge } = await proposedLinkPeers();

  // Alice proposes the removal of Bob's link while Bob types in the same paragraph.
  pair.pause();
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  select(bob, 'See DocsDocs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([['', 'https://example.com', undefined]]);
  expect(bob.surface!.session.bodyText()).toBe('See DocsDocsDocs now!');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['', 'https://example.com', undefined]]);

  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );

  // Bob rejects Alice's removal alone: his proposed link stands on both screens.
  rejectChangesBy(bob, 'Alice');
  converge();
  expect(links(alice)).toEqual([['Docs', 'https://example.com', undefined]]);

  // Proposed again and all accepted: the words stay, with no link and no style.
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.applyTreeOps([{ op: 'acceptAllRevisions' }]).committed).toBe(true);
  converge();
  expect(links(alice)).toEqual([]);
  expect(runStyles(alice)).toEqual([]);
  expect(alice.surface!.session.bodyText()).toBe('See Docs now!');
});

test('the other author taking back a proposed link while its removal is suggested converges', async () => {
  const { pair, alice, bob, converge } = await proposedLinkPeers();
  pair.pause();
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  // Bob deletes his own proposed words while Alice's removal is in flight.
  select(bob, 'See Docs'.length, 'See DocsDocs'.length);
  bob.surface!.deleteBackward();
  pair.resume();
  converge();
  // Bob's proposed words are gone; Alice's plain copy survives beside his struck original.
  expect(bob.surface!.session.bodyText()).toBe('See DocsDocs now');
  expect(links(bob)).toEqual([]);
  expect(runStyles(bob)).toEqual([]);
});

test('taking off your own proposed link withdraws it, converges, undoes, redoes and reopens', async () => {
  const { pair, alice, bob, converge } = await peers();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();

  // Alice takes her own proposed link off while Bob types in the same paragraph.
  pair.pause();
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  select(bob, 'See DocsDocs now'.length);
  bob.surface!.type('!');
  pair.resume();
  converge();
  expect(links(bob)).toEqual([]);
  expect(runStyles(bob)).toEqual([]);
  expect(bob.surface!.session.bodyText()).toBe('See Docs now!');
  expect(serializeOoxmlPart(bob.surface!.session.part())).not.toMatch(/w:author="Alice"/);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([['Docs', 'https://example.com', undefined]]);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(links(bob)).toEqual([]);

  const reopened = mount({ document: new Uint8Array(await bob.save()) });
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
});

test('deleting the words while your own proposed link is withdrawn converges', async () => {
  const { pair, alice, bob, converge } = await peers();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  select(alice, 'See '.length, 'See Docs'.length);
  expect(alice.exec({ type: 'insertHyperlink', href: 'https://example.com' }).ok).toBe(true);
  converge();
  pair.pause();
  select(alice, 'See DocsDo'.length);
  expect(alice.exec({ type: 'removeHyperlink' }).ok).toBe(true);
  // Bob deletes the proposed copy outright while the withdrawal is in flight.
  select(bob, 'See Docs'.length, 'See DocsDocs'.length);
  bob.surface!.deleteBackward();
  pair.resume();
  converge();
  // Both took the copy away; the struck words stand again, with no link and no style.
  expect(bob.surface!.session.bodyText()).toBe('See Docs now');
  expect(links(bob)).toEqual([]);
  expect(runStyles(bob)).toEqual([]);
});
