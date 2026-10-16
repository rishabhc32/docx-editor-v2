/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Review decisions over a collaborator's suggestions. A node a collaboration room creates has the
// room's logical id (`lid:…`), which names no part, so the review commands must find each
// change's story some other way. Accept All, Reject All, single decisions, author changes and
// review navigation work on the peer that received a suggestion and on its author once the room
// renamed its nodes. Every decision converges, undoes, and survives save and reopen.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { revisionSiteNodeIdsOf } from '../../../../core/src/store/store/review-items.ts';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, R, REL, W, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('review-all-on-peers', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const BODY =
  '<w:p><w:r><w:t xml:space="preserve">Payment is due in 30 days.</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>Delivery is free.</w:t></w:r></w:p>';
const WITH_HEADER = zipDocument(
  `${BODY}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>`,
  {
    overrides: `<Override PartName="/word/header1.xml" ContentType="${HEADER_TYPE}"/>`,
    documentRels: `<Relationships xmlns="${REL}"><Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/></Relationships>`,
    extraXml: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Head</w:t></w:r></w:p></w:hdr>`,
    },
  }
);
const WITHOUT_HEADER = zipDocument(`${BODY}<w:sectPr/>`);

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

function select(editor: Editor, paragraph: number, from: number, to = from): void {
  const paragraphId = editor.surface!.session.paragraphIds()[paragraph]!;
  editor.surface!.setSelection({
    anchor: { paragraphId, offset: from },
    head: { paragraphId, offset: to },
  });
}

const lastParagraph = (editor: Editor) => editor.surface!.session.paragraphIds().length - 1;

/** One suggestion of each kind, made through the editor in suggesting mode. */
const SUGGESTIONS = {
  text(editor: Editor) {
    select(editor, 0, 0);
    editor.surface!.type('Net ');
  },
  paragraph(editor: Editor) {
    select(editor, 0, editor.surface!.session.bodyText().indexOf('\n'));
    editor.surface!.splitParagraph();
    editor.surface!.type('Interest accrues.');
  },
  format(editor: Editor) {
    select(editor, lastParagraph(editor), 0, 'Delivery'.length);
    expect(editor.exec({ type: 'toggleMark', mark: 'bold' }).ok).toBe(true);
  },
  header(editor: Editor) {
    expect(editor.exec({ type: 'editHeaderFooter', position: 'header' }).ok).toBe(true);
    editor.surface!.type('Draft ');
    expect(editor.exec({ type: 'exitHeaderFooter' }).ok).toBe(true);
  },
};
type Suggestion = keyof typeof SUGGESTIONS;

const pending = (editor: Editor) =>
  editor.surface!.session.reviewItems().filter((item) => item.kind === 'revision');
/** Whether every pending change is addressed only by logical ids, which name no part. */
const logical = (editor: Editor) =>
  pending(editor).every((item) => revisionSiteNodeIdsOf(item).every((id) => id.startsWith('lid:')));
const authors = (editor: Editor) =>
  pending(editor).flatMap((item) => (item.kind === 'revision' ? [item.author] : []));

async function room(bytes = WITH_HEADER) {
  const pair = await harness.pair(bytes);
  const alice = mount(pair.alice.room.document, 'Alice', pair.alice);
  const bob = mount(pair.bob.room.document, 'Bob', pair.bob);
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  /** Relays pending edits, requires one document on both replicas, and answers it. */
  const converged = (): string => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    const left = alice.surface!.session.currentPackage();
    const right = bob.surface!.session.currentPackage();
    expect(packageFingerprint(right)).toBe(packageFingerprint(left));
    expect(saveReopenDigest(right)).toEqual(saveReopenDigest(left));
    return packageFingerprint(left);
  };
  return { pair, alice, bob, converged };
}

for (const suggestion of Object.keys(SUGGESTIONS) as Suggestion[]) {
  for (const action of ['accept', 'reject'] as const) {
    test(`the peer can ${action} all of a ${suggestion} suggestion, as the document scope does`, async () => {
      const { alice, bob, converged } = await room();
      SUGGESTIONS[suggestion](alice);
      const suggested = converged();
      expect(pending(bob).length).toBeGreaterThan(0);
      expect(logical(bob)).toBe(true);

      expect(bob.exec({ type: 'resolveAllReviewChanges', action })).toMatchObject({
        ok: true,
        changed: true,
        revisions: { skipped: [], remaining: 0 },
      });
      const decided = converged();
      expect(decided).not.toBe(suggested);
      expect(pending(alice)).toEqual([]);

      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged()).toBe(suggested);
      expect(bob.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' }).ok).toBe(
        true
      );
      expect(converged()).toBe(decided);

      const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
      expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
        saveReopenDigest(bob.surface!.session.currentPackage())
      );
      expect(pending(reopened)).toEqual([]);
    });
  }
}

test('the author decides her own suggestions after the room renamed their nodes', async () => {
  const { alice, bob, converged } = await room();
  SUGGESTIONS.text(alice);
  SUGGESTIONS.header(alice);
  converged();
  // Bob's edit reaches Alice as shared state, which renames her new nodes to logical ids.
  select(bob, 1, 0);
  bob.surface!.type('Fast ');
  const suggested = converged();
  expect(logical(alice)).toBe(true);

  expect(alice.exec({ type: 'setReviewChangesAuthor', author: 'Ada' })).toMatchObject({
    ok: true,
    changed: true,
  });
  expect(authors(bob)).toEqual(['Ada', 'Ada']);
  expect(alice.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
  converged();
  expect(bob.surface!.session.bodyText()).toBe(
    'Payment is due in 30 days.\nFast Delivery is free.'
  );
  expect(pending(bob)).toEqual([]);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  expect(converged()).toBe(suggested);
  expect(authors(bob)).toEqual(['Alice', 'Alice']);
});

test('the peer changes the author of every suggestion, then decides each one', async () => {
  const { pair, alice, bob, converged } = await room();
  SUGGESTIONS.text(alice);
  SUGGESTIONS.format(alice);
  SUGGESTIONS.header(alice);
  converged();
  expect(logical(bob)).toBe(true);

  const changed = bob.exec({ type: 'setReviewChangesAuthor', author: 'Counsel' });
  expect(changed).toMatchObject({ ok: true, changed: true });
  expect(changed.revisionAuthors?.updated).toHaveLength(3);
  const reattributed = converged();
  expect(authors(alice)).toEqual(['Counsel', 'Counsel', 'Counsel']);

  // One change by key, one accepted and one rejected from its card.
  const placements = bob.getReviewItems();
  const keyOf = (text: string, kind = 'insert') =>
    placements.find(
      ({ item }) => item.kind === 'revision' && item.revisionKind === kind && item.text === text
    )!.key;
  expect(
    bob.exec({ type: 'resolveAllReviewChanges', action: 'reject', keys: [keyOf('Draft ')] })
  ).toMatchObject({ ok: true, changed: true, revisions: { remaining: 2 } });
  expect(bob.acceptReviewItem(keyOf('Net ')).ok).toBe(true);
  expect(bob.rejectReviewItem(keyOf('', 'format')).ok).toBe(true);
  const decided = converged();
  expect(pending(alice)).toEqual([]);
  expect(alice.surface!.session.bodyText()).toBe(
    'Net Payment is due in 30 days.\nDelivery is free.'
  );

  for (let step = 0; step < 3; step += 1) expect(bob.exec({ type: 'undo' }).ok).toBe(true);
  expect(converged()).toBe(reattributed);
  for (let step = 0; step < 3; step += 1) expect(bob.exec({ type: 'redo' }).ok).toBe(true);
  expect(converged()).toBe(decided);

  // A replica that joins later reads the same decided document.
  const carol = await harness.join(pair.alice, 'carol');
  const rejoined = mount(carol.room.document, 'Carol', carol);
  expect(packageFingerprint(rejoined.surface!.session.currentPackage())).toBe(decided);
});

test('Accept All reaches a header the room created during the session', async () => {
  const { alice, bob, converged } = await room(WITHOUT_HEADER);
  SUGGESTIONS.header(alice);
  select(bob, 1, 0);
  bob.surface!.type('Fast ');
  const suggested = converged();
  // Shared state renamed the whole new part, its root included.
  const header = alice.surface!.session.storyParts()[1]!;
  expect(header.root.id.startsWith('lid:')).toBe(true);
  expect(logical(alice)).toBe(true);

  expect(
    alice.exec({ type: 'resolveAllReviewChanges', action: 'accept', scope: 'document' })
  ).toMatchObject({ ok: true, changed: true, revisions: { skipped: [], remaining: 0 } });
  const accepted = converged();
  expect(pending(bob)).toEqual([]);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  expect(converged()).toBe(suggested);
  expect(alice.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
  expect(converged()).not.toBe(accepted);
  expect(pending(bob)).toEqual([]);
});

test('review navigation starts from a caret in a paragraph the room created', async () => {
  const { alice, bob, converged } = await room(WITHOUT_HEADER);
  expect(alice.setEditingMode('editing').ok).toBe(true);
  select(alice, 1, 'Delivery '.length);
  alice.surface!.splitParagraph();
  expect(alice.setEditingMode('suggesting').ok).toBe(true);
  select(alice, 0, 0);
  alice.surface!.type('one ');
  select(alice, 2, 'is '.length);
  alice.surface!.type('two ');
  converged();
  expect(bob.surface!.session.paragraphIds()[2]!.startsWith('lid:')).toBe(true);

  const landed = (direction: 'next' | 'previous') => {
    select(bob, 2, 0);
    expect(bob.exec({ type: 'navigateReviewChange', direction }).ok).toBe(true);
    const active = bob.getReviewItems().find((placement) => placement.isActive);
    return active?.item.kind === 'revision' ? active.item.text : undefined;
  };
  expect(landed('next')).toBe('two ');
  expect(landed('previous')).toBe('one ');
});

test('Reject All racing the author typing inside her suggestion converges', async () => {
  const { pair, alice, bob, converged } = await room();
  SUGGESTIONS.text(alice);
  SUGGESTIONS.header(alice);
  converged();

  pair.pause();
  expect(bob.exec({ type: 'resolveAllReviewChanges', action: 'reject' }).ok).toBe(true);
  select(alice, 0, 'Net'.length);
  alice.surface!.type('t');
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converged();
  // The rejection removes the change Bob saw; the letter Alice typed meanwhile stays proposed.
  expect(alice.surface!.session.bodyText()).toBe('tPayment is due in 30 days.\nDelivery is free.');
  expect(pending(alice).map((item) => item.kind === 'revision' && item.text)).toEqual(['t']);

  expect(bob.exec({ type: 'undo' }).ok).toBe(true);
  converged();
  expect(alice.surface!.session.bodyText()).toBe(
    'Nett Payment is due in 30 days.\nDelivery is free.'
  );
});

test('a single editor resolves its own suggestions in every story as before', async () => {
  const editor = mount(WITH_HEADER, 'Alice');
  expect(editor.setEditingMode('suggesting').ok).toBe(true);
  for (const suggest of Object.values(SUGGESTIONS)) suggest(editor);
  const ids = pending(editor).flatMap((item) => [...revisionSiteNodeIdsOf(item)]);
  expect(ids.some((id) => id.startsWith('lid:'))).toBe(false);
  const fingerprint = () => packageFingerprint(editor.surface!.session.currentPackage());
  const suggested = fingerprint();

  // After the bold word, so no change is active and the walk starts from the caret.
  select(editor, lastParagraph(editor), 'Delivery is'.length);
  expect(editor.exec({ type: 'navigateReviewChange', direction: 'previous' }).ok).toBe(true);
  const active = editor.getReviewItems().find((placement) => placement.isActive)!;
  expect(active.item.kind === 'revision' && active.item.revisionKind).toBe('format');

  const changed = editor.exec({ type: 'setReviewChangesAuthor', author: 'Counsel' });
  expect(changed.revisionAuthors?.updated).toHaveLength(pending(editor).length);
  const resolved = editor.exec({ type: 'resolveAllReviewChanges', action: 'accept' });
  expect(resolved).toMatchObject({ ok: true, revisions: { skipped: [], remaining: 0 } });
  const decided = fingerprint();
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(editor.exec({ type: 'undo' }).ok).toBe(true);
  expect(fingerprint()).toBe(suggested);
  expect(
    editor.exec({ type: 'resolveAllReviewChanges', action: 'accept', scope: 'document' }).ok
  ).toBe(true);
  expect(fingerprint()).toBe(decided);
});
