/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Two people suggesting in content controls of one shared document.
//
// Alice presses a prompt and types over it while Bob types in a block control. Each keeps one
// suggestion growing as they type, each caret stays after the typed words, and Bob sees
// Alice's caret where hers is. Bob then types at the end of Alice's value, which stays inside
// the control. Both screens agree after every step, through undo, a deletion racing typing, a
// reconnect, and a save.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { pressContentControl, typeKeys } from '../../__tests__/surface-gestures-support';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('prompt-caret', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const mounted: { editor: Editor; container: HTMLElement }[] = [];
afterEach(() => {
  for (const { editor, container } of mounted.splice(0)) {
    editor.destroy();
    container.remove();
  }
  harness.cleanup();
});

const PROMPT = 'Click here to enter your full name';
const LABEL = 'Name: ';
const BODY =
  `<w:p><w:r><w:t xml:space="preserve">${LABEL}</w:t></w:r><w:sdt><w:sdtPr><w:showingPlcHdr/>` +
  '<w:text/></w:sdtPr><w:sdtContent><w:r><w:rPr><w:rStyle w:val="PlaceholderText"/></w:rPr>' +
  `<w:t>${PROMPT}</w:t></w:r></w:sdtContent></w:sdt>` +
  '<w:r><w:t xml:space="preserve"> end.</w:t></w:r></w:p>' +
  '<w:sdt><w:sdtPr><w:richText/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Clause.</w:t></w:r></w:p>' +
  '</w:sdtContent></w:sdt>';

function mount(document: Uint8Array, author: string, peer?: Peer): Editor {
  peer?.detach();
  const container = window.document.createElement('div');
  window.document.body.append(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    mode: 'suggesting',
    modules: [
      reviewModule(),
      ...(peer ? [collaborationModule({ session: peer.room.session })] : []),
    ],
  });
  mounted.push({ editor, container });
  return editor;
}

const containerOf = (editor: Editor) =>
  mounted.find((candidate) => candidate.editor === editor)!.container;
const keys = (editor: Editor, text: string) => typeKeys(editor, containerOf(editor), text);
const paragraphId = (editor: Editor, index: number) =>
  editor.surface!.session.paragraphIds()[index]!;
function caret(editor: Editor, index: number, offset: number) {
  const position = { paragraphId: paragraphId(editor, index), offset };
  editor.surface!.setSelection({ anchor: position, head: position });
}
const head = (editor: Editor) => {
  const { anchor, head } = editor.surface!.state().selection;
  expect(anchor).toEqual(head);
  return head;
};
/** What each control holds, ids, dates and paragraph tags removed. */
const controls = (editor: Editor) =>
  [
    ...serializeOoxmlPart(editor.surface!.session.part())
      .replace(/ w:(date|id)="[^"]*"/g, '')
      .matchAll(/<w:sdtContent>(.*?)<\/w:sdtContent>/gs),
  ].map((match) => match[1]!.replace(/<\/?w:p\b[^>]*>/g, ''));
const ins = (author: string, text: string) =>
  `<w:ins w:author="${author}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:ins>`;
const plainIns = (author: string, text: string) =>
  `<w:ins w:author="${author}"><w:r><w:t>${text}</w:t></w:r></w:ins>`;
const revisions = (editor: Editor) =>
  editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [`${item.revisionKind}:${item.author}`] : []));

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
    const left = alice.surface!.session.currentPackage();
    const right = bob.surface!.session.currentPackage();
    expect(packageFingerprint(right)).toBe(packageFingerprint(left));
    expect(saveReopenDigest(right)).toEqual(saveReopenDigest(left));
  };
  const publishAlice = () =>
    applyAwarenessUpdate(
      pair.bob.awareness,
      encodeAwarenessUpdate(pair.alice.awareness, [pair.alice.awareness.clientID]),
      'test'
    );
  return { pair, alice, bob, flush, converge, publishAlice };
}

test('typing over a prompt beside a peer keeps each caret after its words, and one suggestion each', async () => {
  const { pair, alice, bob, flush, converge, publishAlice } = await peers();

  // Concurrently: Alice presses the prompt and types over it; Bob types in the block control.
  pair.pause();
  pressContentControl(alice, containerOf(alice));
  expect(alice.surface!.state().selection.head.offset).toBe(LABEL.length + PROMPT.length);
  keys(alice, 'Bob');
  caret(bob, 1, 'Clause.'.length);
  keys(bob, ' Two');
  flush();
  pair.resume();
  converge();
  expect(head(alice)).toEqual({ paragraphId: paragraphId(alice, 0), offset: 9 });
  expect(head(bob)).toEqual({ paragraphId: paragraphId(bob, 1), offset: 11 });

  // The next keys continue after the words on both screens.
  keys(alice, 'by');
  converge();
  expect(head(alice).offset).toBe(11);
  expect(bob.surface!.session.bodyText()).toBe('Name: Bobby end.\nClause. Two');
  expect(controls(bob)).toEqual([
    plainIns('Alice', 'Bobby'),
    `<w:r><w:t>Clause.</w:t></w:r>${ins('Bob', ' Two')}`,
  ]);
  expect(revisions(bob)).toEqual(['insert:Alice', 'insert:Bob']);

  // Bob sees Alice's caret where hers is, after "Bobby".
  publishAlice();
  const remote = pair.bob.room.session.remoteSelections()[0]!;
  expect([remote.head.nodeId, remote.head.offset]).toEqual([paragraphId(bob, 0), 11]);
  expect(
    caretAt(bob.surface!.layout(), { paragraphId: remote.head.nodeId, offset: remote.head.offset })
  ).toEqual(caretAt(alice.surface!.layout(), head(alice)));
  expect(containerOf(bob).querySelector('.docx-remote-caret')).not.toBeNull();

  // Bob, a second author, types at the end of Alice's value: inside the control, after hers.
  caret(bob, 0, 11);
  keys(bob, ' Jr');
  converge();
  expect(head(bob).offset).toBe(14);
  expect(controls(alice)[0]).toBe(`${plainIns('Alice', 'Bobby')}${ins('Bob', ' Jr')}`);
  expect(alice.surface!.session.bodyText()).toBe('Name: Bobby Jr end.\nClause. Two');

  // Undo and redo Alice's words.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).not.toContain('by');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(bob.surface!.session.bodyText()).toBe('Name: Bobby Jr end.\nClause. Two');

  // A reconnecting participant and a saved copy read the same document.
  const rejoined = await harness.join(pair.alice, 'returning');
  const returning = mount(rejoined.room.document, 'Cy', rejoined);
  expect(returning.surface!.session.bodyText()).toBe('Name: Bobby Jr end.\nClause. Two');
  const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(revisions(reopened)).toEqual(['insert:Alice', 'insert:Bob', 'insert:Bob']);
});

test('striking the block text while the peer types at its end converges and undoes', async () => {
  // Agreement and the undo are asserted, not who struck what. Bob's keys extend the run that
  // Alice's concurrent deletion strikes, and the merge strikes them with it, as it does in a
  // paragraph outside any control.
  const { pair, alice, bob, flush, converge } = await peers();
  caret(bob, 1, 'Clause.'.length);
  keys(bob, ' Two');
  converge();

  pair.pause();
  const block = paragraphId(alice, 1);
  alice.surface!.setSelection({
    anchor: { paragraphId: block, offset: 0 },
    head: { paragraphId: block, offset: 'Clause. Two'.length },
  });
  alice.surface!.deleteBackward();
  keys(bob, ' more');
  flush();
  pair.resume();
  converge();
  expect(bob.surface!.session.bodyText()).toContain('Clause. Two more');
  expect(revisions(bob)).toContain('insert:Bob');

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(controls(bob)[1]).toBe(`<w:r><w:t>Clause.</w:t></w:r>${ins('Bob', ' Two more')}`);
});
