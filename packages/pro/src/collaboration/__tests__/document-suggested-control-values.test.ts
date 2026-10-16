/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A content-control value one participant suggests replicates as a suggestion.
//
// The pick writes the old display text in a `w:del` and the new one in a `w:ins` inside the
// control, and moves the control's own value record untracked. The peer must see both cards;
// a rejection there must converge, a rejected checkbox must agree on `w14:checked` everywhere,
// and races with the peer's typing and with the peer's own pick must converge and undo.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('suggested-control-values', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const sdt = (properties: string, content: string) =>
  `<w:p><w:r><w:t xml:space="preserve">Term: </w:t></w:r><w:sdt><w:sdtPr>${properties}</w:sdtPr>` +
  `<w:sdtContent>${content}</w:sdtContent></w:sdt></w:p>`;
const DROPDOWN = sdt(
  '<w:dropDownList w:lastValue="1"><w:listItem w:displayText="One" w:value="1"/>' +
    '<w:listItem w:displayText="Two" w:value="2"/><w:listItem w:displayText="Three" w:value="3"/>' +
    '</w:dropDownList>',
  '<w:r><w:t>One</w:t></w:r>'
);
const CHECKBOX = sdt(
  `<w14:checkbox xmlns:w14="${W14}"><w14:checked w14:val="0"/>` +
    '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>' +
    '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/></w14:checkbox>',
  '<w:r><w:sym w:font="MS Gothic" w:char="2610"/></w:r>'
);

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

async function peers(body: string) {
  const pair = await harness.pair(zipDocument(body));
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

const xml = (editor: Editor) => serializeOoxmlPart(editor.surface!.session.part());
const controlId = (editor: Editor) => editor.surface!.layout().contentControls![0]!.id;
const revisions = (editor: Editor) =>
  editor
    .getReviewItems()
    .flatMap((item) => (item.kind === 'revision' ? [`${item.revisionKind}:${item.author}`] : []));

test('a suggested pick converges, reopens, and a peer rejection converges', async () => {
  const { alice, bob, converge } = await peers(DROPDOWN);
  expect(alice.surface!.contentControls.setValue(controlId(alice), '2')).toBe(true);
  converge();

  expect(xml(bob)).toMatch(/<w:del [^>]*w:author="Alice"[^>]*><w:r><w:delText>One</);
  expect(xml(bob)).toMatch(/<w:ins [^>]*w:author="Alice"[^>]*><w:r><w:t>Two</);
  expect(xml(bob)).toContain('w:lastValue="2"');
  expect(revisions(bob)).toEqual(['replace:Alice']);

  const reopened = mount(new Uint8Array(await bob.save()), 'Reader');
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(bob.surface!.session.currentPackage())
  );
  expect(revisions(reopened)).toEqual(['replace:Alice']);

  for (const card of bob.getReviewItems().filter((item) => item.kind === 'revision')) {
    expect(bob.rejectReviewItem(card.key).ok).toBe(true);
  }
  converge();
  for (const editor of [alice, bob]) {
    expect(xml(editor)).not.toContain('<w:del ');
    expect(xml(editor)).not.toContain('<w:ins ');
    expect(editor.surface!.session.bodyText()).toBe('Term: One');
    expect(revisions(editor)).toEqual([]);
  }
});

test('a rejected checkbox toggle puts w14:checked back on every replica', async () => {
  const { alice, bob, converge } = await peers(CHECKBOX);
  expect(alice.surface!.contentControls.setValue(controlId(alice), 'true')).toBe(true);
  converge();
  expect(xml(bob)).toContain('<w14:checked w14:val="1"/>');
  expect(revisions(bob)).toEqual(['replace:Alice']);

  for (const card of bob.getReviewItems().filter((item) => item.kind === 'revision')) {
    expect(bob.rejectReviewItem(card.key).ok).toBe(true);
  }
  converge();
  for (const editor of [alice, bob]) {
    expect(xml(editor)).toContain('<w14:checked w14:val="0"/>');
    expect(xml(editor)).toContain('w:char="2610"');
    expect(xml(editor)).not.toContain('w:char="2612"');
  }
});

test('a suggested pick racing the peer typing beside the control converges and undoes', async () => {
  const { pair, alice, bob, converge } = await peers(DROPDOWN);
  pair.pause();
  expect(alice.surface!.contentControls.setValue(controlId(alice), '2')).toBe(true);
  const paragraphId = bob.surface!.session.paragraphIds()[0]!;
  bob.surface!.setSelection({
    anchor: { paragraphId, offset: 4 },
    head: { paragraphId, offset: 4 },
  });
  bob.surface!.type('s');
  bob.surface!.flushPendingInput();
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  expect(bob.surface!.session.bodyText()).toBe('Terms: OneTwo');
  expect(bob.query({ type: 'contentControls' })).toHaveLength(1);
  expect(revisions(bob)).toEqual(['replace:Alice']);

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toMatch(/w:author="Alice"/);
  expect(bob.surface!.session.bodyText()).toBe('Terms: One');
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(revisions(bob)).toEqual(['replace:Alice']);
});

test('a suggested pick racing the peer’s own pick on the same control converges', async () => {
  // Only agreement is asserted. A value write moves the control's `w:sdtPr`, and a peer's
  // concurrent edit inside the same control can split it in two on every replica; an editing
  // mode pick does the same.
  const { pair, alice, bob, converge } = await peers(DROPDOWN);
  pair.pause();
  expect(alice.surface!.contentControls.setValue(controlId(alice), '2')).toBe(true);
  expect(bob.surface!.contentControls.setValue(controlId(bob), '3')).toBe(true);
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();

  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).not.toMatch(/w:author="Alice"/);
  expect(alice.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(xml(bob)).toMatch(/w:author="Alice"/);
});
