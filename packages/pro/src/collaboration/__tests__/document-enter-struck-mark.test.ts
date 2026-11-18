/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// ENTER INSIDE A PARAGRAPH WHOSE MARK ANOTHER PARTICIPANT PROPOSED REMOVING.
//
// Ben proposes joining two paragraphs, which strikes the first paragraph's mark. Cy then
// presses Enter inside that paragraph. Word inserts Cy's mark BEFORE the struck one, so the
// head ends with Cy's insertion and the tail keeps Ben's deletion. The split copied the strike
// onto the head, so Accept all undid Cy's Enter together with Ben's join, on every replica.
import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { paragraphTextOf, type OoxmlNode } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('enter-struck-mark', { offlineEditing: true });
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

const plain = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function caret(editor: Editor, index: number, offset: number): void {
  const paragraphId = editor.surface!.session.paragraphIds()[index]!;
  editor.surface!.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
}

/** Each paragraph's text and own mark revisions, as `text|kind:author,…`. */
function paragraphs(editor: Editor): string[] {
  const session = editor.surface!.session;
  const named = (node: OoxmlNode | undefined, localName: string) =>
    node && node.kind !== 'textValue'
      ? node.children.find((child) => child.kind !== 'textValue' && child.localName === localName)
      : undefined;
  const out: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind !== 'paragraph') return node.children.forEach(visit);
    const rPr = named(named(node, 'pPr'), 'rPr');
    const marks = (rPr && rPr.kind !== 'textValue' ? rPr.children : []).flatMap((mark) =>
      mark.kind !== 'textValue' && (mark.localName === 'ins' || mark.localName === 'del')
        ? [`${mark.localName}:${mark.attributes.find((a) => a.localName === 'author')?.value}`]
        : []
    );
    out.push(`${paragraphTextOf(session.part(), node.id) ?? ''}|${marks.join(',')}`);
  };
  visit(session.part().root);
  return out;
}

async function open() {
  const pair = await harness.pair(zipDocument(plain('Heading Level 3') + plain('Next paragraph')));
  const ben = mount(pair.alice, 'Ben');
  const cy = mount(pair.bob, 'Cy');
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(ben.surface!.session.currentPackage())).toBe(
      packageFingerprint(cy.surface!.session.currentPackage())
    );
    expect(paragraphs(ben)).toEqual(paragraphs(cy));
    return paragraphs(ben);
  };
  return { pair, ben, cy, converge };
}

test("Enter keeps another participant's proposed join on the original mark, on both replicas", async () => {
  const { ben, cy, converge } = await open();
  caret(ben, 1, 0);
  ben.surface!.deleteBackward();
  expect(converge()).toEqual(['Heading Level 3|del:Ben', 'Next paragraph|']);

  caret(cy, 0, 1);
  cy.surface!.splitParagraph();
  const proposed = ['H|ins:Cy', 'eading Level 3|del:Ben', 'Next paragraph|'];
  expect(converge()).toEqual(proposed);

  // Undo and redo of the Enter replicate, and Ben's proposal is untouched by either.
  expect(cy.exec({ type: 'undo' }).ok).toBe(true);
  expect(converge()).toEqual(['Heading Level 3|del:Ben', 'Next paragraph|']);
  expect(cy.exec({ type: 'redo' }).ok).toBe(true);
  expect(converge()).toEqual(proposed);

  // Accept all keeps Cy's break and takes Ben's join; reject all restores the original.
  const resolve = (editor: Editor, op: 'acceptAllRevisions' | 'rejectAllRevisions') =>
    editor.surface!.session.applyTreeOps([{ op }]).committed;
  expect(resolve(ben, 'acceptAllRevisions')).toBe(true);
  expect(converge()).toEqual(['H|', 'eading Level 3Next paragraph|']);
  expect(ben.exec({ type: 'undo' }).ok).toBe(true);
  expect(converge()).toEqual(proposed);
  expect(resolve(cy, 'rejectAllRevisions')).toBe(true);
  expect(converge()).toEqual(['Heading Level 3|', 'Next paragraph|']);
  expect(cy.exec({ type: 'undo' }).ok).toBe(true);
  expect(converge()).toEqual(proposed);

  // The proposals survive save and reopen as they stand on the replicas.
  const saved = new Uint8Array(await ben.save());
  const reopened = createDocxEditor({ document: saved, modules: [reviewModule()] });
  editors.push(reopened);
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  reopened.attach(container);
  expect(paragraphs(reopened)).toEqual(proposed);
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(ben.surface!.session.currentPackage())
  );
});

test('a concurrent join and Enter both survive the reconnect and converge', async () => {
  // Offline, Cy splits the paragraph before Ben's strike reaches Cy's replica, so the strike
  // follows the paragraph identity the head keeps rather than the mark the tail keeps. What is
  // pinned here is that neither proposal is lost and both replicas agree.
  const { pair, ben, cy, converge } = await open();
  pair.pause();
  caret(ben, 1, 0);
  ben.surface!.deleteBackward();
  caret(cy, 0, 1);
  cy.surface!.splitParagraph();
  pair.resume();
  const merged = converge();
  expect(merged.map((paragraph) => paragraph.split('|')[0])).toEqual([
    'H',
    'eading Level 3',
    'Next paragraph',
  ]);
  const marks = merged.flatMap((paragraph) => paragraph.split('|')[1]!.split(',')).sort();
  expect(marks.filter(Boolean)).toEqual(['del:Ben', 'ins:Cy']);
});
