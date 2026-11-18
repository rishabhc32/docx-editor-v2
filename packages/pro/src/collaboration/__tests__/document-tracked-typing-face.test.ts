/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// One reviewer types beside another reviewer's strike, in a collaboration room. The typed
// words take the face of the text the caret continues — bold, size and font included, struck
// words counted — on both peers, through undo, redo, a concurrent strike, a concurrent reject,
// reconnect, and save and reopen.
import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart } from '@docx-editor.dev/core/store';
import { DocxEditor } from '@docx-editor.dev/editor-api/browser';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('tracked-typing-face', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const FACE = '<w:rPr><w:b/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>';
const run = (text: string) => `<w:r>${FACE}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const BODY = `<w:p>${run('Alpha beta gamma')}</w:p><w:p>${run('Delta epsilon')}</w:p>`;

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
    reviewDisplayMode: 'all-markup',
    modules: [reviewModule(), collaborationModule({ session: peer.room.session })],
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

/** Every run Ben's insertions hold, as `face|text`, in document order. */
function bensRuns(editor: Editor): string[] {
  const out: string[] = [];
  for (const [, inner] of xml(editor).matchAll(/<w:ins [^>]*w:author="Ben"[^>]*>(.*?)<\/w:ins>/g))
    for (const [, face, text] of inner!.matchAll(/<w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t[^>]*>([^<]*)/g))
      out.push(`${face === FACE ? 'face' : (face ?? 'none')}|${text}`);
  return out;
}

test("typing beside another reviewer's strike keeps the face on both peers", async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const ana = mount(pair.alice, 'Ana');
  const ben = mount(pair.bob, 'Ben');
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(ana.surface!.session.currentPackage())).toBe(
      packageFingerprint(ben.surface!.session.currentPackage())
    );
    expect(bensRuns(ana)).toEqual(bensRuns(ben));
  };

  // Ana strikes the end of the first paragraph and the start of the second.
  select(ana, 0, 11, 16);
  ana.surface!.deleteBackward();
  select(ana, 1, 0, 5);
  ana.surface!.deleteBackward();
  converge();
  expect(xml(ben).match(/<w:del [^>]*w:author="Ana"/g)).toHaveLength(2);

  // Ben types inside her first strike, which ends its paragraph: no unstruck run follows.
  select(ben, 0, 13);
  ben.surface!.type('X');
  converge();
  expect(bensRuns(ana)).toEqual(['face|X']);

  // Undo and redo move the typed words, and the face comes back with them.
  expect(ben.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(bensRuns(ana)).toEqual([]);
  expect(ben.exec({ type: 'redo' }).ok).toBe(true);
  converge();
  expect(bensRuns(ana)).toEqual(['face|X']);

  // Offline, Ana strikes the rest of the second paragraph while Ben types in front of her
  // strike at its start. They reconnect.
  pair.pause();
  select(ana, 1, 5, 'Delta epsilon'.length);
  ana.surface!.deleteBackward();
  select(ben, 1, 0);
  ben.surface!.type('Y');
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  expect(bensRuns(ana)).toEqual(['face|X', 'face|Y']);

  // Offline, Ana rejects all her strikes while Ben types inside the one that now ends the
  // second paragraph. They reconnect.
  pair.pause();
  const anaApi = DocxEditor.createBrowser(ana, { author: 'Ana' });
  await anaApi.run(async (context) => {
    context.document.body.revisions.rejectAll();
    await context.sync();
  });
  anaApi.dispose();
  select(ben, 1, 'YDelta'.length);
  ben.surface!.type('W');
  pair.alice.room.session.flushPendingJournals();
  pair.bob.room.session.flushPendingJournals();
  pair.resume();
  converge();
  // Reject all takes Ben's earlier proposals too; the one he typed meanwhile stays.
  expect(xml(ana)).not.toContain('<w:del ');
  expect(bensRuns(ana)).toEqual(['face|W']);
  expect(ana.surface!.session.bodyText()).toContain('Delta epsilonW');

  // Saved and reopened, the face is still there.
  const saved = new Uint8Array(await ben.save());
  const reopened = createDocxEditor({ document: saved, modules: [reviewModule()] });
  editors.push(reopened);
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  reopened.attach(container);
  expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
    saveReopenDigest(ben.surface!.session.currentPackage())
  );
  expect(bensRuns(reopened)).toEqual(bensRuns(ben));
});
