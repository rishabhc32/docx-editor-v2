/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A suggested deletion that covers an equation, in a collaboration room. The equation is one
// unit of paragraph text, and the strike now covers it as `w:del` around the `m:oMath` or
// `m:oMathPara`. Both editors must see that strike, through concurrent typing beside it, a
// direct deletion of it by the other participant, undo, redo, accept, reject, a late joiner,
// and save and reopen.
import { afterEach, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { serializeOoxmlPart, type OoxmlNode } from '@docx-editor.dev/core/store';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint, saveReopenDigest } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const harness = createPeerHarness('tracked-equation-deletion', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const INLINE = `<m:oMath xmlns:m="${M}"><m:r><m:t>x=1</m:t></m:r></m:oMath>`;
const DISPLAY =
  `<m:oMathPara xmlns:m="${M}"><m:oMath><m:r><m:t>y=2</m:t></m:r></m:oMath>` + '</m:oMathPara>';
/** "Alpha ", the equation at offset 6, then " Omega." */
const document_ = (equation: string) =>
  zipDocument(`<w:p>${run('Alpha ')}${equation}${run(' Omega.')}</w:p>`);

function mount(peer: Peer, author: string, mode: 'suggesting' | 'edit' = 'suggesting') {
  peer.detach();
  const container = document.createElement('div');
  document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document: peer.room.document,
    author,
    mode,
    reviewDisplayMode: 'all-markup',
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

/** Math elements inside a `w:del`, and math elements outside one, by local name. */
function math(editor: Editor): { struck: string[]; kept: string[] } {
  const found = { struck: [] as string[], kept: [] as string[] };
  const walk = (node: OoxmlNode, struck: boolean): void => {
    if (node.kind === 'textValue') return;
    if (node.namespaceUri === M && /^oMath(Para)?$/.test(node.localName)) {
      (struck ? found.struck : found.kept).push(node.localName);
      return;
    }
    for (const child of node.children) walk(child, struck || node.kind === 'revisionDelete');
  };
  walk(editor.surface!.session.part().root, false);
  return found;
}

async function room(equation: string, bobMode: 'suggesting' | 'edit' = 'suggesting') {
  const pair = await harness.pair(document_(equation));
  const alice = mount(pair.alice, 'Alice');
  const bob = mount(pair.bob, 'Bob', bobMode);
  const sync = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
  };
  const converge = () => {
    sync();
    for (const peer of [pair.alice, pair.bob]) {
      expect(peer.room.session.statusSnapshot().status).toBe('ready');
    }
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  return { pair, alice, bob, sync, converge };
}

const cases = [
  { name: 'an inline equation', equation: INLINE, element: 'oMath' },
  { name: 'a display equation', equation: DISPLAY, element: 'oMathPara' },
] as const;

for (const { name, equation, element } of cases) {
  describe(`a suggested deletion over ${name}`, () => {
    test('strikes it on both editors through undo, redo, accept, reject, and rejoin', async () => {
      const { pair, alice, bob, converge } = await room(equation);
      select(alice, 3, 10);
      alice.surface!.deleteBackward();
      expect(alice.surface!.state().lastRejection).toBeNull();
      converge();
      expect(math(bob)).toEqual({ struck: [element], kept: [] });
      expect(xml(bob)).toMatch(/<w:del [^>]*w:author="Alice"/);
      expect(xml(bob).match(/<w:del /g)).toHaveLength(1);

      expect(alice.exec({ type: 'undo' }).ok).toBe(true);
      converge();
      expect(math(bob)).toEqual({ struck: [], kept: [element] });
      expect(xml(bob)).not.toMatch(/<w:del /);
      expect(alice.exec({ type: 'redo' }).ok).toBe(true);
      converge();
      expect(math(bob)).toEqual({ struck: [element], kept: [] });

      // A late joiner reads the same strike, and save and reopen keeps it.
      const rejoined = mount(await harness.join(pair.alice, 'rejoined'), 'Returning');
      expect(packageFingerprint(rejoined.surface!.session.currentPackage())).toBe(
        packageFingerprint(alice.surface!.session.currentPackage())
      );
      expect(math(rejoined)).toEqual({ struck: [element], kept: [] });
      const reopened = createDocxEditor({
        document: new Uint8Array(await bob.save()),
        modules: [reviewModule()],
      });
      editors.push(reopened);
      const container = document.createElement('div');
      document.body.append(container);
      containers.push(container);
      reopened.attach(container);
      expect(math(reopened)).toEqual({ struck: [element], kept: [] });
      expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
        saveReopenDigest(alice.surface!.session.currentPackage())
      );

      // The other participant decides it. Accept removes the equation with the words.
      const decide = (action: 'accept' | 'reject') =>
        bob.exec({ type: 'resolveAllReviewChanges', action, scope: 'document' });
      expect(decide('accept')).toMatchObject({ ok: true });
      converge();
      expect(math(alice)).toEqual({ struck: [], kept: [] });
      expect(bob.exec({ type: 'undo' }).ok).toBe(true);
      converge();
      expect(math(alice)).toEqual({ struck: [element], kept: [] });
      // Reject keeps the equation, unmarked.
      expect(decide('reject')).toMatchObject({ ok: true });
      converge();
      expect(math(alice)).toEqual({ struck: [], kept: [element] });
      expect(xml(alice)).not.toMatch(/<w:del /);
    });

    test('converges with concurrent typing beside it on the other editor', async () => {
      const { pair, alice, bob, sync, converge } = await room(equation);
      pair.pause();
      select(alice, 6, 7);
      alice.surface!.deleteBackward();
      select(bob, 0);
      bob.surface!.type('B ');
      select(bob, 9);
      bob.surface!.type('!');
      sync();
      pair.resume();
      converge();
      expect(math(alice)).toEqual({ struck: [element], kept: [] });
      expect(xml(alice)).toMatch(/<w:ins [^>]*w:author="Bob"/);
    });

    test('converges when the other editor deletes the equation outright', async () => {
      const { pair, alice, bob, sync, converge } = await room(equation, 'edit');
      pair.pause();
      select(alice, 3, 10);
      alice.surface!.deleteBackward();
      select(bob, 6, 7);
      bob.surface!.deleteBackward();
      sync();
      pair.resume();
      converge();
      // The direct deletion removes it; nothing keeps an unmarked copy.
      expect(math(alice).kept).toEqual([]);
    });
  });
}
