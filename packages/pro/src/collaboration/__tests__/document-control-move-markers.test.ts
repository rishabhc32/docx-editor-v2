/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A tracked move whose source is inside a content control. The control's content holds the
// move's range markers, which `CT_SdtContentRun` allows as `EG_RangeMarkupElements`, so the
// control stays a typed control (format 1.4.4.2, canonicalModelVersion 2). Shared records keep
// the writer's node kinds, so every participant, a participant who joins later, a decision on
// the move, its undo, and the saved file see one typed control with the same content.

import { afterEach, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPackage,
} from '@docx-editor.dev/core/store';
import { reviewModule } from '../../review/review-module.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createDocumentCollaboration } from '../document-session.ts';
import { zipDocument } from './document-peer-support.ts';

const MARK = 'w:author="Alice" w:date="2026-10-01T10:00:00Z"';
/** "Old" moves out of a text control to the next paragraph; "Kept" stays in the control. */
const MOVED = zipDocument(
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '<w:text/></w:sdtPr><w:sdtContent>' +
    `<w:moveFromRangeStart w:id="1" ${MARK} w:name="move1"/>` +
    `<w:moveFrom w:id="2" ${MARK}><w:r><w:t>Old</w:t></w:r></w:moveFrom>` +
    '<w:moveFromRangeEnd w:id="1"/><w:r><w:t>Kept</w:t></w:r></w:sdtContent></w:sdt></w:p>' +
    `<w:p><w:moveToRangeStart w:id="3" ${MARK} w:name="move1"/>` +
    `<w:moveTo w:id="4" ${MARK}><w:r><w:t>Old</w:t></w:r></w:moveTo>` +
    '<w:moveToRangeEnd w:id="3"/></w:p>'
);

const open: { close(): void }[] = [];
afterEach(() => {
  for (const peer of open.splice(0).reverse()) peer.close();
});

async function peer(name: string, host?: { ydoc: Y.Doc }) {
  const ydoc = new Y.Doc();
  const awareness = new Awareness(ydoc);
  if (host) Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(host.ydoc));
  const room = await createDocumentCollaboration({
    ydoc,
    awareness,
    documentId: 'control-move-markers',
    identity: { actorId: name, name },
    bootstrap: host ? { kind: 'join' } : { kind: 'create', document: MOVED },
    offlineEditing: true,
  });
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: room.document,
    author: name,
    modules: [reviewModule(), collaborationModule({ session: room.session })],
  });
  const created = {
    ydoc,
    room,
    editor,
    close() {
      editor.destroy();
      container.remove();
      room.destroy();
      awareness.destroy();
      ydoc.destroy();
    },
  };
  open.push(created);
  return created;
}
type Peer = Awaited<ReturnType<typeof peer>>;

function sync(a: Peer, b: Peer) {
  a.room.session.flushPendingJournals();
  b.room.session.flushPendingJournals();
  Y.applyUpdate(b.ydoc, Y.encodeStateAsUpdate(a.ydoc, Y.encodeStateVector(b.ydoc)), 'relay');
  Y.applyUpdate(a.ydoc, Y.encodeStateAsUpdate(b.ydoc, Y.encodeStateVector(a.ydoc)), 'relay');
}

async function savedPart(peer: Peer) {
  const result = readOoxmlPackage(new Uint8Array(await peer.editor.save()));
  if (!result.ok) throw new Error(result.reason);
  return result.package.parts.get(result.package.mainDocumentPart)!;
}

async function converged(a: Peer, b: Peer) {
  expect(canonicalOoxmlFingerprint(await savedPart(a))).toBe(
    canonicalOoxmlFingerprint(await savedPart(b))
  );
}

/** The text of each typed control in the participant's editor. */
function controlTexts(peer: Peer): string[] {
  const part = peer.editor.surface!.session.part();
  return contentControlsIn(part.root).map(({ node }) => contentControlTextOf(node));
}

const decideAll = (action: 'accept' | 'reject') =>
  ({ type: 'resolveAllReviewChanges', action, scope: 'document' }) as const;

test('a control holding a move source stays typed on every participant and in the saved file', async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  await converged(alice, bob);
  expect(controlTexts(alice)).toEqual(['OldKept']);
  expect(controlTexts(bob)).toEqual(['OldKept']);

  const late = await peer('Late', alice);
  await converged(alice, late);
  expect(controlTexts(late)).toEqual(['OldKept']);

  const saved = await savedPart(bob);
  expect(contentControlsIn(saved.root)).toHaveLength(1);
});

test('accepting the move inside a control converges, undoes, and reopens with the control typed', async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlTexts(alice)).toEqual(['Kept']);
  expect(controlTexts(bob)).toEqual(['Kept']);

  expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlTexts(alice)).toEqual(['OldKept']);

  expect(bob.editor.exec({ type: 'redo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  const saved = await savedPart(alice);
  expect(contentControlsIn(saved.root).map(({ node }) => contentControlTextOf(node))).toEqual([
    'Kept',
  ]);
});
