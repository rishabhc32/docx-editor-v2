/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Content control prompts in a shared room. Words typed over a prompt in suggesting mode replace
// it without a revision, and accepting them leaves the control holding them for every
// participant: no `w:showingPlcHdr`, and no `PlaceholderText` style on the words. Rejecting them
// shows the control's own prompt again. A file an older version saved, with the prompt struck
// and the words pending, ends the prompt on Accept too. An edit that empties a control without a
// declared type shows its prompt again, as for a rich text control. All of it holds after undo
// and redo, for a participant who joins later, and in the saved file.

import { afterEach, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import {
  canonicalOoxmlFingerprint,
  contentControlPropertiesOf,
  contentControlTextOf,
  contentControlsIn,
  readOoxmlPackage,
  serializeOoxmlPart,
} from '@docx-editor.dev/core/store';
import { reviewModule } from '../../review/review-module.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { createDocumentCollaboration } from '../document-session.ts';
import { zipDocument } from './document-peer-support.ts';

const PROMPT = 'Click here to enter text.';
const DATE = '2026-10-01T10:00:00Z';
const STYLED = '<w:rPr><w:rStyle w:val="PlaceholderText"/><w:b/></w:rPr>';
/** "Name: " and a text control holding `content`, then `tail`, and a paragraph after. */
const line = (content: string, tail: string) =>
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>' +
  `<w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>` +
  `${tail}</w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p>`;
const END = '<w:r><w:t xml:space="preserve"> end</w:t></w:r>';
const PROMPTED = zipDocument(line(`<w:r>${STYLED}<w:t>${PROMPT}</w:t></w:r>`, ''));
/** The prompt struck and Alice's words pending, as versions before the replacement saved them. */
const STRUCK = zipDocument(
  line(
    `<w:del w:id="1" w:author="Alice" w:date="${DATE}"><w:r>${STYLED}<w:delText>${PROMPT}</w:delText></w:r></w:del>` +
      `<w:ins w:id="2" w:author="Alice" w:date="${DATE}"><w:r>${STYLED}<w:t>Bob</w:t></w:r></w:ins>`,
    END
  )
);
const UNTYPED = zipDocument(
  '<w:p><w:r><w:t xml:space="preserve">Name: </w:t></w:r><w:sdt><w:sdtPr><w:id w:val="7"/>' +
    '</w:sdtPr><w:sdtContent><w:r><w:t>Hi</w:t></w:r></w:sdtContent></w:sdt>' +
    '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>'
);

const open: { close(): void }[] = [];
afterEach(() => {
  for (const peer of open.splice(0).reverse()) peer.close();
});

async function peer(name: string, host?: { ydoc: Y.Doc }, fixture = PROMPTED) {
  const ydoc = new Y.Doc();
  const awareness = new Awareness(ydoc);
  if (host) Y.applyUpdate(ydoc, Y.encodeStateAsUpdate(host.ydoc));
  const room = await createDocumentCollaboration({
    ydoc,
    awareness,
    documentId: 'control-prompts',
    identity: { actorId: name, name },
    bootstrap: host ? { kind: 'join' } : { kind: 'create', document: fixture },
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

/** The control's text and state as a participant's editor holds it. */
function controlOf(peer: Peer) {
  const part = peer.editor.surface!.session.part();
  const control = contentControlsIn(part.root)[0]?.node;
  const xml = /<w:sdt>.*<\/w:sdt>/s.exec(serializeOoxmlPart(part))?.[0] ?? '';
  return {
    text: control ? contentControlTextOf(control) : null,
    prompt: xml.includes('showingPlcHdr'),
    promptStyle: xml.includes('PlaceholderText'),
    pending: /<w:(ins|del) /.test(xml),
  };
}

const bodyText = (peer: Peer) => peer.editor.surface!.session.bodyText();

/** Alice selects the prompt whole and types over it in suggesting mode. */
function suggestOverPrompt(alice: Peer, text: string) {
  expect(alice.editor.setEditingMode('suggesting').ok).toBe(true);
  const surface = alice.editor.surface!;
  expect(surface.contentControls.navigate('next')).toBe(true);
  surface.type(text);
  expect(surface.state().lastRejection ?? null).toBeNull();
}

/** Put the caret at `offset` of the first paragraph and type `text`. */
function typeAt(peer: Peer, offset: number, text: string) {
  const surface = peer.editor.surface!;
  const position = { paragraphId: surface.session.paragraphIds()[0]!, offset };
  surface.setSelection({ anchor: position, head: position });
  surface.type(text);
  expect(surface.state().lastRejection ?? null).toBeNull();
}

const decideAll = (action: 'accept' | 'reject') =>
  ({ type: 'resolveAllReviewChanges', action, scope: 'document' }) as const;
const ACCEPTED = { text: 'Bob', prompt: false, promptStyle: false, pending: false };
/** The prompt replaced without a revision, and only the words proposed. */
const PROPOSED = { text: 'Bob', prompt: false, promptStyle: false, pending: true };
const SHOWN = { text: PROMPT, prompt: true, promptStyle: true, pending: false };

async function expectSavedWithoutPrompt(peer: Peer) {
  const xml = serializeOoxmlPart(await savedPart(peer));
  expect(xml).not.toContain('showingPlcHdr');
  expect(xml).not.toContain('PlaceholderText');
}

test('words typed over a prompt, accepted, converge, undo, redo, rejoin, and reopen', async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  suggestOverPrompt(alice, 'Bob');
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(bob)).toEqual(PROPOSED);
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(ACCEPTED);
  expect(controlOf(bob)).toEqual(ACCEPTED);
  expect(bodyText(alice)).toBe('Name: Bob\nAfter.');

  expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(PROPOSED);
  expect(bob.editor.exec({ type: 'redo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(ACCEPTED);

  const returning = await peer('Returning', bob);
  await converged(bob, returning);
  expect(controlOf(returning)).toEqual(ACCEPTED);
  expect(bodyText(returning)).toBe('Name: Bob\nAfter.');
  await expectSavedWithoutPrompt(bob);
});

test("words typed over a prompt, rejected, show the control's prompt on every peer", async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  suggestOverPrompt(alice, 'Bob');
  sync(alice, bob);
  expect(bob.editor.exec(decideAll('reject')).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(SHOWN);
  expect(controlOf(bob)).toEqual(SHOWN);
  expect(bodyText(alice)).toBe(`Name: ${PROMPT}\nAfter.`);

  expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(PROPOSED);
  expect(bob.editor.exec({ type: 'redo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(SHOWN);

  const returning = await peer('Returning', bob);
  await converged(bob, returning);
  expect(controlOf(returning)).toEqual(SHOWN);
});

test('accepting while the author keeps typing in the suggestion converges', async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  suggestOverPrompt(alice, 'Bob');
  sync(alice, bob);
  // Concurrently: Bob accepts, and Alice adds to her words before she sees the decision. The
  // prompt is gone, so "Bob" spans 6..9 in "Name: Bob", and offset 8 is inside the words.
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  typeAt(alice, 8, 'X');
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(controlOf(bob));
  expect(controlOf(bob)).toMatchObject({ text: 'BoXb', prompt: false, promptStyle: false });
});

test('accepting while another participant removes the control converges', async () => {
  const alice = await peer('Alice');
  const bob = await peer('Bob', alice);
  suggestOverPrompt(alice, 'Bob');
  sync(alice, bob);
  const control = contentControlsIn(alice.editor.surface!.session.part().root)[0]!.node;
  // Concurrently: Bob accepts, and Alice removes the whole control with its content.
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  const removed = alice.editor.surface!.session.applyTreeOps([
    { op: 'removeContentControl', controlId: control.id, keepContent: false },
  ]);
  expect(removed.committed).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(controlOf(bob));
  // The removal takes the control and its words on both peers.
  expect(controlOf(bob).text).toBeNull();
  expect(bodyText(bob)).toBe('Name: \nAfter.');
});

/** The struck prompt and Alice's pending words, as the older file holds them. */
const STRUCK_PENDING = { text: 'Bob', prompt: true, promptStyle: true, pending: true };

test('a struck prompt an older version saved ends on Accept, for every peer', async () => {
  const alice = await peer('Alice', undefined, STRUCK);
  const bob = await peer('Bob', alice, STRUCK);
  expect(controlOf(bob)).toEqual(STRUCK_PENDING);
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(ACCEPTED);
  expect(controlOf(bob)).toEqual(ACCEPTED);
  expect(bodyText(alice)).toBe('Name: Bob end\nAfter.');

  expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(STRUCK_PENDING);
  expect(bob.editor.exec({ type: 'redo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(ACCEPTED);

  const returning = await peer('Returning', bob, STRUCK);
  await converged(bob, returning);
  expect(controlOf(returning)).toEqual(ACCEPTED);
  await expectSavedWithoutPrompt(bob);
});

test('accepting a struck prompt while another participant removes the control converges', async () => {
  const alice = await peer('Alice', undefined, STRUCK);
  const bob = await peer('Bob', alice, STRUCK);
  const control = contentControlsIn(alice.editor.surface!.session.part().root)[0]!.node;
  // Concurrently: Bob accepts, and Alice removes the whole control with its content.
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  const removed = alice.editor.surface!.session.applyTreeOps([
    { op: 'removeContentControl', controlId: control.id, keepContent: false },
  ]);
  expect(removed.committed).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(controlOf(bob));
  expect(controlOf(bob).text).toBeNull();
  expect(bodyText(bob)).toBe('Name:  end\nAfter.');
});

test('accepting and rejecting a struck prompt at once converges on one outcome', async () => {
  const alice = await peer('Alice', undefined, STRUCK);
  const bob = await peer('Bob', alice, STRUCK);
  // Concurrently: Bob accepts the words, and Alice rejects them.
  expect(bob.editor.exec(decideAll('accept')).ok).toBe(true);
  expect(alice.editor.exec(decideAll('reject')).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(controlOf(bob));
  // Whichever decision the room keeps, the control is consistent: a prompt that shows only
  // the prompt, or words that are not a prompt.
  for (const { node } of contentControlsIn(bob.editor.surface!.session.part().root)) {
    const shown = contentControlTextOf(node);
    if (contentControlPropertiesOf(node).showingPlaceholder) expect(shown).toBe(PROMPT);
    else expect(shown).not.toContain(PROMPT);
  }
});

/** Select `[start, end)` of the first paragraph and delete it, in editing mode. */
function deleteRange(peer: Peer, start: number, end: number) {
  const surface = peer.editor.surface!;
  const paragraphId = surface.session.paragraphIds()[0]!;
  surface.setSelection({
    anchor: { paragraphId, offset: start },
    head: { paragraphId, offset: end },
  });
  surface.deleteBackward();
  expect(surface.state().lastRejection ?? null).toBeNull();
}

const FILLED = { text: 'Hi', prompt: false, promptStyle: false, pending: false };
/** The type's prompt, in the placeholder style the open gives a prompt. */
const TYPE_PROMPT = 'Click or tap here to enter text.';
const EMPTIED = { text: TYPE_PROMPT, prompt: true, promptStyle: true, pending: false };

test('emptying an untyped control shows its prompt on every peer, and reopens', async () => {
  const alice = await peer('Alice', undefined, UNTYPED);
  const bob = await peer('Bob', alice, UNTYPED);
  deleteRange(alice, 6, 8);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(EMPTIED);
  expect(controlOf(bob)).toEqual(EMPTIED);

  expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(bob)).toEqual(FILLED);
  expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true);
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(bob)).toEqual(EMPTIED);

  // The restored prompt is a prompt for the other participant too: typing replaces it.
  typeAt(bob, 8, 'Yo');
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual({ ...FILLED, text: 'Yo' });

  const returning = await peer('Returning', bob, UNTYPED);
  await converged(bob, returning);
  expect(controlOf(returning)).toEqual({ ...FILLED, text: 'Yo' });
});

test('emptying an untyped control while another participant types in it converges', async () => {
  const alice = await peer('Alice', undefined, UNTYPED);
  const bob = await peer('Bob', alice, UNTYPED);
  // Concurrently: Alice deletes "Hi", and Bob types between its two letters.
  deleteRange(alice, 6, 8);
  typeAt(bob, 7, 'X');
  sync(alice, bob);
  await converged(alice, bob);
  expect(controlOf(alice)).toEqual(controlOf(bob));
  // Neither edit is lost, and every control that says it shows its prompt holds only that.
  expect(bob.editor.surface!.session.bodyText()).toContain('X');
  for (const { node } of contentControlsIn(bob.editor.surface!.session.part().root)) {
    if (contentControlPropertiesOf(node).showingPlaceholder) {
      expect(contentControlTextOf(node)).toBe(TYPE_PROMPT);
    }
  }
});
