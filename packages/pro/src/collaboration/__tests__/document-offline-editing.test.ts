/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Offline editing: a session created with `offlineEditing` keeps admitting local edits while
// the transport is `disconnected`, and the buffered updates merge on reconnect like any
// concurrent online edit. The default stays refuse-while-disconnected.

import { afterEach, describe, expect, test } from 'bun:test';
import type { TreeDocOp } from '@docx-editor.dev/core/store';
import { createPeerHarness, zipDocument } from './document-peer-support.ts';
import { collaborationDocx } from './support.ts';

const proseBytes = (): Uint8Array => collaborationDocx();
const tableBytes = (): Uint8Array =>
  zipDocument(
    '<w:p><w:r><w:t>before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>' +
      '</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r>' +
      '<w:t>a1</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>'
  );

const offlineHarness = createPeerHarness('offline-editing', { offlineEditing: true });
const defaultHarness = createPeerHarness('offline-editing-default');
const pausedHarness = createPeerHarness('offline-editing-paused', { offlineEditing: false });

afterEach(() => {
  offlineHarness.cleanup();
  defaultHarness.cleanup();
  pausedHarness.cleanup();
});

function insertTextOp(paragraphId: string, offset: number, text: string): TreeDocOp {
  return { op: 'insertText', paragraphId, offset, text };
}

describe('offline editing', () => {
  test('a disconnected replica keeps editing and converges on reconnect', async () => {
    const { alice, bob, pause, resume } = await offlineHarness.pair(proseBytes());

    pause();
    alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test-drop');
    expect(alice.room.session.statusSnapshot().status).toBe('disconnected');

    // Both sides type while the wire is down: alice offline, bob still 'ready'.
    offlineHarness.apply(alice, [
      insertTextOp(offlineHarness.paragraphIdAt(alice, 0), 0, 'offline '),
    ]);
    offlineHarness.apply(bob, [insertTextOp(offlineHarness.paragraphIdAt(bob, 1), 0, 'online ')]);

    alice.room.session.setTransportStatus('ready');
    resume();

    offlineHarness.expectConverged(alice, bob);
  });

  test('undo still works for edits made while disconnected', async () => {
    const { alice, bob, pause, resume } = await offlineHarness.pair(proseBytes());
    pause();
    alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test-drop');
    offlineHarness.apply(alice, [
      insertTextOp(offlineHarness.paragraphIdAt(alice, 0), 0, 'undo-me '),
    ]);
    expect(alice.room.session.canUndo()).toBe(true);
    expect(alice.room.session.undo()).toBe(true);
    alice.room.session.setTransportStatus('ready');
    resume();
    offlineHarness.expectConverged(alice, bob);
  });

  test('error and destroyed stay refused with offline editing on', async () => {
    const { alice } = await offlineHarness.pair(proseBytes());
    alice.room.session.setTransportStatus('error', 'transport', 'test-terminal');
    expect(
      alice.room.session.gateOperations(
        [insertTextOp(offlineHarness.paragraphIdAt(alice, 0), 0, 'x')],
        { kind: 'body' }
      )
    ).toBe('collaboration-session-not-ready');
  });

  test('the default keeps admitting edits while disconnected, as Yjs editors do', async () => {
    const { alice } = await defaultHarness.pair(proseBytes());
    alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test-drop');
    expect(
      alice.room.session.gateOperations(
        [insertTextOp(defaultHarness.paragraphIdAt(alice, 0), 0, 'x')],
        { kind: 'body' }
      )
    ).toBeNull();
  });

  test('a suggested table format waits for the room while disconnected', async () => {
    const { alice } = await offlineHarness.pair(tableBytes());
    const tableId = offlineHarness
      .packageOf(alice)
      .parts.get('/word/document.xml')!
      .root.children.flatMap((body) => (body.kind === 'textValue' ? [] : body.children))
      .find((node) => node.kind === 'table')!.id;
    const center = (revision?: { readonly author: string }): TreeDocOp[] => [
      {
        op: 'setTableProperties',
        tableId,
        justification: 'center',
        ...(revision ? { revision } : {}),
      } as TreeDocOp,
    ];
    const gate = (ops: readonly TreeDocOp[]) =>
      alice.room.session.gateOperations(ops, { kind: 'body' });
    expect(gate(center({ author: 'alice' }))).toBeNull();
    alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test-drop');
    expect(gate(center({ author: 'alice' }))).toBe('collaboration-session-not-ready');
    // The same format, untracked, and a suggested text edit still go through.
    expect(gate(center())).toBeNull();
    expect(
      gate([
        {
          ...insertTextOp(offlineHarness.paragraphIdAt(alice, 0), 0, 'x'),
          revision: { author: 'alice' },
        } as TreeDocOp,
      ])
    ).toBeNull();
    alice.room.session.setTransportStatus('ready');
    expect(gate(center({ author: 'alice' }))).toBeNull();
  });

  test('offlineEditing: false refuses edits while disconnected', async () => {
    const { alice } = await pausedHarness.pair(proseBytes());
    alice.room.session.setTransportStatus('disconnected', 'transport-disconnected', 'test-drop');
    expect(
      alice.room.session.gateOperations(
        [insertTextOp(pausedHarness.paragraphIdAt(alice, 0), 0, 'x')],
        { kind: 'body' }
      )
    ).toBe('collaboration-session-not-ready');
  });
});
