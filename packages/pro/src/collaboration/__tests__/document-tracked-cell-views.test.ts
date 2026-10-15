/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A participant reading No Markup while another proposes deleting a column.
//
// The view is local: each participant lays the shared tree out in their own display mode. The
// proposal arrives as ordinary tracked markup, so the reader's No Markup must hide the column
// the moment it arrives, as Accept would, and show it again when the proposal is withdrawn.

import { afterEach, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { collaborationModule, reviewModule } from '../../index';
import { createPeerHarness, zipDocument, type Peer } from './document-peer-support';
import { packageFingerprint } from './document-support';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('tracked-cell-views', { offlineEditing: true });
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const p = (t: string) => `<w:p><w:r><w:t xml:space="preserve">${t}</w:t></w:r></w:p>`;
const cell = (t: string) => `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(t)}</w:tc>`;
// Paragraphs in order: before, a1, a2, a3, b1, b2, b3, after.
const BODY =
  p('before') +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>' +
  '<w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
  `<w:tr>${cell('a1')}${cell('a2')}${cell('a3')}</w:tr>` +
  `<w:tr>${cell('b1')}${cell('b2')}${cell('b3')}</w:tr></w:tbl>` +
  p('after');
const A2 = 2;

function mount(document: Uint8Array, author: string, peer: Peer): Editor {
  peer.detach();
  const container = window.document.createElement('div');
  window.document.body.append(container);
  containers.push(container);
  const editor = createDocxEditor({
    container,
    document,
    author,
    modules: [reviewModule(), collaborationModule({ session: peer.room.session })],
  });
  editors.push(editor);
  return editor;
}

/** The text of every laid-out cell, row by row, and how wide each cell is. */
function laidOutCells(editor: Editor) {
  const layout = editor.surface!.layout();
  const ids = editor.surface!.session.paragraphIds();
  const texts = ['before', 'a1', 'a2', 'a3', 'b1', 'b2', 'b3', 'after'];
  return layout.pages
    .flatMap((page) => page.fragments)
    .flatMap((block) =>
      block.kind === 'table'
        ? block.rows.map((row) =>
            row.cells.map((laid) => {
              const first = laid.blocks[0];
              const paragraph = first?.kind === 'paragraph' ? first.paragraphId : '';
              return { text: texts[ids.indexOf(paragraph)], width: laid.box.width };
            })
          )
        : []
    );
}

test('a column deletion another participant proposes leaves No Markup for the reader', async () => {
  const pair = await harness.pair(zipDocument(BODY));
  const alice = mount(pair.alice.room.document, 'Alice', pair.alice);
  const bob = mount(pair.bob.room.document, 'Bob', pair.bob);
  const converge = () => {
    pair.alice.room.session.flushPendingJournals();
    pair.bob.room.session.flushPendingJournals();
    expect(packageFingerprint(alice.surface!.session.currentPackage())).toBe(
      packageFingerprint(bob.surface!.session.currentPackage())
    );
  };
  bob.surface!.setRevisionDisplayMode('proposed');
  const before = laidOutCells(bob);
  expect(before.map((row) => row.map((laid) => laid.text))).toEqual([
    ['a1', 'a2', 'a3'],
    ['b1', 'b2', 'b3'],
  ]);

  expect(alice.exec({ type: 'setEditingMode', mode: 'suggesting' }).ok).toBe(true);
  const position = { paragraphId: alice.surface!.session.paragraphIds()[A2]!, offset: 0 };
  alice.surface!.setSelection({ anchor: position, head: position });
  expect(alice.exec({ type: 'deleteColumn' })).toEqual({ ok: true, changed: true });
  converge();

  // The column is gone from Bob's view and the cell before it took its width.
  const hidden = laidOutCells(bob);
  expect(hidden.map((row) => row.map((laid) => laid.text))).toEqual([
    ['a1', 'a3'],
    ['b1', 'b3'],
  ]);
  for (const [index, row] of hidden.entries()) {
    expect(row[0]!.width).toBeCloseTo(before[index]![0]!.width + before[index]![1]!.width, 3);
    expect(row[1]!.width).toBeCloseTo(before[index]![2]!.width, 3);
  }
  // Alice still sees the proposal in All Markup.
  expect(laidOutCells(alice).flat()).toHaveLength(6);

  // Withdrawn by undo, the column comes back for Bob.
  expect(alice.exec({ type: 'undo' }).ok).toBe(true);
  converge();
  expect(laidOutCells(bob)).toEqual(before);
});
