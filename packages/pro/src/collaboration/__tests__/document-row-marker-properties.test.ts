/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Resolving a tracked row drops the `w:trPr` that held only its marker. Every replica must see
// the same rows, keep the other row properties and a peer's concurrent typing, converge when a
// peer writes into the removed `w:trPr`, and undo the whole decision.

import { afterEach, expect, test } from 'bun:test';
import {
  planRevisionBatch,
  revisionItemsOf,
  serializeOoxmlPart,
  type OoxmlNode,
} from '@docx-editor.dev/core/store';
import {
  createPeerHarness,
  nodeText,
  walk,
  zipDocument,
  type Peer,
} from './document-peer-support.ts';

const peers = createPeerHarness('row-marker-properties-room');
afterEach(() => peers.cleanup());

const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (properties: string, text: string) => `<w:tr>${properties}${cell(text)}</w:tr>`;
const BODY =
  '<w:tbl><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>' +
  row('<w:trPr><w:ins w:id="1" w:author="Ada"/></w:trPr>', 'inserted') +
  row('<w:trPr><w:del w:id="2" w:author="Ada"/></w:trPr>', 'deleted') +
  row('<w:trPr><w:trHeight w:val="500"/><w:del w:id="3" w:author="Ada"/></w:trPr>', 'tall') +
  row('<w:trPr><w:cantSplit/><w:ins w:id="4" w:author="Ada"/></w:trPr>', 'unsplit') +
  row('<w:trPr/>', 'authored') +
  '</w:tbl><w:p/>';
const KEPT = {
  reject: [
    row('', 'deleted'),
    row('<w:trPr><w:trHeight w:val="500"/></w:trPr>', 'tall'),
    row('<w:trPr/>', 'authored'),
  ].join(''),
  accept: [
    row('', 'inserted'),
    row('<w:trPr><w:cantSplit/></w:trPr>', 'unsplit'),
    row('<w:trPr/>', 'authored'),
  ].join(''),
};
/** The row each action keeps whose `w:trPr` held only its marker. */
const BARE = { reject: 'deleted', accept: 'inserted' } as const;

function rowsOf(peer: Peer): string {
  // The room seed gives every paragraph an identity, which the expected rows leave out.
  const xml = serializeOoxmlPart(peer.store.bodyStore().part).replace(/ w14:\w+Id="\w+"/g, '');
  return xml.slice(xml.indexOf('</w:tblGrid>') + '</w:tblGrid>'.length, xml.indexOf('</w:tbl>'));
}

function rowHolding(peer: Peer, text: string): { row: OoxmlNode; paragraph: OoxmlNode } {
  let found: { row: OoxmlNode; paragraph: OoxmlNode } | undefined;
  walk(peer.store.bodyStore().part.root, (node) => {
    if (node.kind !== 'tableRow' || nodeText(node) !== text) return;
    walk(node, (inner) => {
      if (inner.kind === 'paragraph') found ??= { row: node, paragraph: inner };
    });
  });
  if (!found) throw new Error(`no row holds ${text}`);
  return found;
}

function tableOf(peer: Peer): OoxmlNode {
  let table: OoxmlNode | undefined;
  walk(peer.store.bodyStore().part.root, (node) => {
    if (node.kind === 'table') table ??= node;
  });
  if (!table) throw new Error('no table');
  return table;
}

for (const action of ['accept', 'reject'] as const) {
  test(`${action}: every replica drops the emptied row properties, and undo restores them`, async () => {
    const { alice, bob } = await peers.pair(zipDocument(BODY));
    const before = rowsOf(bob);
    const batch = planRevisionBatch(alice.store.bodyStore().part, action);
    expect(batch.result.skipped).toEqual([]);
    peers.apply(alice, batch.ops);
    peers.expectConverged(alice, bob);
    expect(rowsOf(bob)).toBe(KEPT[action]);
    expect(revisionItemsOf(bob.store.bodyStore().part)).toHaveLength(0);
    const carol = await peers.join(alice, 'carol');
    peers.expectConverged(alice, carol);
    expect(rowsOf(carol)).toBe(KEPT[action]);

    expect(alice.room.session.undo()).toBe(true);
    peers.expectConverged(alice, bob);
    expect(rowsOf(bob)).toBe(before);
    expect(alice.room.session.redo()).toBe(true);
    peers.expectConverged(alice, bob);
    expect(rowsOf(bob)).toBe(KEPT[action]);
    const reconnected = await peers.remount(bob);
    peers.expectConverged(alice, reconnected);
    expect(rowsOf(reconnected)).toBe(KEPT[action]);
  });

  test(`${action}: a peer typing in the row while it is resolved keeps the typing`, async () => {
    const { alice, bob, pause, resume } = await peers.pair(zipDocument(BODY));
    const { paragraph } = rowHolding(bob, BARE[action]);
    pause();
    peers.apply(alice, planRevisionBatch(alice.store.bodyStore().part, action).ops);
    peers.apply(bob, [{ op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'typed ' }]);
    resume();
    peers.expectConverged(alice, bob);
    expect(rowsOf(bob)).toBe(KEPT[action].replace(`>${BARE[action]}<`, `>typed ${BARE[action]}<`));
  });

  test(`${action}: a peer sizing the row while it is resolved converges on the removal`, async () => {
    const { alice, bob, pause, resume } = await peers.pair(zipDocument(BODY));
    const { row: target } = rowHolding(bob, BARE[action]);
    pause();
    peers.apply(alice, planRevisionBatch(alice.store.bodyStore().part, action).ops);
    peers.apply(bob, [
      { op: 'setTableRowHeight', tableId: tableOf(bob).id, rowId: target.id, heightTwips: 720 },
    ]);
    resume();
    peers.expectConverged(alice, bob);
    expect(revisionItemsOf(bob.store.bodyStore().part)).toHaveLength(0);
    // The height lands in the `w:trPr` the decision removed, so the removal wins. A `w:pPr` that
    // one participant clears while another sets a property in it resolves the same way.
    expect(rowsOf(bob)).toBe(KEPT[action]);
  });
}
