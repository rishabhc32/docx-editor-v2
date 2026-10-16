/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Concurrent writes of a table row's or cell's properties.
//
// Two people sizing one row, shading one cell or dragging one column divider at the same
// moment each write the row's `w:trPr`, the cell's `w:tcPr`, or a child of one. Merged, the
// row or cell held two of an element its schema allows once. Validation accepted it, because
// the tree keeps those containers generic, but later edits of that table refused with
// `duplicate-property-container` or left the other copy behind, and the saved document was
// not valid for Word.
//
// Every replica now shows the first live copy in shared child order, a properties element
// ahead of the cells or paragraphs it leads, and a winning container with the properties its
// hidden copies hold. These cases pin that the room stays `ready`, converges to one valid
// document, saves and reopens as shown, keeps taking table edits, and undoes on both sides,
// through shared stores and through two mounted editors.

import { afterEach, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createDocxEditor } from '@docx-editor.dev/core/editor';
import { caretAt } from '@docx-editor.dev/core/layout';
import { type OoxmlElement, type OoxmlNode, type TreeDocOp } from '@docx-editor.dev/core/store';
import { tableColumnDividerResizeTargetFrom } from '../../../../core/src/layout/table-interaction-targets.ts';
import { collaborationModule } from '../collaboration-module.ts';
import { readCollaborationDocument } from '../document-read.ts';
import {
  createPeerHarness,
  nodeText,
  walk,
  W,
  zipDocument,
  type Peer,
} from './document-peer-support.ts';
import {
  loadPackage,
  packageDigest,
  packageFingerprint,
  saveReopenDigest,
} from './document-support.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const harness = createPeerHarness('concurrent-table-properties-room');
type Editor = ReturnType<typeof createDocxEditor>;
const editors: Editor[] = [];
const containers: HTMLElement[] = [];

afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy();
  for (const container of containers.splice(0)) container.remove();
  harness.cleanup();
});

const p = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
const TC_W = '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>';

/** A two-by-two table whose cells start with `cellProperties`, its first row `rowProperties`. */
function table(cellProperties = '', rowProperties = ''): Uint8Array {
  const cell = (text: string) => `<w:tc>${cellProperties}${p(text)}</w:tc>`;
  return zipDocument(
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      `<w:tr>${rowProperties}${cell('A')}${cell('B')}</w:tr>` +
      `<w:tr>${cell('C')}${cell('D')}</w:tr>` +
      `</w:tbl>${p('After')}<w:sectPr/>`
  );
}

/** The main story of a shared store, or of a mounted editor. */
function rootOf(source: Peer | Editor): OoxmlElement {
  return 'room' in source
    ? source.store.bodyStore().part.root
    : source.surface!.session.part().root;
}

function elements(
  source: Peer | Editor,
  kind: OoxmlNode['kind'],
  localName?: string
): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  walk(rootOf(source), (node) => {
    if (node.kind === 'textValue' || node.kind !== kind) return;
    if (localName === undefined || node.localName === localName) found.push(node);
  });
  return found;
}

function ids(peer: Peer) {
  const [tableNode] = elements(peer, 'table');
  return {
    table: tableNode!.id,
    rows: elements(peer, 'tableRow').map((node) => node.id),
    cells: elements(peer, 'tableCell').map((node) => node.id),
    columns: elements(peer, 'generic', 'gridCol').map((node) => node.id),
  };
}

function wmlChildren(node: OoxmlElement, localName: string): OoxmlElement[] {
  const found: OoxmlElement[] = [];
  for (const child of node.children) {
    if (child.kind !== 'textValue' && child.namespaceUri === W && child.localName === localName)
      found.push(child);
  }
  return found;
}

function attribute(node: OoxmlElement, localName: string): string | null {
  return node.attributes.find((entry) => entry.localName === localName)?.value ?? null;
}

/** `w:val` of every `w:trHeight` in the row's properties, in shown order. */
function rowHeights(source: Peer | Editor, row = 0): string[] {
  const properties = wmlChildren(elements(source, 'tableRow')[row]!, 'trPr');
  return properties.flatMap((trPr) =>
    wmlChildren(trPr, 'trHeight').map((height) => attribute(height, 'val')!)
  );
}

/** Each named `w:tcPr` child of one cell, as its attribute values. */
function cellProperties(
  source: Peer | Editor,
  cell: number,
  localName: string
): Record<string, string>[] {
  const properties = wmlChildren(elements(source, 'tableCell')[cell]!, 'tcPr');
  return properties.flatMap((tcPr) =>
    wmlChildren(tcPr, localName).map((child) =>
      Object.fromEntries(child.attributes.map((entry) => [entry.localName, entry.value]))
    )
  );
}

/**
 * Every row and cell holds at most one properties element, ahead of the cells or paragraphs
 * it leads, and no properties element holds two of one child.
 */
function expectValidTables(source: Peer | Editor): void {
  const leads = (node: OoxmlElement, localName: string, content: readonly string[]) => {
    const at = node.children.findIndex(
      (child) => child.kind !== 'textValue' && child.localName === localName
    );
    const first = node.children.findIndex((child) => content.includes(child.kind));
    if (at >= 0 && first >= 0) expect(at).toBeLessThan(first);
  };
  const unique = (node: OoxmlElement) => {
    const names = node.children.flatMap((child) =>
      child.kind !== 'textValue' && child.namespaceUri === W ? [child.localName] : []
    );
    expect(new Set(names).size).toBe(names.length);
  };
  walk(rootOf(source), (node) => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'tableRow') {
      expect(wmlChildren(node, 'trPr').length).toBeLessThanOrEqual(1);
      expect(wmlChildren(node, 'tblPrEx').length).toBeLessThanOrEqual(1);
      leads(node, 'trPr', ['tableCell']);
    }
    if (node.kind === 'tableCell') {
      expect(wmlChildren(node, 'tcPr').length).toBeLessThanOrEqual(1);
      leads(node, 'tcPr', ['paragraph', 'table']);
    }
    if (node.kind === 'generic' && ['trPr', 'tcPr', 'tcBorders'].includes(node.localName)) {
      unique(node);
    }
  });
}

/** Ready, converged, valid, and saved as shown, on every replica. */
function expectHealthy(...peers: Peer[]): void {
  for (const peer of peers) {
    expect(peer.room.session.status()).toBe('ready');
    expectValidTables(peer);
    const saved = loadPackage(readCollaborationDocument(peer.ydoc));
    expect(packageDigest(saved)).toEqual(saveReopenDigest(harness.packageOf(peer)));
  }
  for (const peer of peers.slice(1)) harness.expectConverged(peers[0]!, peer);
}

function rowHeight(peer: Peer, heightTwips: number, row = 0): TreeDocOp {
  const table = ids(peer);
  return { op: 'setTableRowHeight', tableId: table.table, rowId: table.rows[row]!, heightTwips };
}

function fill(peer: Peer, value: string, cell = 0): TreeDocOp {
  const table = ids(peer);
  return {
    op: 'setTableCellFill',
    tableId: table.table,
    cellIds: [table.cells[cell]!],
    color: { kind: 'hex', value },
  };
}

function verticalAlignment(peer: Peer, alignment: 'top' | 'center' | 'bottom'): TreeDocOp {
  const table = ids(peer);
  return {
    op: 'setTableCellVerticalAlignment',
    tableId: table.table,
    cellIds: [table.cells[0]!],
    alignment,
  };
}

function columnWidths(peer: Peer, left: number, right: number): TreeDocOp {
  const table = ids(peer);
  return {
    op: 'setTableColumnWidths',
    tableId: table.table,
    leftGridColumnId: table.columns[0]!,
    rightGridColumnId: table.columns[1]!,
    leftWidthTwips: left,
    rightWidthTwips: right,
  };
}

/** Each peer's edit, applied while neither sees the other's. */
async function concurrently(
  bytes: Uint8Array,
  aliceOps: (peer: Peer) => TreeDocOp[],
  bobOps: (peer: Peer) => TreeDocOp[]
): Promise<{ alice: Peer; bob: Peer }> {
  const { alice, bob, pause, resume } = await harness.pair(bytes);
  pause();
  harness.apply(alice, aliceOps(alice));
  harness.apply(bob, bobOps(bob));
  resume();
  return { alice, bob };
}

describe('concurrent table properties', () => {
  test('two heights for a row without w:trPr show one, and the table takes edits', async () => {
    const { alice, bob } = await concurrently(
      table(),
      (peer) => [rowHeight(peer, 400)],
      (peer) => [rowHeight(peer, 800)]
    );
    expectHealthy(alice, bob);
    const shown = rowHeights(alice);
    expect(shown).toHaveLength(1);
    expect(['400', '800']).toContain(shown[0]!);
    const late = await harness.join(alice, 'late');
    expectHealthy(alice, bob, late);
    // Duplicates refused every later table edit with `duplicate-property-container`.
    const loser = shown[0] === '400' ? bob : alice;
    harness.apply(loser, [rowHeight(loser, 600)]);
    harness.apply(loser, [fill(loser, '00FF00', 1)]);
    expectHealthy(alice, bob, late);
    expect(rowHeights(late)).toEqual(['600']);
    expect(cellProperties(late, 1, 'shd')).toEqual([{ val: 'clear', fill: '00FF00' }]);
  });

  test('two row heights for a row that has one keep one w:trHeight', async () => {
    const { alice, bob } = await concurrently(
      table('', '<w:trPr><w:trHeight w:val="300"/></w:trPr>'),
      (peer) => [rowHeight(peer, 400)],
      (peer) => [rowHeight(peer, 800)]
    );
    expectHealthy(alice, bob);
    expect(rowHeights(bob)).toHaveLength(1);
    // A later height replaced only the first copy and left the other value in the file.
    harness.apply(alice, [rowHeight(alice, 600)]);
    expectHealthy(alice, bob);
    expect(rowHeights(bob)).toEqual(['600']);
  });

  test('each author undoes a concurrent row height, and redoes it', async () => {
    const { alice, bob } = await concurrently(
      table(),
      (peer) => [rowHeight(peer, 400)],
      (peer) => [rowHeight(peer, 800)]
    );
    expectHealthy(alice, bob);
    expect(rowHeights(alice)).toHaveLength(1);
    expect(alice.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(rowHeights(bob)).toHaveLength(1);
    expect(bob.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(rowHeights(alice)).toEqual([]);
    expect(alice.room.session.redo()).toBe(true);
    expectHealthy(alice, bob);
    expect(rowHeights(bob)).toEqual(['400']);
  });

  test('two fills of one cell show one w:shd, and the cell keeps taking fills', async () => {
    const { alice, bob } = await concurrently(
      table(TC_W),
      (peer) => [fill(peer, 'FF0000')],
      (peer) => [fill(peer, '0000FF')]
    );
    expectHealthy(alice, bob);
    const [shown] = cellProperties(alice, 0, 'shd');
    expect(['FF0000', '0000FF']).toContain(shown!.fill!);
    expect(cellProperties(bob, 0, 'shd')).toEqual([shown!]);
    harness.apply(bob, [fill(bob, '00FF00')]);
    expectHealthy(alice, bob);
    expect(cellProperties(alice, 0, 'shd')).toEqual([{ val: 'clear', fill: '00FF00' }]);
  });

  test('each author undoes a concurrent fill', async () => {
    const { alice, bob } = await concurrently(
      table(TC_W),
      (peer) => [fill(peer, 'FF0000')],
      (peer) => [fill(peer, '0000FF')]
    );
    expectHealthy(alice, bob);
    expect(cellProperties(alice, 0, 'shd')).toHaveLength(1);
    expect(alice.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(cellProperties(bob, 0, 'shd')).toHaveLength(1);
    expect(bob.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(cellProperties(alice, 0, 'shd')).toEqual([]);
    expect(cellProperties(alice, 0, 'tcW')).toEqual([{ w: '3000', type: 'dxa' }]);
  });

  test('a fill and a vertical alignment for a cell without w:tcPr both stay', async () => {
    const { alice, bob } = await concurrently(
      table(),
      (peer) => [fill(peer, 'FF0000')],
      (peer) => [verticalAlignment(peer, 'center')]
    );
    expectHealthy(alice, bob);
    for (const peer of [alice, bob]) {
      expect(cellProperties(peer, 0, 'shd')).toEqual([{ val: 'clear', fill: 'FF0000' }]);
      expect(cellProperties(peer, 0, 'vAlign')).toEqual([{ val: 'center' }]);
    }
    // Each undo takes back its own author's property only.
    expect(alice.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(cellProperties(bob, 0, 'shd')).toEqual([]);
    expect(cellProperties(bob, 0, 'vAlign')).toEqual([{ val: 'center' }]);
    expect(bob.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(cellProperties(alice, 0, 'vAlign')).toEqual([]);
    // Both authors keep editing the cell.
    harness.apply(alice, [verticalAlignment(alice, 'bottom')]);
    harness.apply(bob, [fill(bob, '00FF00')]);
    expectHealthy(alice, bob);
    expect(cellProperties(bob, 0, 'vAlign')).toEqual([{ val: 'bottom' }]);
    expect(cellProperties(alice, 0, 'shd')).toEqual([{ val: 'clear', fill: '00FF00' }]);
  });

  test('two column divider drags keep one w:tcW in every cell', async () => {
    const { alice, bob } = await concurrently(
      table(TC_W),
      (peer) => [columnWidths(peer, 2000, 4000)],
      (peer) => [columnWidths(peer, 4000, 2000)]
    );
    expectHealthy(alice, bob);
    for (const cell of [0, 1, 2, 3]) {
      const widths = cellProperties(alice, cell, 'tcW');
      expect(widths).toHaveLength(1);
      expect(['2000', '4000']).toContain(widths[0]!.w!);
    }
    harness.apply(bob, [columnWidths(bob, 3500, 2500)]);
    expectHealthy(alice, bob);
    expect([0, 1].map((cell) => cellProperties(alice, cell, 'tcW')[0]!.w)).toEqual([
      '3500',
      '2500',
    ]);
  });

  test('each author undoes a concurrent column divider drag', async () => {
    const { alice, bob } = await concurrently(
      table(TC_W),
      (peer) => [columnWidths(peer, 2000, 4000)],
      (peer) => [columnWidths(peer, 4000, 2000)]
    );
    expect(alice.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    expect(bob.room.session.undo()).toBe(true);
    expectHealthy(alice, bob);
    for (const cell of [0, 1, 2, 3]) {
      expect(cellProperties(alice, cell, 'tcW')).toEqual([{ w: '3000', type: 'dxa' }]);
    }
  });

  test('two border edits of one cell side keep one side element', async () => {
    const spec = (value: string) =>
      ({ style: 'single', size: 8, color: { kind: 'hex', value } }) as const;
    const border = (peer: Peer, value: string): TreeDocOp => ({
      op: 'setTableCellBorders',
      tableId: ids(peer).table,
      cellIds: [ids(peer).cells[0]!],
      scope: 'left',
      spec: spec(value),
    });
    const { alice, bob } = await concurrently(
      table('<w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4"/></w:tcBorders></w:tcPr>'),
      (peer) => [border(peer, 'FF0000')],
      (peer) => [border(peer, '0000FF')]
    );
    expectHealthy(alice, bob);
    harness.apply(alice, [fill(alice, '00FF00')]);
    expectHealthy(alice, bob);
    expect(cellProperties(bob, 0, 'shd')).toEqual([{ val: 'clear', fill: '00FF00' }]);
  });

  test('two tracked deletions of one row keep one deletion on the row and each cell', async () => {
    const remove = (peer: Peer, author: string): TreeDocOp => ({
      op: 'deleteTableRow',
      tableId: ids(peer).table,
      rowId: ids(peer).rows[0]!,
      revision: { author },
    });
    const { alice, bob } = await concurrently(
      table(),
      (peer) => [remove(peer, 'alice')],
      (peer) => [remove(peer, 'bob')]
    );
    expectHealthy(alice, bob);
    const [trPr] = wmlChildren(elements(alice, 'tableRow')[0]!, 'trPr');
    expect(wmlChildren(trPr!, 'del')).toHaveLength(1);
    for (const cell of [0, 1]) expect(cellProperties(alice, cell, 'cellDel')).toHaveLength(1);
    harness.apply(bob, [rowHeight(bob, 500)]);
    expectHealthy(alice, bob);
    expect(rowHeights(alice)).toEqual(['500']);
  });

  test('a row height written while a peer inserts a first column keeps w:trPr ahead', async () => {
    // Concurrent inserts at one index order by Yjs client id, so try both orders.
    for (const clientIds of [
      { alice: 1, bob: 2 },
      { alice: 2, bob: 1 },
    ]) {
      const ordered = createPeerHarness('concurrent-table-leading-room', undefined, clientIds);
      try {
        const { alice, bob, pause, resume } = await ordered.pair(table());
        pause();
        ordered.apply(alice, [rowHeight(alice, 400)]);
        ordered.apply(bob, [
          {
            op: 'insertTableColumn',
            tableId: ids(bob).table,
            where: 'left',
            gridColumnId: ids(bob).columns[0]!,
          },
        ]);
        resume();
        for (const peer of [alice, bob]) {
          expect(peer.room.session.status()).toBe('ready');
          expectValidTables(peer);
        }
        ordered.expectConverged(alice, bob);
        expect(rowHeights(bob)).toEqual(['400']);
        // Later edits address the order every replica shows.
        ordered.apply(bob, [rowHeight(bob, 500)]);
        ordered.apply(alice, [
          {
            op: 'insertText',
            paragraphId: ordered.paragraphIdAt(alice, 0),
            offset: 0,
            text: 'New',
          },
        ]);
        ordered.expectConverged(alice, bob);
        expectValidTables(alice);
        expect(rowHeights(alice)).toEqual(['500']);
        expect(nodeText(elements(bob, 'tableCell')[0]!)).toBe('New');
      } finally {
        ordered.cleanup();
      }
    }
  });

  test('paragraph and run properties written at once in a cell still show once', async () => {
    // The paragraph and run rules this file's table rules extend, unchanged.
    const { alice, bob } = await concurrently(
      table(TC_W),
      (peer) => [
        {
          op: 'setParagraphProperties',
          paragraphId: harness.paragraphIdAt(peer, 0),
          properties: [{ localName: 'jc', attributes: { val: 'center' } }],
        },
      ],
      (peer) => [
        {
          op: 'setParagraphProperties',
          paragraphId: harness.paragraphIdAt(peer, 0),
          properties: [{ localName: 'jc', attributes: { val: 'right' } }],
        },
        {
          op: 'setRunProperties',
          paragraphId: harness.paragraphIdAt(peer, 0),
          start: 0,
          end: 1,
          properties: [{ localName: 'b' }],
        },
      ]
    );
    expectHealthy(alice, bob);
    const [paragraph] = elements(bob, 'paragraph');
    expect(paragraph!.children[0]!.kind).toBe('paragraphProperties');
    expect(elements(bob, 'paragraphProperties')).toHaveLength(1);
    expect(wmlChildren(elements(alice, 'paragraphProperties')[0]!, 'jc')).toHaveLength(1);
    expect(nodeText(paragraph!)).toBe('A');
  });
});

describe('concurrent table properties from two editors', () => {
  async function mount(bytes: Uint8Array) {
    const peers = await harness.pair(bytes);
    const mounted = [peers.alice, peers.bob].map((peer) => {
      peer.detach();
      const container = document.createElement('div');
      document.body.append(container);
      containers.push(container);
      const editor = createDocxEditor({
        container,
        zoom: 0.75,
        document: peer.room.document,
        modules: [collaborationModule({ session: peer.room.session })],
      });
      editors.push(editor);
      return { editor, container };
    });
    const sync = () => {
      peers.alice.room.session.flushPendingJournals();
      peers.bob.room.session.flushPendingJournals();
    };
    const converged = () => {
      sync();
      const [left, right] = mounted.map(({ editor }) => editor.surface!.session.currentPackage());
      expect(packageFingerprint(right!)).toBe(packageFingerprint(left!));
      expect(saveReopenDigest(right!)).toEqual(saveReopenDigest(left!));
      for (const { editor } of mounted) expectValidTables(editor);
    };
    return { peers, alice: mounted[0]!, bob: mounted[1]!, sync, converged };
  }

  const tableOf = (editor: Editor) => {
    const found = editor
      .surface!.layout()
      .pages[0]!.fragments.find((fragment) => fragment.kind === 'table');
    if (found?.kind !== 'table') throw new Error('missing table');
    return found;
  };

  const selectCell = (editor: Editor, cell = 0) => {
    const block = tableOf(editor).rows[0]!.cells[cell]!.blocks[0]!;
    if (block.kind !== 'paragraph') throw new Error('missing cell paragraph');
    const position = { paragraphId: block.paragraphId, offset: 0 };
    editor.surface!.setSelection({ anchor: position, head: position });
  };

  const pointer = (type: string, x: number, y: number, buttons = 0) =>
    new PointerEvent(type, {
      bubbles: true,
      cancelable: true,
      button: 0,
      buttons,
      pointerId: 1,
      pointerType: 'mouse',
      clientX: 100 + 72 + x,
      clientY: 50 + 72 + y,
    });

  /** Drag the first row's bottom divider down by `points`, as a pointer does. */
  const dragRowDivider = (
    { editor, container }: { editor: Editor; container: HTMLElement },
    points: number
  ) => {
    const pages = container.querySelector<HTMLElement>('.docx-pages')!;
    Object.defineProperty(pages, 'getBoundingClientRect', {
      configurable: true,
      value: () => ({
        left: 100,
        top: 50,
        right: 1100,
        bottom: 1050,
        width: 1000,
        height: 1000,
        x: 100,
        y: 50,
      }),
    });
    const before = tableOf(editor);
    const x = before.box.x + before.box.width / 4;
    const bottom = before.rows[0]!.box.y + before.rows[0]!.box.height;
    pages.dispatchEvent(pointer('pointermove', x, bottom));
    const handle = container.querySelector<HTMLElement>('.docx-table-row-divider-handle');
    expect(handle).not.toBeNull();
    handle!.dispatchEvent(pointer('pointerdown', x, bottom));
    document.dispatchEvent(pointer('pointermove', x, bottom + points, 1));
    document.dispatchEvent(pointer('pointerup', x, bottom + points));
  };

  const resizeColumns = (editor: Editor, left: number, right: number) => {
    const table = tableOf(editor);
    const row = table.rows[0]!;
    const target = tableColumnDividerResizeTargetFrom(
      editor.surface!.layout().revision,
      { table, row, rowIndex: 0 },
      row.cells[0]!.gridColumnId!,
      row.cells[1]!.gridColumnId!
    );
    return editor.exec({
      type: 'commitTableColumnDividerResize',
      target,
      leftWidthTwips: left,
      rightWidthTwips: right,
    }).ok;
  };

  test('row divider drags at once leave one row height, and both authors undo', async () => {
    const { peers, alice, bob, sync, converged } = await mount(table(TC_W));
    peers.pause();
    dragRowDivider(alice, 20);
    dragRowDivider(bob, 40);
    sync();
    peers.resume();
    converged();
    expect(rowHeights(alice.editor)).toHaveLength(1);
    expect(tableOf(bob.editor).rows[0]!.box).toEqual(tableOf(alice.editor).rows[0]!.box);
    // The resized row moves what follows it: a remote caret there lands where its author sees it.
    const after = alice.editor
      .surface!.layout()
      .pages[0]!.fragments.find((fragment) => fragment.kind === 'paragraph');
    if (after?.kind !== 'paragraph') throw new Error('missing paragraph after the table');
    const position = { paragraphId: after.paragraphId, offset: 0 };
    alice.editor.surface!.setSelection({ anchor: position, head: position });
    applyAwarenessUpdate(
      peers.bob.awareness,
      encodeAwarenessUpdate(peers.alice.awareness, [peers.alice.awareness.clientID]),
      'test'
    );
    const remote = peers.bob.room.session.remoteSelections()[0]!;
    expect(
      caretAt(bob.editor.surface!.layout(), {
        paragraphId: remote.head.nodeId,
        offset: remote.head.offset,
      })
    ).toEqual(caretAt(alice.editor.surface!.layout(), position));
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(rowHeights(bob.editor)).toHaveLength(1);
    expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(rowHeights(alice.editor)).toEqual([]);
    expect(alice.editor.exec({ type: 'redo' }).ok).toBe(true);
    converged();
    expect(rowHeights(bob.editor)).toHaveLength(1);
  });

  test('column divider drags at once keep one w:tcW per cell, and both authors undo', async () => {
    const { peers, alice, bob, sync, converged } = await mount(table(TC_W));
    peers.pause();
    expect(resizeColumns(alice.editor, 2000, 4000)).toBe(true);
    expect(resizeColumns(bob.editor, 4000, 2000)).toBe(true);
    sync();
    peers.resume();
    converged();
    expect(tableOf(bob.editor).columnEdges).toEqual(tableOf(alice.editor).columnEdges);
    for (const cell of [0, 1, 2, 3]) {
      expect(cellProperties(alice.editor, cell, 'tcW')).toHaveLength(1);
    }
    // The table still takes a resize, and saves and reopens as both editors show it.
    expect(resizeColumns(bob.editor, 3500, 2500)).toBe(true);
    converged();
    expect(cellProperties(alice.editor, 0, 'tcW')).toEqual([{ w: '3500', type: 'dxa' }]);
    const reopened = createDocxEditor({
      container: document.createElement('div'),
      document: await alice.editor.save(),
    });
    editors.push(reopened);
    expect(saveReopenDigest(reopened.surface!.session.currentPackage())).toEqual(
      saveReopenDigest(bob.editor.surface!.session.currentPackage())
    );
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
  });

  test('a fill and a vertical alignment from two editors on one cell both stay', async () => {
    const { peers, alice, bob, sync, converged } = await mount(table());
    selectCell(alice.editor);
    selectCell(bob.editor);
    peers.pause();
    const red = { kind: 'hex', value: 'FF0000' } as const;
    expect(alice.editor.exec({ type: 'setCellFill', color: red }).ok).toBe(true);
    expect(bob.editor.exec({ type: 'setTableCellVerticalAlignment', alignment: 'center' }).ok).toBe(
      true
    );
    sync();
    peers.resume();
    converged();
    for (const { editor } of [alice, bob]) {
      expect(cellProperties(editor, 0, 'shd')).toEqual([{ val: 'clear', fill: 'FF0000' }]);
      expect(cellProperties(editor, 0, 'vAlign')).toEqual([{ val: 'center' }]);
    }
    // Two copies of the cell's `w:tcPr` refused this fill with `duplicate-property-container`.
    const green = { kind: 'hex', value: '00FF00' } as const;
    expect(bob.editor.exec({ type: 'setCellFill', color: green }).ok).toBe(true);
    converged();
    expect(cellProperties(alice.editor, 0, 'shd')).toEqual([{ val: 'clear', fill: '00FF00' }]);
    expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(alice.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(cellProperties(bob.editor, 0, 'shd')).toEqual([]);
    expect(bob.editor.exec({ type: 'undo' }).ok).toBe(true);
    converged();
    expect(cellProperties(alice.editor, 0, 'vAlign')).toEqual([]);
  });
});
