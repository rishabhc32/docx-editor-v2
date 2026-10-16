/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The visibility rule for table rows, cells and grids, whose property containers the tree
// keeps generic, on child lists edits do not produce.

import { describe, expect, test } from 'bun:test';
import { WML_NAMESPACE_URI } from '@docx-editor.dev/core/store';
import {
  partitionChildIds,
  partitionWithMerges,
  singletonLayout,
  type ChildShell,
} from '../document/singleton-children.ts';
import { ruleKindOf } from '../document/singleton-table-rules.ts';

const peer = (replica: string, counter: number) => `lid:${replica.repeat(32)}:${counter}`;
const seed = (counter: number) => `/word/document.xml#${counter}`;

const generic = (localName: string): ChildShell => ({
  kind: 'generic',
  namespaceUri: WML_NAMESPACE_URI,
  localName,
});
const typed = (kind: string, localName: string): ChildShell => ({
  kind,
  namespaceUri: WML_NAMESPACE_URI,
  localName,
});

const TRPR_A = peer('a', 1);
const TRPR_B = peer('b', 1);
const TBLPREX = seed(1);
const CELL = seed(2);
const NEW_CELL = peer('b', 2);
const TCPR = peer('a', 3);
const PARAGRAPH = seed(3);
const BOOKMARK = seed(4);

const shells: Record<string, ChildShell> = {
  [TRPR_A]: generic('trPr'),
  [TRPR_B]: generic('trPr'),
  [TBLPREX]: generic('tblPrEx'),
  [CELL]: typed('tableCell', 'tc'),
  [NEW_CELL]: typed('tableCell', 'tc'),
  [TCPR]: generic('tcPr'),
  [PARAGRAPH]: typed('paragraph', 'p'),
  [BOOKMARK]: typed('bookmarkStart', 'bookmarkStart'),
};
const shellOf = (id: string): ChildShell | null => shells[id] ?? null;

describe('table singleton child layout', () => {
  test('a generic WML element goes by its name, other nodes by their kind', () => {
    expect(ruleKindOf(generic('tcPr'))).toBe('wml:tcPr');
    expect(ruleKindOf(typed('tableRow', 'tr'))).toBe('tableRow');
    expect(ruleKindOf({ kind: 'generic', namespaceUri: 'urn:other', localName: 'tcPr' })).toBe(
      'generic'
    );
  });

  test("a second w:trPr hides behind the first, which also shows the second's children", () => {
    const result = partitionWithMerges('tableRow', [TRPR_A, CELL, TRPR_B], shellOf);
    expect(result.shown).toEqual([TRPR_A, CELL]);
    expect(result.hidden).toEqual([TRPR_B]);
    expect([...(result.merges ?? [])]).toEqual([[TRPR_A, [TRPR_B]]]);
  });

  test('a w:trPr moves ahead of a cell, but not ahead of a w:tblPrEx', () => {
    expect(singletonLayout('tableRow', [TBLPREX, NEW_CELL, TRPR_A], shellOf)?.order).toEqual([
      0, 2, 1,
    ]);
    // The editor writes a new `w:trPr` first in the row, so that order stays as written.
    expect(singletonLayout('tableRow', [TRPR_A, TBLPREX, CELL], shellOf)).toBeNull();
  });

  test('a w:tcPr moves ahead of a paragraph, but not ahead of a bookmark', () => {
    expect(singletonLayout('tableCell', [PARAGRAPH, TCPR], shellOf)?.order).toEqual([1, 0]);
    // The editor writes a new `w:tcPr` after a cell's leading bookmarks, so it stays there.
    expect(singletonLayout('tableCell', [BOOKMARK, TCPR, PARAGRAPH], shellOf)).toBeNull();
  });

  test('a property a peer wrote twice shows once, and a file’s own two stay', () => {
    const local: Record<string, ChildShell> = {
      [peer('a', 4)]: generic('shd'),
      [peer('b', 4)]: generic('vAlign'),
      [peer('b', 5)]: generic('shd'),
      [seed(5)]: generic('trHeight'),
      [seed(6)]: generic('trHeight'),
    };
    const localShell = (id: string) => local[id] ?? null;
    const layout = singletonLayout(
      ruleKindOf(generic('tcPr')),
      [peer('a', 4), peer('b', 4), peer('b', 5)],
      localShell
    );
    expect(layout?.order).toEqual([0, 1]);
    expect([...(layout?.companions ?? [])]).toEqual([[0, [2]]]);
    expect(singletonLayout('wml:trPr', [seed(5), seed(6)], localShell)).toBeNull();
  });

  test('a w:tblGrid shows one w:tblGridChange and every w:gridCol', () => {
    const local: Record<string, ChildShell> = {
      [seed(7)]: generic('gridCol'),
      [seed(8)]: generic('gridCol'),
      [peer('a', 6)]: generic('tblGridChange'),
      [peer('b', 6)]: generic('tblGridChange'),
    };
    const result = partitionChildIds(
      'tableGrid',
      [seed(7), seed(8), peer('a', 6), peer('b', 6)],
      (id) => local[id] ?? null
    );
    expect(result.shown).toEqual([seed(7), seed(8), peer('a', 6)]);
    expect(result.hidden).toEqual([peer('b', 6)]);
  });

  test('a repeated listing of a hidden copy hides with it', () => {
    // Two undos that each restore one deleted child list it twice.
    const local: Record<string, ChildShell> = {
      [peer('a', 7)]: typed('paragraphProperties', 'pPr'),
      [seed(9)]: typed('paragraphProperties', 'pPr'),
      [seed(10)]: typed('run', 'r'),
    };
    const result = partitionChildIds(
      'paragraph',
      [peer('a', 7), seed(9), seed(10), seed(9)],
      (id) => local[id] ?? null
    );
    expect(result.shown).toEqual([peer('a', 7), seed(10)]);
  });
});
