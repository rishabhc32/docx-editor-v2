// Save then load for the repository's DOCX fixtures, the ones the round-trip tests use: the
// reopened document has the same view and gives the same result for each review card and for
// every card at once. The reference is the engine's own save and reopen of the file: the
// engine's package writer rebuilds some parts (a text box's fallback copy takes its main copy's
// content), and that is the engine's, not this save's. Sharded to keep each process's memory
// small.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { readOoxmlPackage, writeOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { canonicalOoxmlFingerprint } from '../../../store/package/ooxml-tree.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { settleDecidedTables } from '../change/settle-columns.ts';
import { composeOrderedPart } from '../compose/compose.ts';
import { orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { setStatus } from '../model/mutations.ts';
import type { OrderedDocument } from '../model/types.ts';
import { takeBack, withUndoRecord } from '../model/undo-record.ts';
import { normalizedBody, semanticBody } from './matrix/compare.ts';

const FIXTURES = join(import.meta.dir, '../../../../../../e2e/fixtures');
const MAX_BYTES = 2_000_000;
/**
 * The 521-page file, and a file of 6.8 MB of XML in 260 KB: its saves and reopens here take
 * over 3 GB, and the round-trip tests cover its load and compose.
 */
const SKIPPED = new Set(['typing-perf-521pp.docx', 'list-pagination-break.docx']);
/** Cards decided one by one in each file; the rest are decided all at once. */
const CARDS = 4;

const FILES = readdirSync(FIXTURES)
  .filter((name) => name.endsWith('.docx') && !SKIPPED.has(name))
  .filter((name) => statSync(join(FIXTURES, name)).size < MAX_BYTES)
  .sort();

function view(doc: OrderedDocument, ids: readonly string[], status?: 'accepted' | 'rejected') {
  const { record } = withUndoRecord(doc, () => {
    if (status) {
      for (const id of ids) setStatus(doc, id, status);
      settleDecidedTables(doc, ids, orderedIdMinter('check:1'));
    }
  });
  const parts = [...doc.stories.keys()].map((name) => {
    const part = composeOrderedPart(doc, name).part;
    return status
      ? semanticBody(normalizedBody(part, (id) => id))
      : canonicalOoxmlFingerprint(part);
  });
  takeBack(doc, record);
  return parts;
}

/** The tests of one shard of the fixtures. */
export function saveFixturesShard(index: number, count: number): void {
  const files = FILES.filter((_, k) => k % count === index);
  describe(`save then load, the repository fixtures, shard ${index + 1} of ${count}`, () => {
    test('finds the fixtures', () => {
      expect(FILES.length).toBeGreaterThan(90);
    });
    for (const name of files)
      test(name, () => {
        const read = readOoxmlPackage(new Uint8Array(readFileSync(join(FIXTURES, name))), {});
        // A file the engine itself refuses has nothing to load
        if (!read.ok) return;
        const saved = readOoxmlPackage(
          orderedPackageBytes(loadOrderedPackage(read.package), read.package),
          {}
        );
        if (!saved.ok) throw new Error(`${name}: the saved file does not open: ${saved.reason}`);
        const again = loadOrderedPackage(saved.package);
        const engine = readOoxmlPackage(writeOoxmlPackage(read.package), {});
        if (!engine.ok) throw new Error(`${name}: the engine's own save does not open`);
        const doc = loadOrderedPackage(engine.package);
        expect({ name, view: view(again, []) }).toEqual({ name, view: view(doc, []) });
        const live = orderedCards(doc);
        const later = orderedCards(again);
        expect({ name, cards: later.map((c) => c.suggestions) }).toEqual({
          name,
          cards: live.map((c) => c.suggestions),
        });
        const all = live.flatMap((card) => card.suggestions);
        for (const status of ['accepted', 'rejected'] as const) {
          expect({ name, status, all: view(again, all, status) }).toEqual({
            name,
            status,
            all: view(doc, all, status),
          });
          for (const card of live.slice(0, CARDS))
            expect({
              name,
              card: card.id,
              status,
              view: view(again, card.suggestions, status),
            }).toEqual({
              name,
              card: card.id,
              status,
              view: view(doc, card.suggestions, status),
            });
        }
      });
  });
}
