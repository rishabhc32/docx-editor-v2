// The coverage guard: every entry of the editor's feature list (`docs/site/data/word-features.ts`)
// has matrix kinds that put it in a story, a by-name refusal test, or a stated reason it has no
// story content. A new entry with none of them fails here, so no feature is missed later. Also:
// every known difference still matches a matrix cell, and every refusal case is refused by name
// with nothing changed.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { KINDS } from './catalogue.ts';
import { COVERAGE } from './coverage.ts';
import { KNOWN, knownFor } from './known.ts';
import { REFUSALS, runRefusal } from './refusals.ts';
import { SETS, type SetName } from './scenarios.ts';

const FEATURES = fileURLToPath(
  new URL('../../../../../../../docs/site/data/word-features.ts', import.meta.url)
);

/** The ids of the feature list's entries, read from its source. */
function featureIds(): string[] {
  const source = readFileSync(FEATURES, 'utf8');
  return [...source.matchAll(/^\s*id: '([a-z0-9.-]+)',$/gm)].map((match) => match[1] ?? '');
}

/** Every kind a scenario exercises, across the sets. */
function exercised(): Set<string> {
  const out = new Set<string>();
  for (const set of Object.values(SETS)) for (const s of set) for (const k of s.kinds) out.add(k);
  return out;
}

describe('the matrix covers the feature list', () => {
  const ids = featureIds();

  test('the feature list reads', () => {
    expect(ids.length).toBeGreaterThan(90);
    expect(new Set(ids).size).toBe(ids.length);
  });

  test('every entry has kinds, a refusal, or a reason it has no story content', () => {
    const missing = ids.filter((id) => {
      const entry = COVERAGE[id];
      return (
        !entry ||
        ((entry.kinds?.length ?? 0) === 0 &&
          (entry.refusals?.length ?? 0) === 0 &&
          (entry.outside ?? '').length === 0)
      );
    });
    expect(missing).toEqual([]);
  });

  test('the map names no entry the list lacks', () => {
    expect(Object.keys(COVERAGE).filter((id) => !ids.includes(id))).toEqual([]);
  });

  test('every kind the map names is in the matrix, and a scenario meets it', () => {
    const known = new Set(KINDS.map((kind) => kind.key));
    const met = exercised();
    const named = Object.values(COVERAGE).flatMap((entry) => entry.kinds ?? []);
    expect(named.filter((key) => !known.has(key))).toEqual([]);
    expect(named.filter((key) => !met.has(key))).toEqual([]);
  });

  test('every refusal the map names has its case', () => {
    const cases = new Set(REFUSALS.map((refusal) => refusal.id));
    const named = Object.values(COVERAGE).flatMap((entry) => entry.refusals ?? []);
    expect(named.filter((id) => !cases.has(id))).toEqual([]);
  });
});

describe('what the first release does not cover is refused by name', () => {
  for (const refusal of REFUSALS)
    test(refusal.name, () => {
      expect(runRefusal(refusal)).toEqual({ reason: 'not-covered', unchanged: true });
    });
});

test('every known difference matches a matrix cell', () => {
  const used = new Set<number>();
  for (const [set, scenarios] of Object.entries(SETS) as [SetName, (typeof SETS)[SetName]][])
    for (const s of scenarios)
      for (const entry of knownFor(set, s.name)) used.add(KNOWN.indexOf(entry));
  expect(KNOWN.map((entry, k) => (used.has(k) ? null : entry.cause)).filter(Boolean)).toEqual([]);
});
