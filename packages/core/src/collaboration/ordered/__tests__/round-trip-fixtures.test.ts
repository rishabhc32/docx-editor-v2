// Load then compose gives back every story part of the repository's own DOCX fixtures, once
// reusing untouched blocks and once drawing every block again from its items. Fixtures of 2 MB
// or more are skipped, and so is the 521-page typing fixture, to keep this file's memory small.

import { describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { readOoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { composeOrderedPart, loadOrderedPackage } from '../index.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import { expectSameFingerprint, expectValid } from './oracle.ts';

const FIXTURES = join(import.meta.dir, '../../../../../../e2e/fixtures');
const MAX_BYTES = 2_000_000;
const SKIPPED = new Set(['typing-perf-521pp.docx']);

const files = readdirSync(FIXTURES)
  .filter((name) => name.endsWith('.docx') && !SKIPPED.has(name))
  .filter((name) => statSync(join(FIXTURES, name)).size < MAX_BYTES)
  .sort();

describe('ordered load and compose: the repository fixtures', () => {
  test('finds the fixtures', () => {
    expect(files.length).toBeGreaterThan(90);
  });

  for (const name of files) {
    test(name, () => {
      const read = readOoxmlPackage(new Uint8Array(readFileSync(join(FIXTURES, name))), {});
      // A file the engine itself refuses has nothing to load
      if (!read.ok) return;
      const doc = loadOrderedPackage(read.package);
      for (const [storyName, story] of doc.stories) {
        for (const rebuild of [false, true]) {
          const label = `${name} ${storyName} rebuild=${rebuild}`;
          const { part, composed } = composeOrderedPart(doc, storyName, { rebuild });
          expectSameFingerprint(part, story.part, label);
          // A text box's story is no part: its blocks are checked in the host part, drawn there
          if (!isTextBoxStoryName(storyName)) expectValid(part, label);
          expect({ label, prompts: composed.prompts.length }).toEqual({ label, prompts: 0 });
        }
      }
    });
  }
});
