// Load then compose gives back the same story parts, for every kind of item Word puts in a
// story: once reusing untouched paragraphs, and once drawing every paragraph again from its
// items. The view of each paragraph reads as the engine's model text.

import { describe, expect, test } from 'bun:test';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import { paragraphModelTextOf } from '../../../store/store/paragraph-model-text.ts';
import { composeOrderedPart, loadOrderedPackage } from '../index.ts';
import { isTextBoxStoryName } from '../model/text-boxes.ts';
import { expectSameFingerprint, expectValid } from './oracle.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, packageOf } from './word-fixtures.ts';

describe('ordered load and compose: the item-kind matrix', () => {
  test('the matrix holds every kind once', () => {
    expect(new Set(MATRIX_ITEMS.map((item) => item.key)).size).toBe(64);
  });

  for (const item of MATRIX_ITEMS) {
    test(`${item.key}: ${item.name}`, () => {
      const pkg = packageOf(docx(bodyOf(item), MATRIX_EXTRAS));
      const doc = loadOrderedPackage(pkg);
      // A text box is a story of its own, drawn back into its object (group 9.9)
      expect([...doc.stories.keys()].filter((name) => !isTextBoxStoryName(name))).toEqual([
        '/word/document.xml',
        '/word/endnotes.xml',
        '/word/footnotes.xml',
      ]);
      for (const [name, story] of doc.stories) {
        for (const rebuild of [false, true]) {
          const label = `${item.key} ${name} rebuild=${rebuild}`;
          const { part, composed } = composeOrderedPart(doc, name, { rebuild });
          expectSameFingerprint(part, story.part, label);
          expectValid(part, label);
          expect(composed.prompts).toEqual([]);
          for (const view of composed.paragraphs) {
            const paragraph = findNode(story.part, view.id);
            if (!paragraph || paragraph.kind !== 'paragraph') continue;
            expect({ label, text: view.chars.map((c) => c.ch).join('') }).toEqual({
              label,
              text: paragraphModelTextOf(paragraph),
            });
          }
        }
      }
    });
  }
});
