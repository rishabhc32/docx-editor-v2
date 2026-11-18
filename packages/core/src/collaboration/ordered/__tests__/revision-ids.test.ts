// Word revision ids (`w:id`) are the same on every copy, whenever each copy composes: a file's
// revision keeps its own id, and every other id is a function of its suggestion's id and the
// role of the mark, past every id the file holds.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { composeOrderedParagraph, composeOrderedPart } from '../compose/compose.ts';
import { loadOrderedPackage } from '../load/load.ts';
import { revisionIdFor } from '../model/revision-ids.ts';
import type { OrderedDocument } from '../model/types.ts';
import { seededRandom } from './random.ts';
import {
  addressOf,
  apply,
  at,
  BODY,
  EDIT,
  load,
  markOf,
  opens,
  partsOf,
} from './change-apply-fixture.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

/** A run of changes over every kind that writes a revision mark. */
function changes(doc: OrderedDocument, step: number): Parameters<typeof apply>[2] | null {
  const rows = opens(doc, 'tr');
  const [link] = opens(doc, 'hyperlink');
  switch (step) {
    case 0:
      return { kind: 'insert', ...EDIT, place: at(doc, 0, 4), text: 'new ' };
    case 1:
      return { kind: 'delete', ...EDIT, from: at(doc, 8, 0), to: at(doc, 9, 6) };
    case 2:
      return { kind: 'split', ...EDIT, place: at(doc, 0, 8) };
    case 3:
      return {
        kind: 'format',
        ...EDIT,
        from: at(doc, 0, 0),
        to: at(doc, 0, 3),
        properties: [{ localName: 'i' }],
      };
    case 4:
      return {
        kind: 'paragraph-property',
        ...EDIT,
        marks: [markOf(doc, 3)],
        properties: [{ localName: 'jc', attributes: { val: 'right' } }],
      };
    case 5:
      return rows[0]
        ? { kind: 'insert-row', ...EDIT, row: addressOf(rows[0]), where: 'above' }
        : null;
    case 6:
      return link
        ? { kind: 'link-target', ...EDIT, link: addressOf(link), href: 'https://e.com/x' }
        : null;
    case 7:
      return {
        kind: 'move',
        ...EDIT,
        from: at(doc, 3, 0),
        to: at(doc, 3, 5),
        place: at(doc, 3, 11),
      };
    case 8:
      return { kind: 'insert-tab', ...EDIT, place: at(doc, 1, 0) };
    case 9:
      // The table holds a new row, so the engine refuses a column there; a note instead
      return { kind: 'insert-note', ...EDIT, place: at(doc, 0, 2), note: 'endnote', text: ' See.' };
    case 10:
      return { kind: 'join', ...EDIT, mark: markOf(doc, 1) };
    default:
      return null;
  }
}

const STEPS = 11;

function run(composeBetween: (doc: OrderedDocument, step: number) => void): OrderedDocument {
  const doc = load();
  for (let step = 0; step < STEPS; step += 1) {
    const body = changes(doc, step);
    if (body === null) throw new Error(`step ${step} has no target`);
    apply(doc, `ana:${step + 1}`, body, step % 2 === 0 ? 'Ana' : 'Ben');
    composeBetween(doc, step);
  }
  return doc;
}

describe('Word revision ids', () => {
  test('two copies that compose at different times write the same parts', () => {
    const late = run(() => undefined);
    const early = run((doc) => {
      partsOf(doc);
    });
    const random = seededRandom(7);
    const scattered = run((doc) => {
      // One paragraph here and there, in no order: the ids it draws are minted on the way
      const count = doc.stories.get(BODY)?.tree.summary.marks ?? 0;
      for (let k = 0; k < 3; k += 1)
        composeOrderedParagraph(doc, BODY, Math.floor(random() * count));
    });
    const expected = partsOf(late);
    expect(partsOf(early)).toEqual(expected);
    expect(partsOf(scattered)).toEqual(expected);
    // Composing again changes nothing
    expect(partsOf(late)).toEqual(expected);
    expect(expected.join('')).toContain('w:id="1');
  });

  test("a file's revision keeps its id, and a new mark avoids every id the file holds", () => {
    const fresh = load();
    apply(fresh, 'ana:1', { kind: 'insert', ...EDIT, place: at(fresh, 0, 0), text: 'x' });
    const minted = revisionIdFor(fresh, 'S:ana:1.1', 'ins');
    // A file whose own insertion already has the id the new mark would take
    const REV = 'w:author="Cara" w:date="2026-01-01T00:00:00Z"';
    const file = docx(
      p(t('Alpha ') + `<w:ins w:id="${minted}" ${REV}>${t('beta')}</w:ins>`) + p(t('Gamma.'))
    );
    const doc = loadOrderedPackage(packageOf(file));
    apply(doc, 'ana:1', { kind: 'insert', ...EDIT, place: at(doc, 1, 0), text: 'x' });
    const xml = serializeOoxmlPart(composeOrderedPart(doc, BODY).part);
    const ids = [...xml.matchAll(/<w:ins w:author="(\w+)"[^>]* w:id="(\d+)"/g)].map((m) => [
      m[1],
      m[2],
    ]);
    expect(ids).toHaveLength(2);
    expect(ids[0]).toEqual(['Cara', minted]);
    expect(ids[1]?.[0]).toBe('Ana');
    expect(ids[1]?.[1]).not.toBe(minted);
    expect(Number(ids[1]?.[1])).toBeLessThan(2 ** 31);
  });
});
