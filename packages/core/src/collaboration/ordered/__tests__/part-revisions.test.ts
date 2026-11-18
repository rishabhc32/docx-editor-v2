// A file's tracked changes outside the stories as suggestions: each change record
// group of the styles part that the editor alone shows as a card is one pending suggestion with
// no items, loaded after the stories with the same id on every copy. A decision turns its status,
// and compose writes the styles part from the file's with the decided records resolved by the
// engine, in any order of decisions; a take-back gives the file's records back.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedSeenCards } from '../change/apply-decide.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPackage } from '../compose/package.ts';
import { orderedCardOf } from '../card-of.ts';
import { orderedCards } from '../decide.ts';
import { withUndoRecord, takeBack } from '../model/undo-record.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedDocument } from '../model/types.ts';
import { STYLED } from './style-revision-fixture.ts';
import { docx, p, t, W } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const STYLES = '/word/styles.xml';

function opened() {
  return openOrderedWordFile(STYLED, { file: 1 });
}

/** The styles part a save writes. */
function savedStyles(doc: OrderedDocument, base = opened().base): string {
  return serializeOoxmlPart(composeOrderedPackage(doc, base).parts.get(STYLES)!);
}

/** The suggestion of the record group of `element`. */
function recordOf(doc: OrderedDocument, element: string, id: string): string {
  for (const [suggestion, revision] of doc.partRevisions)
    if (revision.element === element && revision.revision.id === id) return suggestion;
  throw new Error(`no ${element} ${id}`);
}

function decide(
  doc: OrderedDocument,
  id: string,
  suggestions: readonly string[],
  status: 'accepted' | 'rejected'
) {
  return withUndoRecord(doc, () =>
    applyOrderedChange(
      doc,
      { v: 1, id, kind: 'decide', cards: orderedSeenCards(doc, suggestions), status },
      { person: testPerson('Olive'), date: '2026-10-02T00:00:00Z', ids: orderedIdMinter(id) }
    )
  );
}

describe('change records outside the stories', () => {
  test('each record group of the styles part is one card, after the story cards', () => {
    const { document: doc } = opened();
    const records = [...doc.partRevisions.values()].map(
      ({ element, styleId, revision }) => `${element} ${revision.id} ${styleId ?? '(defaults)'}`
    );
    expect(records).toEqual([
      'rPrChange 10 (defaults)',
      'rPrChange 1 Normal',
      'pPrChange 7 Quote',
      'tblPrChange 8 TableGrid',
      'tcPrChange 9 TableGrid',
    ]);
    const cards = orderedCards(doc);
    expect(cards).toHaveLength(7);
    // The story cards keep the ids they had, and each record is a card of its own
    expect(cards.slice(2).map((card) => card.suggestions)).toEqual(
      [...doc.partRevisions.keys()].map((id) => [id])
    );
    for (const id of doc.partRevisions.keys())
      expect(orderedCardOf(doc, id)?.suggestions).toEqual([id]);
    // Every copy derives the same ids from the Word file
    expect([...opened().document.partRevisions.keys()]).toEqual([...doc.partRevisions.keys()]);
  });

  test("the default style's record is not the body insertion with the same w:id", () => {
    const { document: doc } = opened();
    const normal = recordOf(doc, 'rPrChange', '1');
    const body = orderedCards(doc)[0]!.suggestions[0]!;
    expect(normal).not.toBe(body);
    expect(doc.suggestions.get(normal)?.attribution).toEqual(
      doc.suggestions.get(body)?.attribution
    );
  });

  test('a record the engine cannot decide both ways makes no suggestion', () => {
    // No author (the editor alone shows it read-only), and no old properties (no reject)
    const styles =
      `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:styleId="Quote">` +
      '<w:name w:val="Quote"/><w:pPr><w:jc w:val="center"/>' +
      '<w:pPrChange w:id="5"><w:pPr/></w:pPrChange></w:pPr><w:rPr><w:i/>' +
      '<w:rPrChange w:id="6" w:author="Ana"/></w:rPr></w:style></w:styles>';
    const bytes = docx(p(t('Plain.')), { parts: { 'word/styles.xml': styles } });
    const { document: doc } = openOrderedWordFile(bytes, { file: 1 });
    expect(doc.partRevisions.size).toBe(0);
    expect(orderedCards(doc)).toEqual([]);
  });

  for (const status of ['accepted', 'rejected'] as const)
    test(`a ${status} record is resolved in the save; its take-back gives the file's back`, () => {
      const { document: doc, base } = opened();
      const before = savedStyles(doc, base);
      expect(before).toBe(serializeOoxmlPart(base.parts.get(STYLES)!));
      const quote = recordOf(doc, 'pPrChange', '7');
      const { record } = decide(doc, 'olive:1', [quote], status);
      expect(doc.suggestions.get(quote)?.status).toBe(status);
      const after = savedStyles(doc, base);
      expect(after).not.toContain('<w:pPrChange ');
      expect(after).toContain(status === 'accepted' ? 'w:after="300"' : 'w:after="100"');
      // The other records stay pending, as the file wrote them
      expect(after.match(/<w:[a-zA-Z]+PrChange /g)).toHaveLength(4);
      expect(orderedCards(doc)).toHaveLength(6);
      takeBack(doc, record);
      expect(savedStyles(doc, base)).toBe(before);
      expect(orderedCards(doc)).toHaveLength(7);
    });

  test('decisions give one styles part in every order', () => {
    const first = opened().document;
    const second = opened().document;
    const ids = [...first.partRevisions.keys()];
    const accepted = ids.filter((_, k) => k % 2 === 0);
    const rejected = ids.filter((_, k) => k % 2 === 1);
    decide(first, 'olive:1', accepted, 'accepted');
    decide(first, 'olive:2', rejected, 'rejected');
    decide(second, 'olive:2', rejected, 'rejected');
    decide(second, 'olive:1', accepted, 'accepted');
    const saved = savedStyles(first);
    expect(saved).not.toContain('PrChange');
    expect(savedStyles(second)).toBe(saved);
    expect(orderedCards(first).map((card) => card.suggestions)).toEqual(
      orderedCards(second).map((card) => card.suggestions)
    );
  });
});
