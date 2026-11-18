// A loaded part is untrusted: the load refuses by name what passes its limits or is not a
// story, and new revision ids never repeat an id the file holds.

import { describe, expect, test } from 'bun:test';
import { revisionIdFor } from '../model/revision-ids.ts';
import {
  loadOrderedDocument,
  loadOrderedPackage,
  newOrderedSuggestion,
  OrderedRefusal,
  type OrderedRefusalReason,
} from '../index.ts';
import { REVISION_BODIES } from './revision-fixtures.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';
import { testAttribution } from './test-persons.ts';

function refusalOf(load: () => unknown): OrderedRefusalReason | null {
  try {
    load();
    return null;
  } catch (error) {
    if (error instanceof OrderedRefusal) return error.reason;
    throw error;
  }
}

const nestedTables = (depth: number): string => {
  let inner = p(t('deep'));
  for (let k = 0; k < depth; k += 1)
    inner =
      '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>' +
      `<w:tr><w:tc>${inner}</w:tc></w:tr></w:tbl>${p('')}`;
  return inner;
};

describe('ordered load limits', () => {
  test('refuses more items than the limit', () => {
    const pkg = packageOf(docx(p(t('one') + t('two') + t('three') + t('four'))));
    expect(refusalOf(() => loadOrderedPackage(pkg, { maxItems: 3 }))).toBe('too-many-items');
    expect(refusalOf(() => loadOrderedPackage(pkg, { maxItems: 100 }))).toBeNull();
  });

  test('refuses nesting past the depth limit', () => {
    const pkg = packageOf(docx(nestedTables(6)));
    expect(refusalOf(() => loadOrderedPackage(pkg, { maxDepth: 8 }))).toBe('too-deep');
    expect(refusalOf(() => loadOrderedPackage(pkg))).toBeNull();
  });

  test('a limit can be lowered, never raised past the engine ceiling', () => {
    const pkg = packageOf(docx(p(t('one'))));
    expect(
      refusalOf(() => loadOrderedPackage(pkg, { maxItems: Number.POSITIVE_INFINITY }))
    ).toBeNull();
    expect(refusalOf(() => loadOrderedPackage(pkg, { maxDepth: -1 }))).toBeNull();
  });

  test('refuses a part that is not a story', () => {
    const pkg = packageOf(docx(p(t('one'))));
    const styles = pkg.parts.get('/word/styles.xml');
    if (!styles) throw new Error('no styles');
    expect(refusalOf(() => loadOrderedDocument(new Map([['/word/styles.xml', styles]])))).toBe(
      'malformed-story'
    );
  });

  test('a new revision id goes past every id the file holds', () => {
    const body = REVISION_BODIES['two authors'];
    if (body === undefined) throw new Error('two authors');
    const doc = loadOrderedPackage(packageOf(docx(body)));
    const fresh = newOrderedSuggestion(doc, testAttribution('Dan'));
    const id = Number(revisionIdFor(doc, fresh.id, 'ins'));
    expect(id).toBeGreaterThan(918);
    // The same suggestion and role keep their id
    expect(revisionIdFor(doc, fresh.id, 'ins')).toBe(String(id));
  });
});
