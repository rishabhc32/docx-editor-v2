// A file's pending revisions load as suggestions and compose back to the same part. Deciding a
// suggestion (a status change) composes to what the engine's own accept or reject gives on the
// same file, and decisions give one result in every order.

import { describe, expect, test } from 'bun:test';
import { canonicalOoxmlFingerprint, type OoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyTreeOp } from '../../../store/store/index.ts';
import {
  composeOrderedPart,
  loadOrderedPackage,
  setOrderedStatus,
  withOrderedUndoRecord,
  type OrderedDocument,
  type OrderedSuggestionStatus,
} from '../index.ts';
import { expectSameFingerprint, expectValid } from './oracle.ts';
import { REVISION_BODIES } from './revision-fixtures.ts';
import { docx, packageOf } from './word-fixtures.ts';

const BODY = '/word/document.xml';

function fixture(body: string): {
  doc: OrderedDocument;
  part: OoxmlPart;
  load: () => OrderedDocument;
} {
  const pkg = packageOf(docx(body));
  const part = pkg.parts.get(BODY);
  if (!part) throw new Error('no body part');
  return { doc: loadOrderedPackage(pkg), part, load: () => loadOrderedPackage(pkg) };
}

/** The file's revision triple a suggestion was loaded from (its first one). */
function tripleOf(doc: OrderedDocument, suggestion: string) {
  const attribution = doc.suggestions.get(suggestion)?.attribution;
  const id = [...doc.revisionIds.byRole].find(([key]) => key.startsWith(`${suggestion}/`))?.[1];
  if (!attribution || id === undefined) throw new Error(`no triple for ${suggestion}`);
  return {
    id,
    author: attribution.author,
    ...(attribution.date === undefined ? {} : { date: attribution.date }),
  };
}

function decide(
  doc: OrderedDocument,
  decisions: readonly (readonly [string, OrderedSuggestionStatus])[]
) {
  for (const [suggestion, status] of decisions)
    withOrderedUndoRecord(doc, () => setOrderedStatus(doc, suggestion, status));
  return composeOrderedPart(doc, BODY).part;
}

describe('pending revisions from a file', () => {
  for (const [name, body] of Object.entries(REVISION_BODIES)) {
    test(`${name}: load as suggestions and compose back to the same part`, () => {
      const { doc, part } = fixture(body);
      expect(doc.suggestions.size).toBeGreaterThan(0);
      for (const rebuild of [false, true]) {
        const composed = composeOrderedPart(doc, BODY, { rebuild }).part;
        expectSameFingerprint(composed, part, `${name} rebuild=${rebuild}`);
        expectValid(composed, name);
      }
    });

    test(`${name}: each decision gives the engine's result`, () => {
      const { doc, part, load } = fixture(body);
      for (const suggestion of doc.suggestions.keys()) {
        for (const action of ['accept', 'reject'] as const) {
          const engine = applyTreeOp(part, {
            op: action === 'accept' ? 'acceptRevision' : 'rejectRevision',
            revision: tripleOf(doc, suggestion),
          });
          expect({ name, suggestion, action, ok: engine.ok }).toEqual({
            name,
            suggestion,
            action,
            ok: true,
          });
          if (!engine.ok) continue;
          const model = decide(load(), [
            [suggestion, action === 'accept' ? 'accepted' : 'rejected'],
          ]);
          expectSameFingerprint(model, engine.part, `${name} ${action} ${suggestion}`);
          expectValid(model, `${name} ${action} ${suggestion}`);
        }
      }
    });
  }

  test('two decisions give one result in either order', () => {
    for (const name of ['two authors', 'deletion inside an insertion']) {
      const body = REVISION_BODIES[name];
      if (body === undefined) throw new Error(name);
      const { load } = fixture(body);
      const [a, b] = [...load().suggestions.keys()];
      if (a === undefined || b === undefined) throw new Error(`${name} needs two suggestions`);
      for (const first of ['accepted', 'rejected'] as const)
        for (const second of ['accepted', 'rejected'] as const) {
          const one = decide(load(), [
            [a, first],
            [b, second],
          ]);
          const other = decide(load(), [
            [b, second],
            [a, first],
          ]);
          expect(canonicalOoxmlFingerprint(one)).toBe(canonicalOoxmlFingerprint(other));
        }
    }
  });

  test('a decision keeps every item and its address', () => {
    const body = REVISION_BODIES.insertion;
    if (body === undefined) throw new Error('insertion');
    const { doc } = fixture(body);
    const addresses = () =>
      [...doc.stories.values()].flatMap((story) =>
        story.tree.toArray().map((item) => `${item.space}|${item.offset}+${item.length}`)
      );
    const before = addresses();
    decide(doc, [['S1', 'accepted']]);
    expect(addresses()).toEqual(before);
    decide(doc, [['S1', 'rejected']]);
    expect(addresses()).toEqual(before);
  });
});
