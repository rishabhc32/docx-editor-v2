// Before a save, a composed part holds each node id once: a node placed twice is copied under new
// ids, and the part's text and shape stay as they were.

import { describe, expect, test } from 'bun:test';
import { serializeOoxmlPart, validateOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { withUniqueIds } from '../compose/unique-ids.ts';
import { withChildren } from '../ooxml-nodes.ts';
import { docx, p, packageOf, t } from './word-fixtures.ts';

describe('unique node ids before a save', () => {
  test('a node placed twice is copied under new ids', () => {
    const part = packageOf(docx(p(t('Alpha.')) + p(t('Beta.')))).parts.get('/word/document.xml');
    if (!part) throw new Error('fixture');
    const body = part.root.children.find((c) => c.kind !== 'textValue' && c.localName === 'body');
    if (!body || body.kind === 'textValue') throw new Error('fixture');
    const [first] = body.children;
    if (!first) throw new Error('fixture');
    // The first paragraph twice: one object, one set of ids
    const twice = withChildren(body, [first, ...body.children]);
    const doubled = {
      ...part,
      root: withChildren(
        part.root,
        part.root.children.map((c) => (c === body ? twice : c))
      ),
    };
    expect(validateOoxmlPart(doubled).ok).toBe(false);
    const fixed = withUniqueIds(doubled);
    expect(validateOoxmlPart(fixed).ok).toBe(true);
    expect(serializeOoxmlPart(fixed)).toBe(serializeOoxmlPart(doubled));
    // A part that holds each id once is the same part
    expect(withUniqueIds(part)).toBe(part);
  });
});
