// `removeHyperlink` with a `styleId`: the link's character style comes off its words with it.

import { describe, expect, test } from 'bun:test';
import {
  readOoxmlPart,
  validateOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { applyTreeOp, type TreeDocOp } from '../store/tree-ops.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

/** The first link's node id. */
function linkIdOf(part: OoxmlPart): string {
  const walk = (node: OoxmlNode): string | null => {
    if (node.kind === 'textValue') return null;
    if (node.kind === 'hyperlink') return node.id;
    for (const child of node.children) {
      const found = walk(child);
      if (found) return found;
    }
    return null;
  };
  const id = walk(part.root);
  if (!id) throw new Error('no link');
  return id;
}

describe('removeHyperlink with a styleId', () => {
  const style = (id: string) => `<w:rStyle w:val="${id}"/>`;
  const part = load(
    '<w:p><w:hyperlink w:anchor="x">' +
      `<w:r><w:rPr>${style('Hyperlink')}<w:b/></w:rPr><w:t>bold</w:t></w:r>` +
      `<w:ins w:author="A" w:id="1"><w:r><w:rPr>${style('Hyperlink')}</w:rPr><w:t>new</w:t></w:r></w:ins>` +
      `<w:r><w:rPr>${style('Strong')}</w:rPr><w:t>strong</w:t></w:r></w:hyperlink></w:p>`
  );

  test('takes that run style off, through revisions, and nothing else', () => {
    const op: TreeDocOp = { op: 'removeHyperlink', linkId: linkIdOf(part), styleId: 'Hyperlink' };
    const result = applyTreeOp(part, op);
    if (!result.ok) throw new Error(result.reason);
    expect(validateOoxmlPart(result.part).ok).toBe(true);
    expect(serializeOoxmlPart(result.part)).toContain(
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>' +
        '<w:ins w:author="A" w:id="1"><w:r><w:t>new</w:t></w:r></w:ins>' +
        `<w:r><w:rPr>${style('Strong')}</w:rPr><w:t>strong</w:t></w:r></w:p>`
    );
  });

  test('without one, the runs keep it; an empty one is refused', () => {
    const linkId = linkIdOf(part);
    const kept = applyTreeOp(part, { op: 'removeHyperlink', linkId });
    if (!kept.ok) throw new Error(kept.reason);
    expect(serializeOoxmlPart(kept.part).match(/w:rStyle w:val="Hyperlink"/g)).toHaveLength(2);
    const refused = applyTreeOp(part, { op: 'removeHyperlink', linkId, styleId: '' });
    expect(refused.ok ? null : refused.reason).toBe('invalid-property-value');
  });
});
