// Splitting a paragraph whose mark carries a pending revision.
//
// A `w:pPr/w:rPr/w:del` deletes the mark that ENDS the paragraph (§17.13.5). Word splits by
// inserting a fresh mark before that one, so the struck mark stays at the end of the tail and
// the head ends with a mark nobody proposed removing. The store copied the whole `w:pPr` onto
// the head, deletion included, so Accept all took back the split along with the join.

import { describe, expect, test } from 'bun:test';
import type { OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { apply, part, xml } from './tracked-edit-fixture.ts';

const BEN = { author: 'Ben', date: '2026-10-18T00:00:00Z' };
const CY = { author: 'Cy', date: '2026-10-18T00:01:00Z' };
const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

function paragraphs(source: OoxmlPart): OoxmlNode[] {
  const out: OoxmlNode[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'paragraph') out.push(node);
    else if (node.kind !== 'textValue') node.children.forEach(visit);
  };
  visit(source.root);
  return out;
}

/** Each paragraph's text and own mark revisions, as `text|kind:author,…`. */
function summary(source: OoxmlPart): string[] {
  const text = (node: OoxmlNode): string =>
    node.kind === 'textValue' ? node.value : node.children.map(text).join('');
  const named = (node: OoxmlNode | undefined, localName: string) =>
    node && node.kind !== 'textValue'
      ? node.children.find((child) => child.kind !== 'textValue' && child.localName === localName)
      : undefined;
  return paragraphs(source).map((paragraph) => {
    const rPr = named(named(paragraph, 'pPr'), 'rPr');
    const marks = (rPr && rPr.kind !== 'textValue' ? rPr.children : []).flatMap((mark) =>
      mark.kind !== 'textValue' && (mark.localName === 'ins' || mark.localName === 'del')
        ? [`${mark.localName}:${mark.attributes.find((a) => a.localName === 'author')?.value}`]
        : []
    );
    return `${text(paragraph)}|${marks.join(',')}`;
  });
}

/** Ben proposes joining "Heading" to "Next": the store strikes the first mark. */
function struck(): OoxmlPart {
  const before = part(plain('Heading') + plain('Next'));
  return apply(before, {
    op: 'proposeParagraphMerge',
    paragraphId: paragraphs(before)[1]!.id,
    revision: BEN,
  });
}

describe('a split inside a paragraph whose mark another author struck', () => {
  test('an untracked split leaves the deletion on the tail only', () => {
    const source = struck();
    const split = apply(source, {
      op: 'splitParagraph',
      paragraphId: paragraphs(source)[0]!.id,
      offset: 1,
    });
    expect(summary(split)).toEqual(['H|', 'eading|del:Ben', 'Next|']);
    expect(summary(apply(split, { op: 'acceptAllRevisions' }))).toEqual(['H|', 'eadingNext|']);
    // The containers Ben's proposal wrote go with it, on the tail as on the unsplit paragraph.
    const rejected = apply(split, { op: 'rejectAllRevisions' });
    expect(summary(rejected)).toEqual(['H|', 'eading|', 'Next|']);
    expect(xml(rejected)).not.toContain('<w:pPr');
  });

  test('a multi-line paste gives every earlier piece a new mark, and the last the struck one', () => {
    const source = struck();
    const pasted = apply(source, {
      op: 'splitParagraphMany',
      paragraphId: paragraphs(source)[0]!.id,
      offsets: [1, 3],
      revision: CY,
    });
    expect(summary(pasted)).toEqual(['H|ins:Cy', 'ea|ins:Cy', 'ding|del:Ben', 'Next|']);
    expect(summary(apply(pasted, { op: 'acceptAllRevisions' }))).toEqual([
      'H|',
      'ea|',
      'dingNext|',
    ]);
    expect(xml(apply(pasted, { op: 'rejectAllRevisions' }))).toBe(
      xml(
        part(
          '<w:p><w:r><w:t>H</w:t></w:r><w:r><w:t>ea</w:t></w:r><w:r><w:t>ding</w:t></w:r></w:p>' +
            plain('Next')
        )
      )
    );
  });

  test('the many-way split still equals the single splits it stands for', () => {
    const source = struck();
    const id = paragraphs(source)[0]!.id;
    const many = apply(source, { op: 'splitParagraphMany', paragraphId: id, offsets: [1, 3] });
    const singles = apply(apply(source, { op: 'splitParagraph', paragraphId: id, offset: 3 }), {
      op: 'splitParagraph',
      paragraphId: id,
      offset: 1,
    });
    expect(xml(many)).toBe(xml(singles));
  });

  test('a deletion written by Word keeps its source container on the head', () => {
    const source = part(
      '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/>' +
        '<w:del w:id="5" w:author="Ben" w:date="2026-10-18T00:00:00Z"/></w:rPr></w:pPr>' +
        '<w:r><w:t>Heading</w:t></w:r></w:p>' +
        plain('Next')
    );
    const split = apply(source, {
      op: 'splitParagraph',
      paragraphId: paragraphs(source)[0]!.id,
      offset: 1,
    });
    // Only the deletion leaves the head; the alignment and the mark's bold are its own too.
    expect(xml(split)).toContain(
      '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>H</w:t>'
    );
    expect(summary(split)).toEqual(['H|', 'eading|del:Ben', 'Next|']);
  });
});

describe('a split inside a paragraph whose mark is a pending insertion', () => {
  test('both halves keep the insertion, as new text inside an insertion does', () => {
    const before = part(plain('Heading') + plain('Next'));
    const stamped = apply(before, {
      op: 'setParagraphMarkRevision',
      paragraphId: paragraphs(before)[0]!.id,
      kind: 'ins',
      revision: BEN,
    });
    const split = apply(stamped, {
      op: 'splitParagraph',
      paragraphId: paragraphs(stamped)[0]!.id,
      offset: 1,
    });
    expect(summary(split)).toEqual(['H|ins:Ben', 'eading|ins:Ben', 'Next|']);
  });
});
