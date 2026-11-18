import { describe, expect, test } from 'bun:test';
import {
  applyTreeOp,
  canonicalOoxmlFingerprint,
  paragraphTextOf,
  projectOmmlEquation,
  readOoxmlPart,
  segmentsOf,
  serializeOoxmlPart,
  type OoxmlElement,
  type OoxmlPart,
} from '../index.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const metadata = {
  name: '/word/document.xml',
  contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
};

function parse(content: string): OoxmlPart {
  const loaded = readOoxmlPart(
    `<w:document xmlns:w="${W}" xmlns:m="${M}"><w:body><w:p>${content}</w:p></w:body></w:document>`,
    metadata
  );
  if (!loaded.ok) throw new Error(loaded.reason);
  return loaded.part;
}

function paragraphOf(part: OoxmlPart): OoxmlElement {
  const body = part.root.children[0] as OoxmlElement;
  const paragraph = body.children.find((child) => child.kind === 'paragraph');
  if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('missing paragraph');
  return paragraph;
}

const equation = '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>';

describe('OMML equation model atom', () => {
  test('occupies one paragraph offset and round-trips without normalization', () => {
    const part = parse(`<w:r><w:t>A</w:t></w:r>${equation}<w:r><w:t>Z</w:t></w:r>`);
    const paragraph = paragraphOf(part);
    const segment = segmentsOf(paragraph).find(
      (candidate) =>
        candidate.node.kind !== 'textValue' &&
        candidate.node.namespaceUri === M &&
        candidate.node.localName === 'oMath'
    );

    expect(segment).toMatchObject({ start: 1, end: 2 });
    expect(segment?.removeNodeIds).toEqual([segment?.node.id]);
    expect(paragraphTextOf(part, paragraph.id)).toBe('A\uFFFCZ');

    const reopened = readOoxmlPart(serializeOoxmlPart(part), metadata);
    if (!reopened.ok) throw new Error(reopened.reason);
    expect(canonicalOoxmlFingerprint(reopened.part)).toBe(canonicalOoxmlFingerprint(part));
  });

  test('deletes the complete equation through the shared text operation', () => {
    const part = parse(`<w:r><w:t>A</w:t></w:r>${equation}<w:r><w:t>Z</w:t></w:r>`);
    const paragraph = paragraphOf(part);
    const result = applyTreeOp(part, {
      op: 'deleteText',
      paragraphId: paragraph.id,
      start: 1,
      end: 2,
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(paragraphTextOf(result.part, paragraph.id)).toBe('AZ');
    expect(serializeOoxmlPart(result.part)).not.toContain('<m:oMath>');
  });

  test('a tracked deletion strikes the complete equation with the text around it', () => {
    // `w:del` may hold `m:oMath` (`EG_RunLevelElts` → `EG_MathContent`). Leaving the equation
    // outside the strike made Accept remove the text and keep the equation.
    const part = parse(`<w:r><w:t>AB</w:t></w:r>${equation}<w:r><w:t>YZ</w:t></w:r>`);
    const paragraph = paragraphOf(part);
    const revision = { author: 'Ada', date: '2026-01-02T03:04:05Z' };
    const struck = applyTreeOp(part, {
      op: 'deleteText',
      paragraphId: paragraph.id,
      start: 1,
      end: 4,
      revision,
    });
    if (!struck.ok) throw new Error(struck.reason);
    expect(serializeOoxmlPart(struck.part)).toContain(
      '<w:r><w:delText>B</w:delText></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>' +
        '<w:r><w:delText>Y</w:delText></w:r></w:del>'
    );
    expect(paragraphTextOf(struck.part, paragraph.id)).toBe('AB\uFFFCYZ');

    const accepted = applyTreeOp(struck.part, { op: 'acceptAllRevisions' });
    const rejected = applyTreeOp(struck.part, { op: 'rejectAllRevisions' });
    if (!accepted.ok || !rejected.ok) throw new Error('revision resolution refused');
    expect(paragraphTextOf(accepted.part, paragraph.id)).toBe('AZ');
    expect(serializeOoxmlPart(accepted.part)).not.toContain('<m:oMath>');
    expect(paragraphTextOf(rejected.part, paragraph.id)).toBe('AB\uFFFCYZ');
    expect(serializeOoxmlPart(rejected.part)).not.toContain('<w:del ');
  });

  test('typing at a leading equation boundary creates a sibling run', () => {
    const part = parse(`${equation}<w:r><w:t>Z</w:t></w:r>`);
    const paragraph = paragraphOf(part);
    const result = applyTreeOp(part, {
      op: 'insertText',
      paragraphId: paragraph.id,
      offset: 0,
      text: 'A',
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(paragraphTextOf(result.part, paragraph.id)).toBe('A\uFFFCZ');
  });

  test('typing after a trailing equation creates a sibling run after it', () => {
    const part = parse(`<w:r><w:t>A</w:t></w:r>${equation}`);
    const paragraph = paragraphOf(part);
    const result = applyTreeOp(part, {
      op: 'insertText',
      paragraphId: paragraph.id,
      offset: 2,
      text: 'Z',
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(paragraphTextOf(result.part, paragraph.id)).toBe('A\uFFFCZ');
    expect(serializeOoxmlPart(result.part)).toContain('</m:oMath><w:r><w:t>Z</w:t></w:r>');
  });

  test('replaces and removes an equation through explicit atomic operations', () => {
    const part = parse(`<w:r><w:t>A</w:t></w:r>${equation}<w:r><w:t>Z</w:t></w:r>`);
    const paragraph = paragraphOf(part);
    const original = paragraph.children.find(
      (child) =>
        child.kind !== 'textValue' && child.namespaceUri === M && child.localName === 'oMath'
    );
    if (!original) throw new Error('missing equation');

    const replaced = applyTreeOp(part, {
      op: 'setMathEquation',
      equationId: original.id,
      linear: '{a+b}/{2}',
    });
    expect(replaced.ok).toBe(true);
    if (!replaced.ok) return;
    const next = projectOmmlEquation(
      paragraphOf(replaced.part).children.find((child) => child.id === original.id)!
    );
    expect(next?.expression.kind).toBe('fraction');
    expect(paragraphTextOf(replaced.part, paragraph.id)).toBe('A\uFFFCZ');

    const removed = applyTreeOp(replaced.part, {
      op: 'removeMathEquation',
      equationId: original.id,
    });
    expect(removed.ok).toBe(true);
    if (!removed.ok) return;
    expect(paragraphTextOf(removed.part, paragraph.id)).toBe('AZ');
  });

  test('refuses malformed linear math without changing the part', () => {
    const part = parse(equation);
    const paragraph = paragraphOf(part);
    const original = paragraph.children.find(
      (child) =>
        child.kind !== 'textValue' && child.namespaceUri === M && child.localName === 'oMath'
    );
    if (!original) throw new Error('missing equation');
    const result = applyTreeOp(part, {
      op: 'setMathEquation',
      equationId: original.id,
      linear: 'x^',
    });
    expect(result).toEqual({ ok: false, reason: 'invalid-property-value' });
    expect(canonicalOoxmlFingerprint(part)).toBe(canonicalOoxmlFingerprint(parse(equation)));
  });
});
