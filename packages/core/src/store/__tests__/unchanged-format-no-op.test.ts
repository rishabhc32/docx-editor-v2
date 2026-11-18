// A format that restates values the text already has changes nothing, so it writes nothing.
//
// Word makes no change for such a press. The appliers used to rewrite the container in schema
// order and split the run at the range edges, so the saved package moved while the words did
// not, and a collaborator saw a change nobody made.

import { describe, expect, test } from 'bun:test';
import { readOoxmlPart, type OoxmlNode, type OoxmlPart } from '../package/ooxml-tree.ts';
import { serializeOoxmlPart } from '../package/ooxml-serialize.ts';
import { TreeDocumentStore } from '../store/tree-store.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function load(body: string): OoxmlPart {
  const result = readOoxmlPart(`<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`, {
    name: '/word/document.xml',
    contentType: 'app/xml',
  });
  if (!result.ok) throw new Error(result.reason);
  return result.part;
}

function firstParagraphId(part: OoxmlPart): string {
  let found: string | null = null;
  const walk = (node: OoxmlNode): void => {
    if (found || node.kind === 'textValue') return;
    if (node.kind === 'paragraph') found = node.id;
    for (const child of node.children) walk(child);
  };
  walk(part.root);
  if (!found) throw new Error('no paragraph');
  return found;
}

function runCount(part: OoxmlPart): number {
  let count = 0;
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'run') count += 1;
    for (const child of node.children) walk(child);
  };
  walk(part.root);
  return count;
}

/** Apply one op and report whether the store moved at all. */
function applyOne(body: string, op: (paragraphId: string) => TreeDocOp) {
  const store = new TreeDocumentStore(load(body));
  const before = store.part;
  const revision = store.revision;
  const xml = serializeOoxmlPart(before);
  const result = store.transact((context) => {
    context.apply(op(firstParagraphId(before)));
  });
  if (!result.ok) throw new Error(result.detail ?? result.reason);
  return {
    result,
    unchanged: store.part === before && store.revision === revision && !store.canUndo,
    sameXml: serializeOoxmlPart(store.part) === xml,
    runs: runCount(store.part),
  };
}

// `w:b` precedes `w:sz` in `CT_RPr`; authored out of order, so a rewrite would reorder it.
const RUN = `<w:p><w:r><w:rPr><w:sz w:val="28"/><w:b/></w:rPr><w:t>Hello world</w:t></w:r></w:p>`;
const MARK = `<w:p><w:pPr><w:rPr><w:sz w:val="28"/><w:b/></w:rPr></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`;
// `w:pStyle` precedes `w:jc` in `CT_PPr`.
const PARAGRAPH = `<w:p><w:pPr><w:jc w:val="center"/><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Hi</w:t></w:r></w:p>`;
const SAME_RUN = [{ localName: 'sz', attributes: { val: '28' } }, { localName: 'b' }];

describe('a format that changes no value makes no change', () => {
  test('run: the same size and bold over part of the run', () => {
    const outcome = applyOne(RUN, (paragraphId) => ({
      op: 'setRunProperties',
      paragraphId,
      start: 2,
      end: 7,
      properties: SAME_RUN,
    }));
    expect(outcome.result.ok && outcome.result.change).toBeNull();
    expect(outcome.unchanged).toBe(true);
    expect(outcome.sameXml).toBe(true);
    expect(outcome.runs).toBe(1);
  });

  test('run in suggesting mode: no change and no format record', () => {
    const outcome = applyOne(RUN, (paragraphId) => ({
      op: 'setRunProperties',
      paragraphId,
      start: 2,
      end: 7,
      properties: SAME_RUN,
      revision: { author: 'Ada' },
    }));
    expect(outcome.unchanged).toBe(true);
    expect(outcome.sameXml).toBe(true);
  });

  test('clearing a run that has no direct formatting', () => {
    const outcome = applyOne(`<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>`, (paragraphId) => ({
      op: 'setRunProperties',
      paragraphId,
      start: 2,
      end: 7,
      properties: [],
    }));
    expect(outcome.unchanged).toBe(true);
    expect(outcome.runs).toBe(1);
  });

  test('paragraph mark: the same size and bold', () => {
    const outcome = applyOne(MARK, (paragraphId) => ({
      op: 'setParagraphMarkProperties',
      paragraphId,
      properties: SAME_RUN,
    }));
    expect(outcome.unchanged).toBe(true);
    expect(outcome.sameXml).toBe(true);
  });

  test('paragraph: the same style and alignment', () => {
    const outcome = applyOne(PARAGRAPH, (paragraphId) => ({
      op: 'setParagraphProperties',
      paragraphId,
      properties: [
        { localName: 'pStyle', attributes: { val: 'Heading1' } },
        { localName: 'jc', attributes: { val: 'center' } },
      ],
    }));
    expect(outcome.unchanged).toBe(true);
    expect(outcome.sameXml).toBe(true);
  });
});

describe('a format that changes a value still writes it', () => {
  test('run: a new size splits the run and changes it', () => {
    const outcome = applyOne(RUN, (paragraphId) => ({
      op: 'setRunProperties',
      paragraphId,
      start: 2,
      end: 7,
      properties: [{ localName: 'sz', attributes: { val: '32' } }, { localName: 'b' }],
    }));
    expect(outcome.unchanged).toBe(false);
    expect(outcome.runs).toBe(3);
  });

  test('run: dropping one property is a change', () => {
    const outcome = applyOne(RUN, (paragraphId) => ({
      op: 'setRunProperties',
      paragraphId,
      start: 0,
      end: 11,
      properties: [{ localName: 'sz', attributes: { val: '28' } }],
    }));
    expect(outcome.unchanged).toBe(false);
  });

  test('paragraph mark: a new value', () => {
    const outcome = applyOne(MARK, (paragraphId) => ({
      op: 'setParagraphMarkProperties',
      paragraphId,
      properties: [{ localName: 'sz', attributes: { val: '28' } }, { localName: 'i' }],
    }));
    expect(outcome.unchanged).toBe(false);
  });

  test('paragraph: a new alignment', () => {
    const outcome = applyOne(PARAGRAPH, (paragraphId) => ({
      op: 'setParagraphProperties',
      paragraphId,
      properties: [
        { localName: 'pStyle', attributes: { val: 'Heading1' } },
        { localName: 'jc', attributes: { val: 'right' } },
      ],
    }));
    expect(outcome.unchanged).toBe(false);
  });
});
