// One paragraph composed alone equals that paragraph in the whole story's compose, and keeps
// its identity: the loaded node id and `w14:paraId`, whether it is reused or drawn again.

import { describe, expect, test } from 'bun:test';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import { canonicalOoxmlFingerprint, type OoxmlNode } from '../../../store/package/ooxml-tree.ts';
import {
  applyOrderedChange,
  composeOrderedParagraph,
  composeOrderedStory,
  loadOrderedPackage,
  orderedCardSeen,
  orderedIdMinter,
  orderedParagraphCount,
  setOrderedStatus,
  withOrderedUndoRecord,
  type OrderedDocument,
  type OrderedItem,
} from '../index.ts';
import { REVISION_BODIES } from './revision-fixtures.ts';
import {
  bodyOf,
  docx,
  MATRIX_EXTRAS,
  MATRIX_ITEMS,
  p,
  packageOf,
  R,
  t,
  W,
} from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const BODY = '/word/document.xml';
const ENDNOTES = '/word/endnotes.xml';

const DELETED = 'w:id="7" w:author="Cara" w:date="2026-01-01T00:00:00Z"';

/**
 * Two endnotes. Cara's deletion strikes the first one's reference, mark and text, as Word writes
 * a deleted note; the bookmark in it stays out of the deletion.
 */
const DELETED_NOTE = docx(
  p(
    t('The fee') +
      `<w:del ${DELETED}><w:r><w:endnoteReference w:id="1"/></w:r></w:del>` +
      t(' is due') +
      '<w:r><w:endnoteReference w:id="2"/></w:r>'
  ),
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/endnotes.xml':
        `<w:endnotes xmlns:w="${W}" xmlns:r="${R}">` +
        `<w:endnote w:id="1"><w:p><w:pPr><w:rPr><w:del ${DELETED}/></w:rPr></w:pPr>` +
        '<w:bookmarkStart w:id="0" w:name="annex"/>' +
        `<w:del ${DELETED}><w:r><w:endnoteRef/></w:r><w:r><w:delText> See the annex.</w:delText></w:r></w:del>` +
        '<w:bookmarkEnd w:id="0"/></w:p></w:endnote>' +
        `<w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r>${t(' Net of tax.')}</w:p></w:endnote>` +
        '</w:endnotes>',
    },
  }
);

const CARA = 'w:id="5" w:author="Cara" w:date="2026-01-01T00:00:00Z"';

/** A block control of two paragraphs of Cara's pending words, and an empty paragraph after it. */
const BLOCK_CONTROL = docx(
  p(t('Before.')) +
    '<w:sdt><w:sdtPr><w:id w:val="602"/></w:sdtPr><w:sdtContent>' +
    p(`<w:ins ${CARA}>${t('Acme')}</w:ins>`) +
    p(`<w:ins ${CARA}>${t('Corp')}</w:ins>`) +
    '</w:sdtContent></w:sdt><w:p/>' +
    p(t('End.'))
);

/** Every paragraph node in the blocks, in document order (in tables and controls too). */
function paragraphsIn(nodes: readonly OoxmlNode[], into: OoxmlNode[] = []): OoxmlNode[] {
  for (const node of nodes) {
    if (node.kind === 'textValue') continue;
    if (node.kind === 'paragraph') into.push(node);
    else paragraphsIn(node.children, into);
  }
  return into;
}

const addressOf = (item: OrderedItem) => `${item.space}:${item.offset}`;

const paraIdOf = (node: OoxmlNode | null | undefined) =>
  node && node.kind !== 'textValue'
    ? node.attributes.find((a) => a.localName === 'paraId')?.value
    : undefined;

function expectParagraphsMatch(doc: OrderedDocument, label: string, name = BODY): void {
  const story = composeOrderedStory(doc, name);
  const whole = paragraphsIn(story.children);
  const count = orderedParagraphCount(doc, name);
  expect({ label, count }).toEqual({ label, count: story.paragraphs.length });
  expect({ label, count: whole.length }).toEqual({ label, count });
  for (let k = 0; k < count; k += 1) {
    const one = composeOrderedParagraph(doc, name, k);
    const expected = whole[k];
    if (!one || !expected) throw new Error(`${label}: no paragraph ${k}`);
    expect({ label, k, print: canonicalOoxmlFingerprint(one.node) }).toEqual({
      label,
      k,
      print: canonicalOoxmlFingerprint(expected),
    });
    expect(one.node.id).toBe(expected.id);
    // The items an edit reads to place itself
    expect({ label, k, items: one.view.items.map(addressOf) }).toEqual({
      label,
      k,
      items: story.paragraphs[k]?.items.map(addressOf),
    });
    expect(one.view.chars.map((c) => c.ch).join('')).toBe(
      story.paragraphs[k]?.chars.map((c) => c.ch).join('')
    );
  }
  expect(composeOrderedParagraph(doc, name, count)).toBeNull();
}

describe('one-paragraph compose', () => {
  for (const item of MATRIX_ITEMS) {
    test(`matrix ${item.key}`, () => {
      expectParagraphsMatch(
        loadOrderedPackage(packageOf(docx(bodyOf(item), MATRIX_EXTRAS))),
        item.key
      );
    });
  }

  for (const [name, body] of Object.entries(REVISION_BODIES)) {
    test(`${name}, pending and after each decision`, () => {
      const pkg = packageOf(docx(body));
      expectParagraphsMatch(loadOrderedPackage(pkg), `${name} pending`);
      for (const status of ['accepted', 'rejected'] as const) {
        const doc = loadOrderedPackage(pkg);
        for (const id of doc.suggestions.keys()) setOrderedStatus(doc, id, status);
        expectParagraphsMatch(doc, `${name} ${status}`);
      }
    });
  }

  // Backlog row 53: a paragraph drawn alone took the items of a left-out block before it
  test('a note after a note whose reference shows nowhere', () => {
    const pkg = packageOf(DELETED_NOTE);
    expectParagraphsMatch(loadOrderedPackage(pkg), 'pending', ENDNOTES);
    const doc = loadOrderedPackage(pkg);
    for (const id of doc.suggestions.keys()) setOrderedStatus(doc, id, 'accepted');
    expectParagraphsMatch(doc, 'accepted', ENDNOTES);
    const story = composeOrderedStory(doc, ENDNOTES);
    expect(story.paragraphs.map((v) => v.chars.map((c) => c.ch).join(''))).toEqual([
      '\ufffc Net of tax.',
    ]);
  });

  // Backlog row 54: a paragraph drawn alone took the items before a mark a control's emptying
  // dropped
  test('a paragraph after a block control a decision emptied, which drops its later paragraphs', () => {
    const doc = loadOrderedPackage(packageOf(BLOCK_CONTROL));
    const seen = orderedCardSeen(doc, 'S1');
    if (!seen) throw new Error('no card');
    const change = {
      v: 1,
      id: 'olga:1',
      kind: 'decide',
      cards: [{ suggestions: seen }],
      status: 'rejected',
    };
    withOrderedUndoRecord(doc, () =>
      applyOrderedChange(doc, change as never, {
        person: testPerson('Olga'),
        date: '2026-02-01T00:00:00Z',
        ids: orderedIdMinter('olga:1'),
      })
    );
    expectParagraphsMatch(doc, 'rejected');
    const story = composeOrderedStory(doc, BODY);
    expect(story.paragraphs.map((v) => v.chars.map((c) => c.ch).join(''))).toEqual([
      'Before.',
      '     ',
      '',
      'End.',
    ]);
  });

  test('an untouched paragraph is the loaded node; a drawn one keeps its id and paragraph id', () => {
    const body = REVISION_BODIES.insertion;
    if (body === undefined) throw new Error('insertion');
    const pkg = packageOf(docx(body));
    const doc = loadOrderedPackage(pkg);
    const loaded = pkg.parts.get(BODY);
    if (!loaded) throw new Error('no body');
    // The paragraph holds Cara's pending insertion: it is drawn from its items
    const drawn = composeOrderedParagraph(doc, BODY, 0);
    const original = findNode(loaded, drawn?.node.id ?? '');
    expect(drawn?.node).not.toBe(original);
    expect(paraIdOf(drawn?.node)).toBe(paraIdOf(original));

    const plain = loadOrderedPackage(
      packageOf(docx('<w:p w14:paraId="1A2B3C4D"><w:r><w:t>Plain</w:t></w:r></w:p>'))
    );
    const reused = composeOrderedParagraph(plain, BODY, 0);
    const story = plain.stories.get(BODY);
    const loadedNode = story ? findNode(story.part, reused?.node.id ?? '') : null;
    expect(reused?.node === loadedNode).toBe(true);
    expect(paraIdOf(reused?.node)).toBe('1A2B3C4D');
  });
});
