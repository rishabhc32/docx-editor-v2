// A content control's prompt shows only while the control holds nothing else. While a
// suggestion's text fills it, compose names a glossary placeholder block on the control and
// reports the block for the caller to write (Word's form, ECMA-376 §17.5.2.25); a decision
// that empties a control shows its prompt again, from the glossary after a save and reopen.

import { describe, expect, test } from 'bun:test';
import { findNode } from '../../../store/package/ooxml-edit.ts';
import {
  canonicalOoxmlFingerprint,
  type OoxmlElement,
  type OoxmlNode,
  type OoxmlPart,
} from '../../../store/package/ooxml-tree.ts';
import { build } from '../../../store/store/tree-op-tracked.ts';
import {
  composeOrderedPart,
  createOrderedItem,
  insertOrderedItems,
  loadOrderedDocument,
  loadOrderedPackage,
  newOrderedSuggestion,
  setOrderedStatus,
  withOrderedUndoRecord,
  type OrderedDocument,
  type OrderedItem,
} from '../index.ts';
import { expectValid } from './oracle.ts';
import { bodyOf, docx, MATRIX_ITEMS, p, packageOf, t } from './word-fixtures.ts';
import { testAttribution } from './test-persons.ts';

const BODY = '/word/document.xml';

function matrixDoc(key: string): { doc: OrderedDocument; part: OoxmlPart } {
  const item = MATRIX_ITEMS.find((candidate) => candidate.key === key);
  if (!item) throw new Error(key);
  const doc = loadOrderedPackage(packageOf(docx(bodyOf(item))));
  const part = doc.stories.get(BODY)?.part;
  if (!part) throw new Error('no body');
  return { doc, part };
}

function openItem(doc: OrderedDocument, token: string): OrderedItem {
  const found = doc.stories
    .get(BODY)
    ?.tree.toArray()
    .find((item) => item.kind === 'open' && item.token === token);
  if (!found) throw new Error(`no ${token}`);
  return found;
}

/** Ana types `text` right after `after`, as one pending insertion. Returns her suggestion. */
function anaTypes(doc: OrderedDocument, after: OrderedItem, text: string): string {
  return withOrderedUndoRecord(doc, () => {
    const story = doc.stories.get(BODY);
    if (!story) throw new Error('no body');
    const ana = newOrderedSuggestion(doc, testAttribution('Ana', '2026-02-01T00:00:00Z'));
    const typed = createOrderedItem({
      space: `i:${ana.id}`,
      kind: 'text',
      length: text.length,
      text,
      run: `i:${ana.id}`,
      by: ana.id,
    });
    insertOrderedItems(doc, story, after, [typed]);
    return ana.id;
  }).value;
}

const descendants = (node: OoxmlNode, into: OoxmlElement[] = []): OoxmlElement[] => {
  if (node.kind === 'textValue') return into;
  into.push(node);
  for (const child of node.children) descendants(child, into);
  return into;
};
const textOf = (nodes: readonly OoxmlNode[]): string =>
  nodes
    .flatMap((node) => descendants(node))
    .filter((n) => n.localName === 't')
    .map((n) => n.children.map((c) => (c.kind === 'textValue' ? c.value : '')).join(''))
    .join('');
const controlsIn = (part: OoxmlPart) => descendants(part.root).filter((n) => n.localName === 'sdt');
const named = (control: OoxmlElement, name: string) =>
  descendants(control).find((n) => n.localName === name);
const docPartOf = (control: OoxmlElement) =>
  named(control, 'docPart')?.attributes.find((a) => a.localName === 'val')?.value;

describe('content control prompts', () => {
  test('a typed suggestion hides an inline prompt: the control names a glossary block, reported for the caller', () => {
    const { doc } = matrixDoc('placeholder');
    const ana = anaTypes(doc, openItem(doc, 'sdt'), 'X');
    const { part: composed, composed: story } = composeOrderedPart(doc, BODY);
    expectValid(composed);
    const [control] = controlsIn(composed);
    if (!control) throw new Error('no control');
    expect(named(control, 'showingPlcHdr')).toBeUndefined();
    const name = docPartOf(control);
    if (name === undefined) throw new Error('the control names no block');
    expect(name).toMatch(/^[0-9A-F]{32}$/);
    expect(story.prompts.map((prompt) => ({ name: prompt.name, inline: prompt.inline }))).toEqual([
      { name, inline: true },
    ]);
    expect(textOf(story.prompts[0]?.nodes ?? [])).toBe('Click here');
    // A rejection shows the prompt again. The control keeps naming the block, as the engine
    // keeps the reference it added, so the block is still reported for the caller to keep
    withOrderedUndoRecord(doc, () => setOrderedStatus(doc, ana, 'rejected'));
    const back = composeOrderedPart(doc, BODY);
    const [again] = controlsIn(back.part);
    if (!again) throw new Error('no control after the rejection');
    expect(named(again, 'showingPlcHdr')).toBeDefined();
    expect(docPartOf(again)).toBe(name);
    expect(textOf(again.children)).toBe('Click here');
    expect(back.composed.prompts.map((prompt) => prompt.name)).toEqual([name]);
  });

  test('after a save and reopen, a decision that empties the control shows the glossary prompt', () => {
    const { doc } = matrixDoc('placeholder');
    anaTypes(doc, openItem(doc, 'sdt'), 'X');
    const { part: saved, composed } = composeOrderedPart(doc, BODY);
    const prompt = composed.prompts[0];
    if (!prompt) throw new Error('no prompt');
    // The caller wrote the block; the reopened file names it
    const block = build('glossary#p', 'paragraph', 'p', [], prompt.nodes);
    const reopened = loadOrderedDocument(new Map([[BODY, saved]]), {
      placeholders: new Map([[prompt.name, [block]]]),
    });
    const [ana] = reopened.suggestions.keys();
    if (ana === undefined) throw new Error('no suggestion after reopen');
    withOrderedUndoRecord(reopened, () => setOrderedStatus(reopened, ana, 'rejected'));
    const after = composeOrderedPart(reopened, BODY);
    expect(after.composed.emptied).toHaveLength(1);
    const [control] = controlsIn(after.part);
    if (!control) throw new Error('no control');
    expect(named(control, 'showingPlcHdr')).toBeDefined();
    expect(textOf([control])).toBe('Click here');
  });

  test('a typed suggestion hides a block prompt the same way', () => {
    const { doc } = matrixDoc('blockprompt');
    anaTypes(doc, openItem(doc, 'sdtBlock'), 'New ');
    const { part: composed, composed: story } = composeOrderedPart(doc, BODY);
    expectValid(composed);
    const [control] = controlsIn(composed);
    if (!control) throw new Error('no control');
    expect(named(control, 'showingPlcHdr')).toBeUndefined();
    const name = docPartOf(control);
    if (name === undefined) throw new Error('the control names no block');
    expect(story.prompts.map((prompt) => ({ name: prompt.name, inline: prompt.inline }))).toEqual([
      { name, inline: false },
    ]);
    expect(textOf(story.prompts[0]?.nodes ?? [])).toBe('Alpha middle Omega.');
  });

  test("a decision that empties a control records it, inline and block, and both show Word's blank prompt", () => {
    const REV = 'w:id="5" w:author="Cara" w:date="2026-01-01T00:00:00Z"';
    const sdt = (props: string, content: string) =>
      `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
    const doc = loadOrderedPackage(
      packageOf(
        docx(
          p(t('Name: ') + sdt('<w:text/>', `<w:ins ${REV}>${t('Ana')}</w:ins>`)) +
            sdt('<w:alias w:val="Clause"/>', p(`<w:ins ${REV}>${t('Clause text')}</w:ins>`)) +
            p(t('After'))
        )
      )
    );
    const loaded = doc.stories.get(BODY)?.part;
    if (!loaded) throw new Error('no body');
    const ids = controlsIn(loaded).map((control) => control.id);
    withOrderedUndoRecord(doc, () => setOrderedStatus(doc, 'S1', 'rejected'));
    const { part, composed } = composeOrderedPart(doc, BODY);
    expect([...composed.emptied].sort()).toEqual([...ids].sort());
    // Neither names a prompt: each shows Word's blank prompt, five spaces in no style, as the
    // engine's decision leaves it (Word 16.114 samples `~$i6w-*-out.docx`)
    for (const id of ids) {
      const control = findNode(part, id);
      if (!control || control.kind === 'textValue') throw new Error('no control');
      expect(named(control, 'showingPlcHdr')).toBeDefined();
      expect(named(control, 'sdtEndPr')).toBeDefined();
      expect(textOf([control])).toBe('     ');
    }
    expect(canonicalOoxmlFingerprint(part)).not.toBe(canonicalOoxmlFingerprint(loaded));
  });
});
