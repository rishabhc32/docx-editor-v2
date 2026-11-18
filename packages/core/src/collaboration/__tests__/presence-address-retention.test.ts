// Presence address hints must not keep old trees alive.
//
// A hint remembers where a stable paraId resolved, and in which tree, so an unchanged tree
// answers without a lookup. It held that tree's root: each remote caret that resolved once kept
// the tree of its last resolution, up to 512 hints, and with each tree every cache keyed by its
// nodes. A long session with several people kept hundreds of whole document revisions.

import { expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createParagraphAddressResolver, isStoryPart } from '../paragraph-addresses.ts';
import { normalizeParagraphIdentity } from '../../store/package/para-id.ts';
import { readOoxmlPackage } from '../../store/package/ooxml-package.ts';
import type { OoxmlNode, OoxmlPart } from '../../store/package/ooxml-tree.ts';
import { TreePackageStore } from '../../store/store/tree-package-store.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

const PARAGRAPHS = 16;

function stableId(index: number): string {
  return (index + 1).toString(16).toUpperCase().padStart(8, '0');
}

function documentBytes(): Uint8Array {
  const prose = Array.from(
    { length: PARAGRAPHS },
    (_unused, index) =>
      `<w:p w14:paraId="${stableId(index)}" w14:textId="${stableId(index)}">` +
      `<w:r><w:t>Paragraph ${index}</w:t></w:r></w:p>`
  ).join('');
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rId1" Type="${OD}" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${prose}<w:sectPr/></w:body></w:document>`
    ),
  });
}

function openStore(): TreePackageStore {
  const loaded = readOoxmlPackage(documentBytes());
  if (!loaded.ok) throw new Error(loaded.reason);
  const main = loaded.package.parts.get(loaded.package.mainDocumentPart);
  if (!main) throw new Error('missing main part');
  return new TreePackageStore(loaded.package, normalizeParagraphIdentity(main));
}

function paragraphIds(part: OoxmlPart): string[] {
  const ids: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') ids.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(part.root);
  return ids;
}

/** One keystroke as a collaboration commit makes it: no legacy history keeps the old tree. */
function type(store: TreePackageStore, paragraphId: string): void {
  const result = store.transact(
    { kind: 'body' },
    (context) => {
      context.apply({ op: 'insertText', paragraphId, offset: 0, text: 'x' });
    },
    { recordsHistory: false }
  );
  if (!result.ok) throw new Error(result.detail ?? result.reason);
}

test('address hints keep no tree that an edit replaced', async () => {
  const store = openStore();
  const storyParts = () =>
    [...store.currentPackage().parts.values()].filter((part) => isStoryPart(part));
  const resolve = createParagraphAddressResolver(storyParts);
  const ids = paragraphIds(store.bodyStore().part);
  const roots: WeakRef<object>[] = [];
  // In its own frame, so no stack slot keeps a tree for the collector
  const step = (index: number): void => {
    type(store, ids[(index + 1) % PARAGRAPHS]!);
    roots.push(new WeakRef(store.bodyStore().part.root));
    // Each caret names another paragraph, so each hint is checked in another tree
    expect(resolve(stableId(index))?.nodeId).toBe(ids[index]!);
  };
  for (let index = 0; index < PARAGRAPHS; index += 1) step(index);
  await new Promise((resolve) => setTimeout(resolve, 0));
  Bun.gc(true);
  // The current tree, and one more that a stale stack slot may hold for the collector
  expect(roots.filter((ref) => ref.deref() !== undefined).length).toBeLessThanOrEqual(2);
  // Every hint still answers, in the tree the replica holds now
  for (let index = 0; index < PARAGRAPHS; index += 1) {
    const address = resolve(stableId(index));
    expect(address?.nodeId).toBe(ids[index]!);
    expect(address?.part.root).toBe(store.bodyStore().part.root);
  }
});

test('a hint from an earlier tree is checked again before it answers', () => {
  const store = openStore();
  const resolve = createParagraphAddressResolver(() => [store.bodyStore().part]);
  const ids = paragraphIds(store.bodyStore().part);
  expect(resolve(stableId(3))?.nodeId).toBe(ids[3]!);
  expect(resolve(stableId(3))?.part).toBe(store.bodyStore().part);
  // The paragraph goes: a hint from the earlier tree must not name it
  const removed = store.transact(
    { kind: 'body' },
    (context) => {
      context.apply({ op: 'deleteBlock', blockId: ids[3]! });
    },
    { recordsHistory: false }
  );
  if (!removed.ok) throw new Error(removed.detail ?? removed.reason);
  expect(resolve(stableId(3))).toBeNull();
  expect(resolve(stableId(4))?.nodeId).toBe(ids[4]!);
});
