// A converted note's paragraphs take paragraph ids that are new in the notes part they go to.
//
// The source note's `w14:paraId` values are unique in the part it came from, not in the target
// part. Copying them gave two paragraphs of the target part one id; Word writes a unique id on
// every paragraph, and a collaboration session finds a paragraph by it.

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import {
  readOoxmlPackage,
  resolveNotesPart,
  type OoxmlNode,
  type OoxmlPackage,
} from '../package/index.ts';
import { paraIdOf } from '../package/para-id.ts';
import { TreePackageStore } from '../store/tree-package-store.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

const p = (paraId: string, text: string, mark = '') =>
  `<w:p w14:paraId="${paraId}" w14:textId="77777777">${mark}<w:r><w:t>${text}</w:t></w:r></w:p>`;

const separators = (kind: 'footnote' | 'endnote') =>
  `<w:${kind} w:type="separator" w:id="-1">${p('00000001', '')}</w:${kind}>` +
  `<w:${kind} w:type="continuationSeparator" w:id="0">${p('00000002', '')}</w:${kind}>`;

function docx(): Uint8Array {
  const notes = (kind: 'footnote' | 'endnote', body: string) =>
    strToU8(
      `<w:${kind}s xmlns:w="${W}" xmlns:w14="${W14}">${separators(kind)}${body}</w:${kind}s>`
    );
  const ref = (kind: 'footnote' | 'endnote', id: number) =>
    `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/>` +
        `<Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/>` +
        `<Override PartName="/word/endnotes.xml" ContentType="${MAIN}.endnotes+xml"/></Types>`
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}">` +
        `<Relationship Id="rIdFn" Type="${R}/footnotes" Target="footnotes.xml"/>` +
        `<Relationship Id="rIdEn" Type="${R}/endnotes" Target="endnotes.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
        `<w:p w14:paraId="0000000A" w14:textId="77777777"><w:r><w:t>Body</w:t></w:r>` +
        `${ref('footnote', 1)}${ref('footnote', 2)}${ref('endnote', 1)}</w:p>` +
        '<w:sectPr/></w:body></w:document>'
    ),
    // Each footnote's paragraphs carry ids the endnotes part already uses.
    'word/footnotes.xml': notes(
      'footnote',
      `<w:footnote w:id="1">${p('1A2B3C4D', 'one')}${p('2B3C4D5E', 'one, second')}</w:footnote>` +
        `<w:footnote w:id="2">${p('3C4D5E6F', 'two')}</w:footnote>`
    ),
    'word/endnotes.xml': notes(
      'endnote',
      `<w:endnote w:id="1">${p('1A2B3C4D', 'end')}${p('2B3C4D5E', 'end, second')}` +
        `${p('3C4D5E6F', 'end, third')}</w:endnote>`
    ),
  });
}

function openStore(): TreePackageStore {
  const read = readOoxmlPackage(docx());
  if (!read.ok) throw new Error(read.reason);
  const main = read.package.parts.get(read.package.mainDocumentPart)!;
  return new TreePackageStore(read.package, main);
}

function paragraphIds(pkg: OoxmlPackage, kind: 'footnote' | 'endnote'): (string | null)[] {
  const ids: (string | null)[] = [];
  const walk = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'paragraph') ids.push(paraIdOf(node));
    for (const child of node.children) walk(child);
  };
  walk(resolveNotesPart(pkg, kind)!.root);
  return ids;
}

function expectUnique(ids: readonly (string | null)[]): void {
  expect(ids.every((id) => id !== null)).toBe(true);
  expect(new Set(ids).size).toBe(ids.length);
}

describe('convertNote gives the copied paragraphs new ids', () => {
  test('one footnote into an endnotes part that has the same ids', () => {
    const store = openStore();
    const result = store.applyLifecycleOp({ op: 'convertNote', fromKind: 'footnote', noteId: 1 });
    expect(result.ok).toBe(true);
    const ids = paragraphIds(store.currentPackage(), 'endnote');
    // Two separators, the endnote's three paragraphs, and the converted note's two.
    expect(ids).toHaveLength(7);
    expectUnique(ids);
  });

  test('every footnote into the endnotes part', () => {
    const store = openStore();
    const result = store.applyLifecycleOp({ op: 'convertAllNotes', fromKind: 'footnote' });
    expect(result.ok).toBe(true);
    const ids = paragraphIds(store.currentPackage(), 'endnote');
    expect(ids).toHaveLength(8);
    expectUnique(ids);
  });

  test('the ids stay new after undo and redo', () => {
    const store = openStore();
    expect(store.applyLifecycleOp({ op: 'convertNote', fromKind: 'footnote', noteId: 1 }).ok).toBe(
      true
    );
    const converted = paragraphIds(store.currentPackage(), 'endnote');
    store.undo();
    expect(paragraphIds(store.currentPackage(), 'endnote')).toHaveLength(5);
    store.redo();
    expect(paragraphIds(store.currentPackage(), 'endnote')).toEqual(converted);
  });
});
