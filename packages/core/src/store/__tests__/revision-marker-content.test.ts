// Which property containers resolving a paragraph mark may remove.
import { expect, test } from 'bun:test';
import { WML_NAMESPACE_URI, type OoxmlNode } from '../package/ooxml-tree.ts';
import { heldOnlyResolvedMarks } from '../store/revision-marker-content.ts';

const element = (id: string, kind: string, localName: string, children: OoxmlNode[] = []) =>
  ({
    id,
    kind,
    namespaceUri: WML_NAMESPACE_URI,
    localName,
    prefix: 'w',
    attributes: [],
    children,
  }) as unknown as OoxmlNode;

/** `w:pPr` holding `others` and `w:rPr` holding one mark, and the mark's id. */
function properties(others: OoxmlNode[] = []): { node: OoxmlNode; mark: string } {
  const mark = element('/word/document.xml#0.0.2', 'generic', 'del');
  const rPr = element('/word/document.xml#0.0.1', 'runProperties', 'rPr', [mark]);
  return {
    node: element('/word/document.xml#0.0.0', 'paragraphProperties', 'pPr', [...others, rPr]),
    mark: mark.id,
  };
}

test('containers that held only the resolved marks are removed, whoever made them', () => {
  const { node, mark } = properties();
  expect(heldOnlyResolvedMarks(node, new Set([mark]))).toBe(true);
});

test('containers that hold anything else stay', () => {
  const centered = properties([element('/word/document.xml#0.0.3', 'generic', 'jc')]);
  expect(heldOnlyResolvedMarks(centered.node, new Set([centered.mark]))).toBe(false);
  const { node } = properties();
  expect(heldOnlyResolvedMarks(node, new Set())).toBe(false);
});

test('a container that arrives empty holds nothing this resolution drops', () => {
  const empty = element('/word/document.xml#0.0.4', 'paragraphProperties', 'pPr');
  expect(heldOnlyResolvedMarks(empty, new Set())).toBe(false);
});
