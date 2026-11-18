// What a received package tells layout when paragraphs keep their ids but their inner nodes do
// not, and when a split or a join changes a block list. A collaboration view (the ordered
// rooms' list) composes a paragraph again under new inner ids each time, and its Enter makes a
// paragraph under its own id; both used to read as a change to the whole document.

import { describe, expect, test } from 'bun:test';
import { zipSync, strToU8 } from 'fflate';
import { canonicalTreeDifference } from '../package/ooxml-difference.ts';
import { ooxmlTreesEqual } from '../package/ooxml-serialize.ts';
import { readOoxmlPackage, withPart, type OoxmlPackage } from '../package/ooxml-package.ts';
import type { OoxmlElement, OoxmlNode, OoxmlPart } from '../package/ooxml-tree.ts';
import { ORIGIN_IDS } from '../registry/frozen-ids.ts';
import type { TreeDocOp } from '../store/tree-op-types.ts';
import { remotePackageDelta } from '../store/tree-package-remote.ts';
import { TreePackageStore } from '../store/tree-package-store.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

function documentBytes(body: string): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`
    ),
  });
}

const BODY =
  '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>beta gamma</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>delta</w:t></w:r></w:p>' +
  '<w:sectPr/>';

function openStore(body = BODY): TreePackageStore {
  const result = readOoxmlPackage(documentBytes(body));
  if (!result.ok) throw new Error(result.reason);
  const main = result.package.parts.get(result.package.mainDocumentPart);
  if (!main) throw new Error('no main part');
  return new TreePackageStore(result.package, main);
}

function mainPart(pkg: OoxmlPackage): OoxmlPart {
  const part = pkg.parts.get(pkg.mainDocumentPart);
  if (!part) throw new Error('no main part');
  return part;
}

function bodyNode(pkg: OoxmlPackage): OoxmlElement {
  const body = mainPart(pkg).root.children.find(
    (child) => child.kind !== 'textValue' && child.localName === 'body'
  );
  if (!body || body.kind === 'textValue') throw new Error('no body');
  return body;
}

function paragraphs(pkg: OoxmlPackage): OoxmlElement[] {
  return bodyNode(pkg).children.filter(
    (child): child is OoxmlElement => child.kind === 'paragraph'
  );
}

function apply(store: TreePackageStore, op: TreeDocOp): void {
  const result = store.transact({ kind: 'body' }, (ctx) => {
    ctx.apply(op);
  });
  if (!result.ok) throw new Error(result.reason);
}

/** A copy of `node` with every id renamed, `tag` added to each. */
function renamed(node: OoxmlNode, tag: string): OoxmlNode {
  if (node.kind === 'textValue') return { ...node, id: `${node.id}${tag}` };
  return { ...node, id: `${node.id}${tag}`, children: node.children.map((c) => renamed(c, tag)) };
}

/** The package with the body's children replaced. */
function withBody(pkg: OoxmlPackage, children: readonly OoxmlNode[]): OoxmlPackage {
  const main = mainPart(pkg);
  const body = bodyNode(pkg);
  const root = main.root;
  const nextBody: OoxmlElement = { ...body, children };
  return withPart(pkg, {
    ...main,
    root: { ...root, children: root.children.map((c) => (c === body ? nextBody : c)) },
  });
}

/** A paragraph whose inner nodes take new ids, the paragraph keeping its own. */
function innerRenamed(paragraph: OoxmlElement, tag: string, text?: string): OoxmlElement {
  const children = paragraph.children.map((child) => {
    const copy = renamed(child, tag);
    if (text === undefined) return copy;
    const swap = (node: OoxmlNode): OoxmlNode =>
      node.kind === 'textValue'
        ? { ...node, value: text }
        : { ...node, children: node.children.map(swap) };
    return swap(copy);
  });
  return { ...paragraph, children };
}

describe('renamed inner nodes inside paragraphs that keep their ids', () => {
  test('a run under new ids, its text changed, is text-local on that paragraph', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [first, second, third] = paragraphs(current);
    const next = withBody(
      current,
      bodyNode(current).children.map((child) =>
        child === second ? innerRenamed(second!, '~1', 'beta gamma!') : child
      )
    );
    const delta = remotePackageDelta(next, current);
    expect(delta).toMatchObject({
      equal: false,
      impact: 'text-local',
      dirty: [second!.id],
      created: [],
      deleted: [],
    });
    expect(delta.dirty).not.toContain(first!.id);
    expect(delta.dirty).not.toContain(third!.id);
  });

  test('equal content under new inner ids installs, text-local on that paragraph', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [, second] = paragraphs(current);
    const next = withBody(
      current,
      bodyNode(current).children.map((child) =>
        child === second ? innerRenamed(second!, '~2') : child
      )
    );
    expect(remotePackageDelta(next, current)).toMatchObject({
      equal: false,
      impact: 'text-local',
      dirty: [second!.id],
    });
  });

  test('a renamed control inside a paragraph is flow-structural', () => {
    const store = openStore(
      '<w:p><w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent><w:r><w:t>in</w:t></w:r>' +
        '</w:sdtContent></w:sdt></w:p><w:sectPr/>'
    );
    const current = store.currentPackage();
    const [only] = paragraphs(current);
    const next = withBody(
      current,
      bodyNode(current).children.map((child) =>
        child === only ? innerRenamed(only!, '~3') : child
      )
    );
    expect(remotePackageDelta(next, current)).toMatchObject({
      impact: 'flow-structural',
      dirty: [only!.id],
    });
  });

  test('a renamed paragraph inside a paragraph (a text box) stays wholesale', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [first] = paragraphs(current);
    // A paragraph node nested in a paragraph: its id is a block's, not an inline node's
    const nested: OoxmlElement = {
      ...first!,
      children: [...first!.children, renamed(first!, '~n')],
    };
    const base = withBody(
      current,
      bodyNode(current).children.map((child) => (child === first ? nested : child))
    );
    const next = withBody(
      base,
      bodyNode(base).children.map((child) =>
        child === nested
          ? { ...nested, children: [...first!.children, renamed(first!, '~m')] }
          : child
      )
    );
    expect(remotePackageDelta(next, base).impact).toBe('global');
  });
});

describe('a split or a join changes the block list', () => {
  test('a split names the head dirty and the tail created', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [, second] = paragraphs(current);
    apply(store, { op: 'splitParagraph', paragraphId: second!.id, offset: 4 });
    const next = store.currentPackage();
    const tail = paragraphs(next)[2]!;
    const delta = remotePackageDelta(next, current);
    expect(delta).toMatchObject({
      equal: false,
      impact: 'flow-structural',
      dirty: [second!.id],
      created: [tail.id],
      deleted: [],
    });
  });

  test('a join names the kept paragraph dirty and the other deleted', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [, second, third] = paragraphs(current);
    apply(store, { op: 'joinParagraphs', firstId: second!.id, secondId: third!.id });
    const delta = remotePackageDelta(store.currentPackage(), current);
    expect(delta.impact).toBe('flow-structural');
    expect(delta.created).toEqual([]);
    expect(delta.deleted).toHaveLength(1);
    expect([second!.id, third!.id]).toContain(delta.deleted[0]!);
    expect(delta.dirty).toHaveLength(1);
  });

  test('a paragraph replaced under another id is one created and one deleted', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [, second] = paragraphs(current);
    const other = renamed(second!, '~new');
    const next = withBody(
      current,
      bodyNode(current).children.map((child) => (child === second ? other : child))
    );
    expect(remotePackageDelta(next, current)).toMatchObject({
      impact: 'flow-structural',
      dirty: [],
      created: [other.id],
      deleted: [second!.id],
    });
  });

  test('the engine head and tail replaced by a view whose head is new and tail kept', () => {
    // The typist's editor after Enter: the engine split under its ids (head kept, tail new);
    // the view's split keeps the id on the tail and gives the head a new one
    const store = openStore();
    const before = store.currentPackage();
    const [, second] = paragraphs(before);
    apply(store, { op: 'splitParagraph', paragraphId: second!.id, offset: 4 });
    const current = store.currentPackage();
    const [, head, engineTail] = paragraphs(current);
    const viewHead = renamed(head!, '~view');
    const viewTail: OoxmlElement = { ...engineTail!, id: second!.id };
    const next = withBody(
      current,
      bodyNode(current).children.flatMap((child) =>
        child === head ? [viewHead] : child === engineTail ? [viewTail] : [child]
      )
    );
    const delta = remotePackageDelta(next, current);
    expect(delta.impact).toBe('flow-structural');
    expect(delta.created).toEqual([viewHead.id]);
    expect(delta.deleted).toEqual([engineTail!.id]);
    expect(delta.dirty).toEqual([second!.id]);
  });

  test('two kept paragraphs that swap places are wholesale', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [first, second, third] = paragraphs(current);
    const next = withBody(
      current,
      bodyNode(current).children.map((child) =>
        child === first ? third! : child === third ? first! : child
      )
    );
    expect(remotePackageDelta(next, current).impact).toBe('global');
    expect(second).toBeDefined();
  });

  test('a paragraph moved into a table is wholesale', () => {
    const store = openStore(
      '<w:p><w:r><w:t>alpha</w:t></w:r></w:p>' +
        '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
        '<w:sectPr/>'
    );
    const current = store.currentPackage();
    const body = bodyNode(current);
    const [first, table, sectPr] = body.children as OoxmlElement[];
    const row = table!.children.find((c) => c.kind !== 'textValue' && c.localName === 'tr');
    if (!row || row.kind === 'textValue') throw new Error('no row');
    const cell = row.children[0] as OoxmlElement;
    const movedCell: OoxmlElement = { ...cell, children: [...cell.children, first!] };
    const movedRow: OoxmlElement = { ...row, children: [movedCell] };
    const movedTable: OoxmlElement = {
      ...table!,
      children: table!.children.map((c) => (c === row ? movedRow : c)),
    };
    const next = withBody(current, [movedTable, sectPr!]);
    expect(remotePackageDelta(next, current).impact).toBe('global');
  });

  test('the verdict agrees with the oracle when the paragraphs between say the same in place', () => {
    const store = openStore();
    const current = store.currentPackage();
    const [first, second] = paragraphs(current);
    // In place the text reads alpha, beta gamma as before; by id, `first` now says beta gamma
    const newFirst = renamed(first!, '~x');
    const shifted: OoxmlElement = { ...second!, id: first!.id };
    const next = withBody(
      current,
      bodyNode(current).children.map((child) =>
        child === first ? newFirst : child === second ? shifted : child
      )
    );
    const difference = canonicalTreeDifference(mainPart(current), mainPart(next));
    expect(difference.undecided).toBe(false);
    expect(difference.equal).toBe(ooxmlTreesEqual(mainPart(current), mainPart(next)));
    expect(difference.equal).toBe(true);
    expect(difference.createdParagraphIds).toEqual([newFirst.id]);
    expect(difference.deletedParagraphIds).toEqual([second!.id]);
    expect(difference.paragraphIds).toEqual([first!.id]);
    // Equal content under other ids still installs, as a structural change
    expect(remotePackageDelta(next, current)).toMatchObject({
      equal: false,
      impact: 'flow-structural',
    });
  });

  test('a received split publishes the paragraphs it touched, not the document', () => {
    const sender = openStore();
    // The receiver holds the sender's package: the two share every unchanged object, as two
    // replicas do once the first package is installed
    const shared = sender.currentPackage();
    const receiver = new TreePackageStore(shared, mainPart(shared));
    const [, second] = paragraphs(sender.currentPackage());
    apply(sender, { op: 'splitParagraph', paragraphId: second!.id, offset: 4 });
    const result = receiver.publishRemotePackage(sender.currentPackage(), {
      origin: ORIGIN_IDS.mutationRemote,
      actorId: 'bob',
      operationId: 'bob-1',
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.change?.impact).toBe('flow-structural');
    expect(result.change?.dirty).toEqual([second!.id]);
    expect(result.change?.created).toEqual([paragraphs(sender.currentPackage())[2]!.id]);
    expect(result.change?.deleted).toEqual([]);
    expect(paragraphs(receiver.currentPackage())).toHaveLength(4);
  });
});
