// A received package that adds a second list to a numbering part the editor already has.
//
// A collaboration session hands the editor whole packages through the port. The package is
// the authority: after `applyRemotePackage` succeeds, the editor's body and numbering parts
// say what the package says — the new `w:abstractNum`/`w:num` AND the paragraph's `w:numPr`
// that names them — even when this editor made a list of its own that has not replicated
// yet. (The ordered room once composed packages whose list definitions shared node ids; the
// store refused those whole, so the editor kept its old paragraphs. These tests pin the
// store's half of the contract with packages built the way an editor builds them.)

import { describe, expect, test } from 'bun:test';
import { strToU8, zipSync } from 'fflate';
import { createCollaborationDocumentPort } from '../../collaboration/document-port.ts';
import { runWithTransactionActor } from '../package/actor-scoped-ids.ts';
import { ensureListDefinition, type ListKind } from '../package/numbering-part.ts';
import { readOoxmlPackage, type OoxmlPackage } from '../package/ooxml-package.ts';
import {
  serializeOoxmlPart,
  validateOoxmlPart,
  type OoxmlNode,
  type OoxmlPart,
} from '../package/ooxml-tree.ts';
import { ORIGIN_IDS } from '../registry/frozen-ids.ts';
import { TreePackageStore } from '../store/tree-package-store.ts';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WML = 'application/vnd.openxmlformats-officedocument.wordprocessingml';
const NUMBERING = '/word/numbering.xml';
const REMOTE = {
  origin: ORIGIN_IDS.mutationRemote,
  actorId: 'room',
  operationId: 'room-1',
};

/** Three paragraphs; the first is in list 1, which `numbering.xml` defines. */
function numberedDocument(): Uint8Array {
  const plain = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  return zipSync({
    '[Content_Types].xml': strToU8(
      `<Types xmlns="${CT}">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        `<Override PartName="/word/document.xml" ContentType="${WML}.document.main+xml"/>` +
        `<Override PartName="${NUMBERING}" ContentType="${WML}.numbering+xml"/>` +
        '</Types>'
    ),
    '_rels/.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`
    ),
    'word/_rels/document.xml.rels': strToU8(
      `<Relationships xmlns="${REL}"><Relationship Id="rId5" Type="${R}/numbering" Target="numbering.xml"/></Relationships>`
    ),
    'word/document.xml': strToU8(
      `<w:document xmlns:w="${W}"><w:body>` +
        '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>' +
        '<w:r><w:t>One</w:t></w:r></w:p>' +
        `${plain('Two')}${plain('Three')}<w:sectPr/>` +
        '</w:body></w:document>'
    ),
    'word/numbering.xml': strToU8(
      `<w:numbering xmlns:w="${W}">` +
        '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>' +
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>' +
        '<w:lvlText w:val="%1."/></w:lvl></w:abstractNum>' +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
        '</w:numbering>'
    ),
  });
}

function openStore(): TreePackageStore {
  const read = readOoxmlPackage(numberedDocument());
  if (!read.ok) throw new Error(read.reason);
  const main = read.package.parts.get(read.package.mainDocumentPart);
  if (!main) throw new Error('no main document part');
  return new TreePackageStore(read.package, main);
}

const element = (node: OoxmlNode, localName: string): OoxmlNode | undefined =>
  node.kind === 'textValue'
    ? undefined
    : node.children.find(
        (child) =>
          child.kind !== 'textValue' && child.namespaceUri === W && child.localName === localName
      );

const valOf = (node: OoxmlNode | undefined): string | null =>
  node && node.kind !== 'textValue'
    ? (node.attributes.find((attribute) => attribute.localName === 'val')?.value ?? null)
    : null;

function paragraphsOf(part: OoxmlPart): OoxmlNode[] {
  const found: OoxmlNode[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'paragraph') found.push(node);
    else if (node.kind !== 'textValue') node.children.forEach(visit);
  };
  visit(part.root);
  return found;
}

/** The list number each body paragraph's `w:numPr` names, or null. */
function bodyListNumbers(pkg: OoxmlPackage): (string | null)[] {
  const body = pkg.parts.get(pkg.mainDocumentPart);
  if (!body) throw new Error('no main document part');
  return paragraphsOf(body).map((paragraph) => {
    const pPr = element(paragraph, 'pPr');
    const numPr = pPr ? element(pPr, 'numPr') : undefined;
    return numPr ? valOf(element(numPr, 'numId')) : null;
  });
}

/** The `w:numId` of every `w:num` in the numbering part, in order. */
function definedListNumbers(pkg: OoxmlPackage): string[] {
  const numbering = pkg.parts.get(NUMBERING);
  if (!numbering || numbering.root.kind === 'textValue') return [];
  return numbering.root.children.flatMap((child) =>
    child.kind !== 'textValue' && child.localName === 'num'
      ? [child.attributes.find((attribute) => attribute.localName === 'numId')?.value ?? '']
      : []
  );
}

/**
 * Make a new list on one paragraph the way the editor does (`tree-session.ts`): the
 * definition is a package-shell write, then the paragraph's `w:numPr` is a body op. The actor
 * stripes the new `w:numId`, as it does for an attached replica.
 */
function makeList(store: TreePackageStore, index: number, kind: ListKind, actor: string): string {
  const ensured = runWithTransactionActor(actor, () =>
    ensureListDefinition(store.currentPackage(), kind)
  );
  if (!ensured) throw new Error('no list definition');
  store.replacePackageShell(ensured.pkg);
  const paragraph = paragraphsOf(store.bodyStore().part)[index];
  if (!paragraph) throw new Error(`no paragraph ${index}`);
  const result = store.transact({ kind: 'body' }, (context) => {
    context.apply({
      op: 'setListNumbering',
      paragraphId: paragraph.id,
      numId: ensured.numId,
      level: 0,
    });
  });
  if (!result.ok) throw new Error(result.detail ?? result.reason);
  return ensured.numId;
}

function partXml(pkg: OoxmlPackage, name: string): string {
  const part = pkg.parts.get(name);
  if (!part) throw new Error(`no ${name}`);
  return serializeOoxmlPart(part);
}

/** The editor's body and numbering parts say exactly what the received package says. */
function expectPartsOf(store: TreePackageStore, received: OoxmlPackage): void {
  const installed = store.currentPackage();
  expect(partXml(installed, installed.mainDocumentPart)).toBe(
    partXml(received, received.mainDocumentPart)
  );
  expect(serializeOoxmlPart(store.bodyStore().part)).toBe(
    partXml(received, received.mainDocumentPart)
  );
  expect(partXml(installed, NUMBERING)).toBe(partXml(received, NUMBERING));
  expect(validateOoxmlPart(installed.parts.get(NUMBERING)!).ok).toBe(true);
}

describe('a received package adds a second list to an existing numbering part', () => {
  test('the editor takes the new definition and the paragraph that names it', () => {
    const editor = openStore();
    const port = createCollaborationDocumentPort(editor, { documentId: 'second-list' });
    const peer = openStore();
    const second = makeList(peer, 1, 'bullet', 'peer');
    const received = peer.currentPackage();
    expect(bodyListNumbers(received)).toEqual(['1', second, null]);
    expect(definedListNumbers(received)).toEqual(['1', second]);
    expect(definedListNumbers(editor.currentPackage())).toEqual(['1']);

    expect(port.applyRemotePackage(received, REMOTE)).toEqual({ ok: true, changed: true });

    expectPartsOf(editor, received);
    expect(bodyListNumbers(editor.currentPackage())).toEqual(['1', second, null]);
    // Installed as received, not merged with this editor's own copy of the part.
    expect(editor.currentPackage().parts.get(NUMBERING)).toBe(received.parts.get(NUMBERING));
  });

  test('the received package wins over a list this editor made before its write replicated', () => {
    const editor = openStore();
    const port = createCollaborationDocumentPort(editor, { documentId: 'second-list' });
    const local = makeList(editor, 2, 'bullet', 'editor');
    const peer = openStore();
    const second = makeList(peer, 1, 'bullet', 'peer');
    expect(local).not.toBe(second);
    expect(definedListNumbers(editor.currentPackage())).toEqual(['1', local]);
    const received = peer.currentPackage();

    expect(port.applyRemotePackage(received, REMOTE)).toEqual({ ok: true, changed: true });

    expectPartsOf(editor, received);
    expect(bodyListNumbers(editor.currentPackage())).toEqual(['1', second, null]);
    expect(definedListNumbers(editor.currentPackage())).toEqual(['1', second]);

    // The next local write keeps what the package said: the editor's own list does not
    // come back from its package shell.
    const paragraph = paragraphsOf(editor.bodyStore().part)[1]!;
    const typed = editor.transact({ kind: 'body' }, (context) => {
      context.apply({ op: 'insertText', paragraphId: paragraph.id, offset: 0, text: 'x' });
    });
    expect(typed.ok).toBe(true);
    expect(partXml(editor.currentPackage(), NUMBERING)).toBe(partXml(received, NUMBERING));
    expect(bodyListNumbers(editor.currentPackage())).toEqual(['1', second, null]);
  });
});
