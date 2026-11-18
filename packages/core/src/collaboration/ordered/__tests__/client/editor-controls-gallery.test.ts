// Building block gallery controls with real editors and one ordered room (task 9.6): the
// organiser picks a glossary block for an inline and a block gallery control, every copy shows
// what the editor alone writes, and the organiser's undo and redo are exact. A party's pick is
// a direct edit, refused by the host's party rule (`text`).

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { buildingBlocksOf } from '../../../../store/package/building-blocks.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import type { TreeDocOp } from '../../../../store/store/tree-op-types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t, W, WML_CT } from '../word-fixtures.ts';
import {
  closeAll,
  editorTexts,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const GALLERY = '<w:docPartList><w:docPartGallery w:val="Quick Parts"/></w:docPartList>';
const docPart = (name: string, body: string) =>
  `<w:docPart><w:docPartPr><w:name w:val="${name}"/><w:category><w:name w:val="General"/>` +
  `<w:gallery w:val="quickParts"/></w:category></w:docPartPr><w:docPartBody>${body}</w:docPartBody></w:docPart>`;

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
/** A run whose choice names `wps`, which only the glossary's root declares, as Word writes it. */
const SHAPE_RUN =
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:t>a</w:t></mc:Choice>' +
  '<mc:Fallback><w:t>b</w:t></mc:Fallback></mc:AlternateContent></w:r>';

/** An inline and a block gallery control. */
const BODY =
  p(t('Party: ') + sdt(GALLERY, t('Part')) + t('.')) +
  sdt(GALLERY, p(t('Clause'))) +
  p(t('Signed here.'));
/** A glossary of three blocks. */
const GLOSSARY = {
  parts: {
    'word/glossary/document.xml':
      `<w:glossaryDocument xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="${WPS}"><w:docParts>` +
      docPart('Acme', p(t('Acme Ltd.'))) +
      docPart('Terms', p(t('First term.')) + p(t('Second term.'))) +
      docPart('Seal', p(t('Sealed ') + SHAPE_RUN)) +
      '</w:docParts></w:glossaryDocument>',
  },
  contentTypes: `<Override PartName="/word/glossary/document.xml" ContentType="${WML_CT}.document.glossary+xml"/>`,
  relationships:
    '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
};
const FILE = docx(BODY, GLOSSARY);
/** The same gallery controls, temporary: the engine removes each once a block fills it. */
const TEMPORARY_FILE = docx(BODY.replaceAll(GALLERY, `<w:temporary/>${GALLERY}`), GLOSSARY);
/** The same, its document's root declaring the shape's prefixes too, as Word writes both. */
const SHAPES_FILE = docx(BODY, {
  ...GLOSSARY,
  parts: {
    ...GLOSSARY.parts,
    'word/document.xml': `<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wps="${WPS}"><w:body>${BODY}</w:body></w:document>`,
  },
});

const BLOCKS = new Map<string, readonly OoxmlNode[]>();
{
  const pkg = readOoxmlPackage(FILE);
  if (!pkg.ok) throw new Error('fixture');
  for (const block of buildingBlocksOf(pkg.package)) BLOCKS.set(block.name, block.blocks);
}

const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '');

function apply(person: Person, op: TreeDocOp): void {
  const s = surface(person) as unknown as {
    applyGatedTreeOps(ops: readonly TreeDocOp[]): unknown;
  };
  s.applyGatedTreeOps([op]);
}

function controlIds(person: Person): string[] {
  const ids: string[] = [];
  const visit = (node: OoxmlNode): void => {
    if (node.kind === 'textValue') return;
    if (node.kind === 'contentControl') ids.push(node.id);
    for (const child of node.children) visit(child);
  };
  visit(surface(person).session.part().root);
  return ids;
}

const pick = (person: Person, index: number, name: string): TreeDocOp => ({
  op: 'insertBuildingBlock',
  controlId: controlIds(person)[index] ?? '',
  name,
  blocks: BLOCKS.get(name) ?? [],
});

function alone(index: number, name: string, file = FILE): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olive',
    mode: 'edit',
    modules: [stubReviewModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    apply(person, pick(person, index, name));
    return bodyOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
  }
}

describe('building block gallery controls', () => {
  for (const [index, name, where] of [
    [0, 'Acme', 'inline'],
    [1, 'Terms', 'block'],
  ] as const)
    test(`the organiser picks a block for a ${where} gallery, then undo and redo`, async () => {
      const room = localRoom(FILE);
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
      const ben = await join(room, 'Ben', { wordFile: FILE });
      const before = bodyOf(editorXml(olive));
      apply(olive, pick(olive, index, name));
      await settle();
      converged(room, olive, ben);
      const after = bodyOf(editorXml(olive));
      expect(after).toBe(alone(index, name));
      expect(bodyOf(editorXml(ben))).toBe(after);
      surface(olive).undo();
      await settle();
      converged(room, olive, ben);
      expect(bodyOf(editorXml(olive))).toBe(before);
      surface(olive).redo();
      await settle();
      converged(room, olive, ben);
      expect(bodyOf(editorXml(olive))).toBe(after);
      const saved = new Uint8Array(await olive.editor.save());
      const reopened = openOrderedWordFile(saved, { file: 2 });
      expect(reopened.document.stories.size).toBeGreaterThan(0);
      expect(editorTexts(ben)).toEqual(editorTexts(olive));
    });

  for (const [index, name, where] of [
    [0, 'Acme', 'inline'],
    [1, 'Terms', 'block'],
  ] as const)
    test(`a temporary ${where} gallery goes once a block fills it, as in the editor alone`, async () => {
      const file = TEMPORARY_FILE;
      const room = localRoom(file);
      const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
      const ben = await join(room, 'Ben', { wordFile: file });
      const before = bodyOf(editorXml(olive));
      apply(olive, pick(olive, index, name));
      await settle();
      converged(room, olive, ben);
      const after = bodyOf(editorXml(olive));
      expect(after.match(/<w:sdt>/g)).toHaveLength(1);
      expect(after).toBe(alone(index, name, file));
      expect(bodyOf(editorXml(ben))).toBe(after);
      surface(olive).undo();
      await settle();
      converged(room, olive, ben);
      expect(bodyOf(editorXml(olive))).toBe(before);
      surface(olive).redo();
      await settle();
      converged(room, olive, ben);
      expect(bodyOf(editorXml(olive))).toBe(after);
    });

  test("a block whose choice names a prefix of the glossary's root is the room's block", async () => {
    const room = localRoom(SHAPES_FILE);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: SHAPES_FILE,
    });
    const ben = await join(room, 'Ben', { wordFile: SHAPES_FILE });
    apply(olive, pick(olive, 1, 'Seal'));
    await settle();
    converged(room, olive, ben);
    const after = bodyOf(editorXml(olive));
    expect(after).toContain('<mc:Choice Requires="wps">');
    expect(after).toBe(alone(1, 'Seal', SHAPES_FILE));
    expect(bodyOf(editorXml(ben))).toBe(after);
  });

  test("a party's pick is a direct edit: refused with the party rule's reason", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    apply(ana, pick(ana, 0, 'Acme'));
    await settle();
    expect(ana.refusals.map((r) => r.reason)).toEqual(['text']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(screenFollowsList(ana)).toBe(true);
  });
});
