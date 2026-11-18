// Editing API writes that change only package parts, in an ordered room: the document
// properties (`document.properties.title = ...`) and a bookmark deleted from a header the script
// does not write in. Each reaches the other copy and the room's Word file as the editor alone
// writes it. Before, the room refused both as opaque writes (`ordered-opaque-write`).

import { afterEach, expect, test } from 'bun:test';
import { createBrowserAutomationHost } from '../../../../editor/automation-host.ts';
import { documentProperty } from '../../../../store/package/document-property-writes.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { docx, p, R, t, WML_CT } from '../word-fixtures.ts';
import { closeAll, join, localRoom, ORGANISER, settle, surface } from './editor-harness.ts';
import { alone, converged, roomFile, storyAsRead, type Act } from './links-harness.ts';

afterEach(closeAll);

const CORE =
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Old title</dc:title></cp:coreProperties>';
const HEADER =
  `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p>` +
  '<w:bookmarkStart w:id="7" w:name="Mark"/><w:r><w:t>Header text</w:t></w:r>' +
  '<w:bookmarkEnd w:id="7"/></w:p></w:hdr>';

/** A body, a default header holding the bookmark `Mark`, and core document properties. */
const FILE = docx(
  p(t('First page text.')) +
    `<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>`,
  {
    parts: {
      'word/header1.xml': HEADER,
      'docProps/core.xml': CORE,
      '_rels/.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/>` +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '</Relationships>',
    },
    relationships: `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes:
      `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>` +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
  }
);

/** One operation through the editing API, with the document's handle. */
function execute(operation: (document: unknown) => object): Act {
  return (editor) => {
    const host = createBrowserAutomationHost(editor);
    const [read] = host.execute({ operations: [{ op: 'getDocument' }] }).results;
    const document = (read as unknown as { value: { handle: unknown } }).value.handle;
    return host.execute({ operations: [operation(document)] } as never);
  };
}

const setTitle = execute(() => ({ op: 'setDocumentProperties', values: { title: 'New title' } }));
const deleteMark = execute((document) => ({
  op: 'deleteBookmark',
  target: { document, name: 'Mark' },
}));

const headerAsRead = (pkg: OoxmlPackage) => storyAsRead(pkg, '/word/header1.xml');

test("the editing API's document properties reach the other copy and the room's file", async () => {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  expect(setTitle(olive.editor)).toMatchObject({ ok: true });
  await settle(20);
  converged(room, [olive, oscar]);
  for (const person of [olive, oscar])
    expect(documentProperty(surface(person).session.currentPackage(), 'title')).toBe('New title');
  const editorAlone = await alone(FILE, 'Olive', 'edit', setTitle);
  const saved = roomFile(room.room.document, FILE);
  expect(documentProperty(saved, 'title')).toBe(documentProperty(editorAlone, 'title'));
  expect(serializeOoxmlPart(saved.parts.get('/docProps/core.xml')!)).toBe(
    serializeOoxmlPart(editorAlone.parts.get('/docProps/core.xml')!)
  );
});

test("the editing API's deletion of a header's bookmark reaches the other copy and the file", async () => {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  expect(deleteMark(olive.editor)).toMatchObject({ ok: true });
  await settle(20);
  converged(room, [olive, oscar]);
  const editorAlone = await alone(FILE, 'Olive', 'edit', deleteMark);
  const saved = roomFile(room.room.document, FILE);
  expect(headerAsRead(saved)).not.toContain('bookmarkStart');
  expect(headerAsRead(saved)).toBe(headerAsRead(editorAlone));
});
