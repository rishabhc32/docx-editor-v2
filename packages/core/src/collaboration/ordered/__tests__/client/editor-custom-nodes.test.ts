// A custom node written through the editing API (`insertCustomNode`) in an ordered room: the
// control and the payload node its binding names reach the other copy and the room's Word file,
// as the editor alone writes them. Before, the room refused the write (`package-scoped`).

import { afterEach, expect, test } from 'bun:test';
import { createBrowserAutomationHost } from '../../../../editor/automation-host.ts';
import { readCustomXmlNode } from '../../../../store/package/custom-xml-nodes.ts';
import { findCustomXmlDataPart } from '../../../../store/package/custom-xml-part.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { closeAll, FILE, join, localRoom, ORGANISER, settle, surface } from './editor-harness.ts';
import { alone, converged, roomFile, storyAsRead, type Act } from './links-harness.ts';

afterEach(closeAll);

const NS = 'urn:example:fields';

/** The value of the first result of one operation, typed by the caller. */
function valueOf<T>(host: ReturnType<typeof createBrowserAutomationHost>, operation: object): T {
  const [result] = host.execute({ operations: [operation] } as never).results;
  if (result?.status !== 'ok') throw new Error('the read was refused');
  return (result as unknown as { value: T }).value;
}

/** The editing API inserts a chip bound to node `n1` at the start of the first paragraph. */
const insertChip: Act = (editor) => {
  const host = createBrowserAutomationHost(editor);
  const document = valueOf<{ handle: unknown }>(host, { op: 'getDocument' }).handle;
  const body = valueOf<{ handle: unknown }>(host, { op: 'getBody', document }).handle;
  const [first] = valueOf<{ handles: unknown[] }>(host, { op: 'getParagraphs', body }).handles;
  return host.execute({
    operations: [
      {
        op: 'insertCustomNode',
        at: { paragraph: first, at: 'start' },
        tag: 'clm:field',
        text: 'ACME Ltd',
        payload: {
          namespaceUri: NS,
          rootLocalName: 'fields',
          nodeId: 'n1',
          label: 'ACME Ltd',
          data: '{"v":1}',
        },
      },
    ],
  } as never);
};

const storedNode = (pkg: OoxmlPackage) => {
  const store = findCustomXmlDataPart(pkg, pkg.mainDocumentPart, NS);
  return store ? readCustomXmlNode(pkg, store.partName, 'n1') : null;
};

/** The main story as Word reads it, past the controls' ids, which every copy mints itself. */
const bodyAsRead = (pkg: OoxmlPackage) => storyAsRead(pkg).replace(/<w:id w:val="\d+"\/>/g, '');

test("the editing API's custom node reaches the other copy and the room's file", async () => {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
  expect(insertChip(olive.editor)).toMatchObject({ ok: true });
  await settle(20);
  converged(room, [olive, oscar]);
  expect(storedNode(surface(oscar).session.currentPackage())).toEqual({
    id: 'n1',
    label: 'ACME Ltd',
    data: '{"v":1}',
  });
  const editorAlone = await alone(FILE, 'Olive', 'edit', insertChip);
  const saved = roomFile(room.room.document, FILE);
  expect(bodyAsRead(saved)).toBe(bodyAsRead(editorAlone));
  expect(storedNode(saved)).toEqual(storedNode(editorAlone));
});
