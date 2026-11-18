/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Custom nodes (data chips) in an ordered room: two real editors and one room. An insert, an
// update and a removal reach the other copy with their payload, and the room's Word file holds
// what the editor alone holds. Before, the room took every one of them back (`ungated-write`).

import { afterEach, describe, expect, test } from 'bun:test';
import { z } from 'zod';
import {
  findCustomXmlDataPart,
  readCustomXmlNode,
  writeOoxmlPackage,
  type OoxmlPackage,
} from '@docx-editor.dev/core/store';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  type Editor,
  type Person,
} from '../../../core/src/collaboration/ordered/__tests__/client/editor-harness.ts';
import {
  alone,
  converged,
  roomFile,
  storyAsRead,
} from '../../../core/src/collaboration/ordered/__tests__/client/links-harness.ts';
import { docx, p, t } from '../../../core/src/collaboration/ordered/__tests__/word-fixtures.ts';
import {
  customNodesOf,
  defineCustomNode,
  insertCustomNode,
  removeCustomNode,
  updateCustomNode,
} from '../index.ts';
import { customNodeNamespace } from '../custom-nodes/node-payload.ts';

afterEach(closeAll);

const citation = defineCustomNode({
  name: 'citation',
  tagPrefix: 'acme',
  schema: z.object({ sourceId: z.string(), year: z.number().int() }),
  text: (data) => `[${data.sourceId}, ${data.year}]`,
});
const NS = customNodeNamespace(citation);
const FILE = docx(p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')));

/** The first body paragraph's id, at `offset`. */
function at(editor: Editor, offset: number) {
  const paragraphId = editor.surface?.session.paragraphIds()[0];
  if (paragraphId === undefined) throw new Error('no paragraph');
  return { paragraphId, offset };
}

const insert = (year: number) => (editor: Editor) =>
  insertCustomNode(editor, citation, { at: at(editor, 4), data: { sourceId: 'src_9f3', year } });

const remove = (editor: Editor) => {
  const [node] = customNodesOf(editor, { nodes: [citation] });
  if (!node) throw new Error('no chip');
  return removeCustomNode(editor, node.nodeId);
};

/** The chips an editor shows, with their payloads. */
const chips = (editor: Editor) =>
  customNodesOf(editor, { nodes: [citation] }).map((node) => ({
    text: node.text,
    data: node.data,
  }));

/** The main story as Word reads it, past the controls' ids, which every copy mints itself. */
const bodyAsRead = (pkg: OoxmlPackage) => storyAsRead(pkg).replace(/<w:id w:val="\d+"\/>/g, '');

/** The payload node a package's store holds under `nodeId`, or null. */
function storedNode(pkg: OoxmlPackage, nodeId: string) {
  const store = findCustomXmlDataPart(pkg, pkg.mainDocumentPart, NS);
  return store ? readCustomXmlNode(pkg, store.partName, nodeId) : null;
}

/** The organiser's steps in a room, then the same steps in the editor alone: same file. */
async function inRoomAndAlone(
  steps: readonly ((editor: Editor) => unknown)[],
  file = FILE
): Promise<{
  people: Person[];
  saved: OoxmlPackage;
  editorAlone: OoxmlPackage;
}> {
  const room = localRoom(file);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
  const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: file });
  for (const step of steps) {
    expect(step(olive.editor)).toMatchObject({ ok: true });
    await settle(20);
  }
  converged(room, [olive, oscar]);
  const editorAlone = await alone(file, 'Olive', 'edit', (editor) => {
    for (const step of steps) step(editor);
  });
  return { people: [olive, oscar], saved: roomFile(room.room.document, file), editorAlone };
}

describe('custom nodes in an ordered room', () => {
  test("an organiser's new chip reaches the other copy with its payload, and saves", async () => {
    const { people, saved, editorAlone } = await inRoomAndAlone([insert(2024)]);
    for (const person of people)
      expect(chips(person.editor)).toEqual([
        { text: '[src_9f3, 2024]', data: { sourceId: 'src_9f3', year: 2024 } },
      ]);
    expect(bodyAsRead(saved)).toBe(bodyAsRead(editorAlone));
    expect(storedNode(saved, 'cx1')).toEqual(storedNode(editorAlone, 'cx1'));
    expect(storedNode(saved, 'cx1')?.data).toBe('{"sourceId":"src_9f3","year":2024}');
  });

  test('an update and a removal reach the other copy, and the store follows', async () => {
    const update = (editor: Editor) => {
      const [node] = customNodesOf(editor, { nodes: [citation] });
      if (!node) throw new Error('no chip');
      return updateCustomNode(editor, citation, node.nodeId, {
        data: { sourceId: 'src_9f3', year: 2025 },
      });
    };
    const updated = await inRoomAndAlone([insert(2024), update]);
    for (const person of updated.people)
      expect(chips(person.editor)).toEqual([
        { text: '[src_9f3, 2025]', data: { sourceId: 'src_9f3', year: 2025 } },
      ]);
    expect(bodyAsRead(updated.saved)).toBe(bodyAsRead(updated.editorAlone));
    expect(storedNode(updated.saved, 'cx1')).toEqual(storedNode(updated.editorAlone, 'cx1'));
    closeAll();

    const removed = await inRoomAndAlone([insert(2024), remove]);
    for (const person of removed.people) expect(chips(person.editor)).toEqual([]);
    expect(bodyAsRead(removed.saved)).toBe(bodyAsRead(removed.editorAlone));
    expect(storedNode(removed.saved, 'cx1')).toBeNull();
  });

  test("a removal of the file's own chip takes its node out of the file's store", async () => {
    const file = writeOoxmlPackage(await alone(FILE, 'Olive', 'edit', insert(2024)));
    const { people, saved, editorAlone } = await inRoomAndAlone([remove], file);
    for (const person of people) expect(chips(person.editor)).toEqual([]);
    expect(bodyAsRead(saved)).toBe(bodyAsRead(editorAlone));
    expect(storedNode(editorAlone, 'cx1')).toBeNull();
    expect(storedNode(saved, 'cx1')).toBeNull();
  });

  test('an undo takes the chip and its node back on every copy', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    expect(insert(2024)(olive.editor)).toMatchObject({ ok: true });
    await settle(20);
    expect(olive.editor.exec({ type: 'undo' }).ok).toBe(true);
    await settle(20);
    converged(room, [olive, oscar]);
    for (const person of [olive, oscar]) expect(chips(person.editor)).toEqual([]);
    expect(storedNode(roomFile(room.room.document, FILE), 'cx1')).toBeNull();
  });

  test('two people who mint the same node id at once: the second is refused by name', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    // Neither copy has seen the other's chip, so both mint `cx1`
    expect(insert(2024)(olive.editor)).toMatchObject({ ok: true });
    expect(insert(1999)(oscar.editor)).toMatchObject({ ok: true });
    await settle(20);
    expect(olive.refusals).toEqual([]);
    expect(oscar.refusals).toMatchObject([
      { code: 'ordered-refused', reason: 'pending-refused', detail: 'name-taken' },
    ]);
    oscar.refusals.length = 0;
    converged(room, [olive, oscar]);
    for (const person of [olive, oscar])
      expect(chips(person.editor)).toEqual([
        { text: '[src_9f3, 2024]', data: { sourceId: 'src_9f3', year: 2024 } },
      ]);
  });
});
