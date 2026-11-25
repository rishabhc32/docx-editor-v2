// Steps in a room against the editor alone, decided live, on the save reopened and on the room
// folded once and twice.

import { expect } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { checkpointBytes, folded, mediaStore } from '../fold-check.ts';
import {
  BODY,
  editorXml,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

export type Decision = 'accept' | 'reject';
export type Act = (person: Person) => void;

/** What a check compares: the part of a body's XML a reader of the file sees. */
type View = (xml: string) => string;

/** Accept or reject every revision, as the organiser. A file with none has nothing to decide. */
function decideAll(person: Person, decision: Decision): void {
  if (!/<w:(ins|del)\b/.test(editorXml(person))) {
    return;
  }

  const command = { type: 'resolveAllReviewChanges', action: decision, scope: 'document' };
  expect(person.editor.exec(command as never).ok).toBe(true);
}

/** What the editor alone writes: after the steps, and after the decision. */
function alone(
  view: View,
  file: Uint8Array,
  steps: readonly Act[],
  decision: Decision,
  author: string
): { readonly edited: string; readonly decided: string } {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author,
    mode: 'suggesting',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    for (const act of steps) {
      act(person);
    }

    surface(person).flushPendingInput();
    const edited = view(editorXml(person));
    decideAll(person, decision);
    return { edited, decided: view(editorXml(person)) };
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The organiser, Olive, and Ben, a party, in a room on `file` (a Word file or a checkpoint). */
async function seats(file: Uint8Array, number: number) {
  const review = reviewTestModule();
  const room = localRoom(file, { file: number });
  const olive = await join(room, 'Olive', { login: ORGANISER, wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  return { room, olive, ben };
}

/** The room's view, the same on both copies. */
function roomView(view: View, seated: Awaited<ReturnType<typeof seats>>): string {
  for (const person of [seated.olive, seated.ben]) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, seated.room) }).toEqual({
      person: person.name,
      room: true,
    });
  }

  return view(editorXml(seated.olive));
}

/** The checkpoint of the room's document folded `times` times. */
async function foldedBytes(
  doc: OrderedDocument,
  file: Uint8Array,
  times: number
): Promise<Uint8Array> {
  const { base } = openOrderedWordFile(file, { file: 1 });
  const blobs = mediaStore(base);
  let copy = { document: doc, base };
  for (let fold = 1; fold < times; fold += 1) {
    copy = await folded(copy.document, copy.base, fold + 1, blobs);
  }

  return checkpointBytes(copy.document, copy.base, blobs);
}

/**
 * A check of the steps of the organiser, or of Ben, a party, in a room on `file`, decided live,
 * on the save reopened and on the room folded once and twice, each as the editor alone, by
 * `view`.
 */
export const decidedCheck =
  (view: View) =>
  async (
    label: string,
    file: Uint8Array,
    steps: readonly Act[],
    decision: Decision,
    actor: 'olive' | 'ben' = 'olive'
  ): Promise<void> => {
    const seated = await seats(file, 1);
    const expected = alone(view, file, steps, decision, seated[actor].name);
    for (const act of steps) {
      act(seated[actor]);
      await settle();
    }

    expect({ label, edited: roomView(view, seated) }).toEqual({ label, edited: expected.edited });

    // A save reopened, and the room folded once and twice, each decided by the organiser
    const saved = new Uint8Array(await seated.olive.editor.save());
    const copies = [
      ['saved', saved],
      ['folded once', await foldedBytes(seated.room.room.document, file, 1)],
      ['folded twice', await foldedBytes(seated.room.room.document, file, 2)],
    ] as const;
    decideAll(seated.olive, decision);
    await settle();
    expect({ label, live: roomView(view, seated) }).toEqual({ label, live: expected.decided });
    for (const [copy, bytes] of copies) {
      const reopened = await seats(bytes, 2);
      expect({ label, copy, edited: roomView(view, reopened) }).toEqual({
        label,
        copy,
        edited: expected.edited,
      });
      decideAll(reopened.olive, decision);
      await settle();
      const part = composeOrderedPart(reopened.room.room.document, BODY).part;
      expect({ label, copy, decided: view(serializeOoxmlPart(part)) }).toEqual({
        label,
        copy,
        decided: expected.decided,
      });
    }
  };
