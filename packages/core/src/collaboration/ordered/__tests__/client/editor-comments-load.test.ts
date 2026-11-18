// The Word file's comments in a room with two real editors: every copy holds the file's comments
// as the room does, each editor lists the threads the editor alone lists, a room comment
// still adds its comment, and the saved file reopens to the same comments.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { commentedFile, commentsOf } from '../comments-fixture.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Editor,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();
const FILE = commentedFile();

/** An editor's comment cards: id, the comment it answers, resolved state and words. */
function commentCards(editor: Editor): string[] {
  return editor
    .getReviewItems()
    .flatMap((placement) => (placement.kind === 'comment' ? [placement.item] : []))
    .map((item) => {
      const words = item.comment.blocks.map(textOf).join('');
      return `${item.id} ${item.parentId ?? '-'} ${item.resolved} ${words}`;
    });
}

/** The words under a node. */
function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') {
    return node.value;
  }

  return node.children.map(textOf).join('');
}

/** The comment cards of the editor alone on the file. */
function aloneCards(): string[] {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Olive',
    mode: 'edit',
    modules: [review],
  });
  try {
    return commentCards(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

describe("the Word file's comments in a room", () => {
  test('every copy holds them, each editor lists them as alone, and save and reopen keep them', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      review,
      wordFile: FILE,
    });
    const ana = await join(room, 'Ana', { review, wordFile: FILE });
    const opened = openOrderedWordFile(FILE, { file: 1 });
    const fileComments = commentsOf(opened.document);
    expect(fileComments).toHaveLength(11);

    const alone = aloneCards();
    expect(alone).toContain('1 0 false Remark 1.');
    for (const person of [olive, ana]) {
      expect(commentsOf(person.handle.session.orderedDocument())).toEqual(fileComments);
      expect(commentCards(person.editor)).toEqual(alone);
    }

    // A party's suggestion, and a comment on the organiser's words
    caretAt(ana, 2, 0);
    surface(ana).insertPlainText('Now ');
    await settle();
    caretAt(olive, 0, 0, 6);
    expect(olive.handle.session.addSelectionComment('Room note.')).not.toBeNull();
    await settle();
    expect(room.room.joinSnapshot(0).map((change) => change.kind)).toEqual(['insert', 'comment']);
    for (const person of [olive, ana]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(commentsOf(person.handle.session.orderedDocument()).slice(0, 11)).toEqual(
        fileComments
      );
      expect(
        commentCards(person.editor).filter((card) => card.endsWith('Room note.'))
      ).toHaveLength(1);
    }

    const saved = orderedPackageBytes(room.room.document, opened.base);
    const reopened = openOrderedWordFile(saved, { file: 2 }).document;
    const [added, ...more] = [...reopened.comments.values()].slice(11);
    expect(more).toEqual([]);
    expect(added && textOf(added.element)).toBe('Room note.');
    reopened.comments.delete(added?.id ?? '');
    expect(commentsOf(reopened)).toEqual(fileComments);
  });
});
