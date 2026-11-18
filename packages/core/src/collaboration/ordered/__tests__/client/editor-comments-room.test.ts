// Comments travel through the room, with real editors and one room (design D15): a comment, a
// reply and a resolve are changes the room checks, stamps with the person and orders; every page
// shows the Word comments the room holds; save, reopen and a checkpoint keep them with their
// persons; the editor's own Add comment is a comment change, and its reply, resolve and delete
// stay refused (a host replies and resolves through the session).

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import type { OoxmlNode } from '../../../../store/package/ooxml-tree.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import {
  openOrderedCheckpoint,
  openOrderedRoomFile,
  openOrderedWordFile,
} from '../../client/open.ts';
import { writeOrderedCheckpoint } from '../../compose/checkpoint.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import type { OrderedPerson } from '../../model/person.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { commentedFile, end, reference, start } from '../comments-fixture.ts';
import { p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  FILE,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Editor,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const review = reviewTestModule();
const OLIVE: OrderedPerson = { id: 'id-olive-1', name: 'Olive', org: 'Host' };
const ANA: OrderedPerson = { id: 'id-ana-7f3c', name: 'Ana', org: 'Acme' };
const BEN: OrderedPerson = { id: 'id-ben-5e08', name: 'Ben', org: 'Bolt' };

/** A room with the suggestions-only policy on `file`, the organiser Olive, and parties. */
async function negotiation(file = FILE, parties: readonly OrderedPerson[] = [ANA]) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    person: OLIVE,
    mode: 'edit',
    login: ORGANISER,
    review,
    wordFile: file,
  });
  const people: Person[] = [];
  for (const person of parties) {
    people.push(await join(room, person.name, { person, review, wordFile: file }));
  }

  return { room, olive, people };
}

/** The kinds of the room's changes, in order. */
const kinds = (room: OrderedLocalRoom): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

/** The person's refusals as `code/reason`, then none. */
function refused(person: Person): string[] {
  const reasons = person.refusals.map((r) => `${r.code}/${r.reason ?? ''}`);
  person.refusals.length = 0;
  return reasons;
}

/** Every copy is the room's, nothing is refused, and every editor shows its list. */
function converged(room: OrderedLocalRoom, ...people: Person[]): void {
  for (const person of people) {
    expect({ who: person.name, refusals: person.refusals }).toEqual({
      who: person.name,
      refusals: [],
    });
    expect({ who: person.name, room: sameAsRoom(person, room) }).toEqual({
      who: person.name,
      room: true,
    });
    expect({ who: person.name, screen: screenFollowsList(person) }).toEqual({
      who: person.name,
      screen: true,
    });
  }
}

/** The words under a node. */
function textOf(node: OoxmlNode): string {
  if (node.kind === 'textValue') {
    return node.value;
  }

  return node.children.map(textOf).join('');
}

/**
 * An editor's comment cards: the author, the body's paragraphs, the author of the comment it
 * answers, whether it is resolved, and its range as `paragraph:offset-paragraph:offset`.
 */
function cardsOf(editor: Editor): string[] {
  const ids = editor.surface?.session.paragraphIds() ?? [];
  const items = editor
    .getReviewItems()
    .flatMap((placement) => (placement.kind === 'comment' ? [placement.item] : []));
  const authors = new Map(items.map((item) => [item.id, item.comment.author]));
  return items.map((item) => {
    const words = item.comment.blocks.map(textOf).join('/');
    const answers = item.parentId === undefined ? '' : ` re ${authors.get(item.parentId)}`;
    const range = item.range
      ? `${ids.indexOf(item.range.start.paragraphId)}:${item.range.start.offset}-` +
        `${ids.indexOf(item.range.end.paragraphId)}:${item.range.end.offset}`
      : 'none';
    return `${item.comment.author}: ${words}${answers}${item.resolved ? ' (resolved)' : ''} @${range}`;
  });
}

/** The comment cards of the editor alone on a Word file. */
function aloneCards(bytes: Uint8Array): string[] {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: bytes,
    author: 'Lee',
    mode: 'edit',
    modules: [review],
  });
  try {
    return cardsOf(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The comments the room added, as `author person-ID parent done words`. */
function roomComments(doc: OrderedDocument): string[] {
  return [...doc.comments.values()]
    .filter((comment) => comment.person !== null)
    .map((comment) => {
      const words = comment.element.children.map(textOf).join('/');
      return `${comment.author} ${comment.person?.id} ${comment.parent ?? '-'} ${comment.done} ${words}`;
    });
}

/** The room's document as a Word file and as a checkpoint, each with its bytes. */
async function saved(room: OrderedLocalRoom, file: Uint8Array = FILE) {
  const { base } = openOrderedRoomFile(file, { file: 1 });
  const chunks: Uint8Array[] = [];
  await writeOrderedCheckpoint(room.room.document, base, {
    media: async (sha256) => {
      throw new Error(`no media is read: ${sha256}`);
    },
    write: async (chunk) => {
      chunks.push(chunk);
    },
  });
  return {
    word: orderedPackageBytes(room.room.document, base),
    checkpoint: new Uint8Array(Buffer.concat(chunks)),
  };
}

/** The entries of a zip that hold `text`. */
const entriesWith = (bytes: Uint8Array, text: string): string[] =>
  Object.entries(unzipSync(bytes))
    .filter(([, data]) => strFromU8(data).includes(text))
    .map(([name]) => name);

/** Ana's comment on "Due" of the second paragraph, through her held selection. */
async function anaComments(room: OrderedLocalRoom, ana: Person, text = 'Why 30?\nAsk legal.') {
  caretAt(ana, 1, 0, 3);
  const id = ana.handle.session.addSelectionComment(text);
  expect(id).not.toBeNull();
  await settle();
  expect(room.room.document.comments.has(id ?? '')).toBe(true);
  return id ?? '';
}

describe('a comment through the room', () => {
  test("a party's comment shows on every page with its person; save, reopen and checkpoints keep it", async () => {
    const { room, olive, people } = await negotiation();
    const [ana] = people as [Person];
    const id = await anaComments(room, ana);
    expect(kinds(room)).toEqual(['comment']);
    converged(room, olive, ana);
    for (const person of [olive, ana]) {
      expect(cardsOf(person.editor)).toEqual(['Ana (Acme): Why 30?/Ask legal. @1:0-1:3']);
      expect(person.handle.session.orderedDocument().comments.get(id)?.person).toEqual(ANA);
    }

    // The saved Word file: the same card on the editor alone, under the Word name only
    const { word, checkpoint } = await saved(room);
    expect(aloneCards(word)).toEqual(['Ana (Acme): Why 30?/Ask legal. @1:0-1:3']);
    expect(roomComments(openOrderedWordFile(word, { file: 2 }).document)).toEqual([]);
    expect(entriesWith(word, ANA.id)).toEqual([]);

    // The checkpoint keeps the person, under its own number and as file 1 of the next room
    expect(entriesWith(checkpoint, ANA.id)).toEqual(['ordered-checkpoint.json']);
    const kept = [`Ana (Acme) ${ANA.id} - false Why 30?/Ask legal.`];
    expect(roomComments(openOrderedCheckpoint(checkpoint, { file: 1 }).document)).toEqual(kept);
    expect(roomComments(openOrderedCheckpoint(checkpoint, { file: 4 }).document)).toEqual(kept);
    const next = localRoom(checkpoint, { policy: suggestionsOnlyPolicy() });
    const again = await join(next, 'Ana', { person: ANA, review, wordFile: checkpoint });
    expect(roomComments(again.handle.session.orderedDocument())).toEqual(kept);
    expect(cardsOf(again.editor)).toEqual(['Ana (Acme): Why 30?/Ask legal. @1:0-1:3']);

    // In the next room the thread goes on
    expect(again.handle.session.replyToComment(id, 'Still open.')).not.toBeNull();
    await settle();
    expect(kinds(next)).toEqual(['comment-reply']);
    converged(next, again);
  });

  test('a reply threads under the first comment, with both ends in commentsExtended', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    const id = await anaComments(room, ana);
    const reply = ben.handle.session.replyToComment(id, 'Thirty is usual.');
    await settle();

    // A reply to the reply answers the thread's first comment
    expect(olive.handle.session.replyToComment(reply ?? '', 'Agreed.')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment-reply', 'comment-reply']);
    converged(room, olive, ana, ben);
    const cards = [
      'Ana (Acme): Why 30?/Ask legal. @1:0-1:3',
      'Ben (Bolt): Thirty is usual. re Ana (Acme) @1:0-1:3',
      'Olive (Host): Agreed. re Ana (Acme) @1:0-1:3',
    ];
    for (const person of [olive, ana, ben]) {
      expect(cardsOf(person.editor)).toEqual(cards);
    }

    const { word } = await saved(room);
    expect(aloneCards(word)).toEqual(cards);
    const extended = strFromU8(unzipSync(word)['word/commentsExtended.xml'] ?? new Uint8Array());
    const entries = [...extended.matchAll(/<w15:commentEx [^>]*\/>/g)].map((match) => match[0]);
    expect(entries).toHaveLength(3);
    expect(entries.filter((entry) => entry.includes('w15:paraIdParent'))).toHaveLength(2);
  });

  test("the organiser resolves with a note; a party's resolve, and a reply after it, are refused", async () => {
    const { room, olive, people } = await negotiation();
    const [ana] = people as [Person];
    const id = await anaComments(room, ana);

    // The party's copy takes it at once, and gives it back when the room refuses it
    expect(ana.handle.session.resolveComment(id)).toBe(true);
    await settle();
    expect(refused(ana)).toEqual(['ordered-refused/resolve']);
    expect(kinds(room)).toEqual(['comment']);
    converged(room, olive, ana);

    expect(olive.handle.session.resolveComment(id, 'Settled at 30.')).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment-resolve']);
    converged(room, olive, ana);
    const cards = [
      'Ana (Acme): Why 30?/Ask legal. (resolved) @1:0-1:3',
      'Olive (Host): Settled at 30. re Ana (Acme) (resolved) @1:0-1:3',
    ];
    expect(cardsOf(ana.editor)).toEqual(cards);
    expect(aloneCards((await saved(room)).word)).toEqual(cards);

    expect(ana.handle.session.replyToComment(id, 'One more thing.')).not.toBeNull();
    await settle();
    expect(refused(ana)).toEqual(['ordered-refused/resolved']);
    expect(olive.handle.session.resolveComment(id)).toBe(true);
    await settle();
    expect(refused(olive)).toEqual(['ordered-refused/resolved']);
    expect(kinds(room)).toEqual(['comment', 'comment-resolve']);
    converged(room, olive, ana);
  });

  test("a reply to and a resolve of the Word file's own thread", async () => {
    const file = commentedFile();
    const { room, olive, people } = await negotiation(file);
    const [ana] = people as [Person];
    const before = aloneCards(file);

    // Comment 1 answers comment 0: the reply answers comment 0 too
    expect(ana.handle.session.replyToComment('1', 'Me too.')).not.toBeNull();
    await settle();
    expect(olive.handle.session.resolveComment('0')).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['comment-reply', 'comment-resolve']);
    converged(room, olive, ana);
    const thread = room.room.document.comments;
    expect([...thread.values()].filter((comment) => comment.done).map((c) => c.parent)).toEqual([
      null,
      '0',
      '0',
      '0',
    ]);
    const cards = cardsOf(ana.editor);
    expect(cards).toHaveLength(before.length + 1);
    expect(cards.filter((card) => card.includes('(resolved)'))).toHaveLength(4);
    expect(aloneCards((await saved(room, file)).word)).toEqual(cards);
  });

  test('a reply to a file thread whose reference sits in the next paragraph shows on every page', async () => {
    const body = p(start(0) + t('Clause one.') + end(0)) + p(t('Its second line.') + reference(0));
    const file = commentedFile(body);
    const { room, olive, people } = await negotiation(file);
    const [ana] = people as [Person];

    // The reply's reference goes after the thread's, in the second paragraph
    expect(ana.handle.session.replyToComment('0', 'Me too.')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['comment-reply']);
    converged(room, olive, ana);
  });

  test("a comment on another person's pending suggestion is no nested edit", async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    caretAt(ana, 1, 0);
    surface(ana).insertPlainText('Now ');
    await settle();
    caretAt(ben, 1, 0, 4);
    expect(ben.handle.session.addSelectionComment('Why now?')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['insert', 'comment']);
    converged(room, olive, ana, ben);
    expect(cardsOf(olive.editor)).toEqual(['Ben (Bolt): Why now? @1:0-1:4']);
  });
});

describe("a comment's words deleted", () => {
  test('a direct deletion keeps the comment at its place, a range of no words', async () => {
    const { room, olive, people } = await negotiation();
    const [ana] = people as [Person];
    await anaComments(room, ana, 'Due when?');
    caretAt(olive, 1, 0, 4);
    surface(olive).deleteSelection();
    await settle();
    expect(kinds(room)).toEqual(['comment', 'delete']);
    converged(room, olive, ana);
    for (const person of [olive, ana]) {
      expect(cardsOf(person.editor)).toEqual(['Ana (Acme): Due when? @1:0-1:0']);
    }

    expect(aloneCards((await saved(room)).word)).toEqual(['Ana (Acme): Due when? @1:0-1:0']);
  });

  test('a tracked deletion keeps it over the struck words; accepted, at their place', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    await anaComments(room, ana, 'Due when?');
    caretAt(ben, 1, 0, 4);
    surface(ben).deleteSelection();
    await settle();
    converged(room, olive, ana, ben);
    expect(cardsOf(olive.editor)).toEqual(['Ana (Acme): Due when? @1:0-1:3']);
    const [deletion] = [...room.room.document.suggestions.values()].filter(
      (s) => s.status === 'pending'
    );
    expect(olive.handle.session.requestDecision(deletion?.id ?? '', 'accepted')).not.toBeNull();
    await settle();
    converged(room, olive, ana, ben);
    expect(cardsOf(ana.editor)).toEqual(['Ana (Acme): Due when? @1:0-1:0']);
  });

  test('a tracked deletion rejected gives the comment its words back', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    await anaComments(room, ana, 'Due when?');
    caretAt(ben, 1, 0, 4);
    surface(ben).deleteSelection();
    await settle();
    const [deletion] = [...room.room.document.suggestions.values()].filter(
      (s) => s.status === 'pending'
    );
    expect(olive.handle.session.requestDecision(deletion?.id ?? '', 'rejected')).not.toBeNull();
    await settle();
    converged(room, olive, ana, ben);
    expect(cardsOf(ben.editor)).toEqual(['Ana (Acme): Due when? @1:0-1:3']);
  });
});

describe('comments at the same time, and after a restart', () => {
  test('two comments on overlapping words are two threads', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    caretAt(ana, 1, 0, 6);
    caretAt(ben, 1, 4, 9);
    expect(ana.handle.session.addSelectionComment('First.')).not.toBeNull();
    expect(ben.handle.session.addSelectionComment('Second.')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment']);
    converged(room, olive, ana, ben);
    const cards = ['Ana (Acme): First. @1:0-1:6', 'Ben (Bolt): Second. @1:4-1:9'];
    expect(cardsOf(olive.editor)).toEqual(cards);
    expect(aloneCards((await saved(room)).word)).toEqual(cards);
  });

  test('two comments on the same words stay two threads', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    caretAt(ana, 1, 0, 3);
    caretAt(ben, 1, 0, 3);
    expect(ana.handle.session.addSelectionComment('First.')).not.toBeNull();
    expect(ben.handle.session.addSelectionComment('Second.')).not.toBeNull();
    await settle();
    converged(room, olive, ana, ben);
    const cards = ['Ana (Acme): First. @1:0-1:3', 'Ben (Bolt): Second. @1:0-1:3'];
    expect(cardsOf(olive.editor)).toEqual(cards);
    expect(aloneCards((await saved(room)).word)).toEqual(cards);
  });

  test('a reply the room orders after the resolve is refused, and every copy ends equal', async () => {
    const { room, olive, people } = await negotiation();
    const [ana] = people as [Person];
    const id = await anaComments(room, ana);
    expect(olive.handle.session.resolveComment(id)).toBe(true);
    expect(ana.handle.session.replyToComment(id, 'Wait.')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment-resolve']);
    expect(refused(ana)).toEqual(['ordered-refused/resolved']);
    converged(room, olive, ana);
    expect(cardsOf(ana.editor)).toEqual(['Ana (Acme): Why 30?/Ask legal. (resolved) @1:0-1:3']);
  });

  test('a comment the room lost in a restart is sent again and applies once', async () => {
    const room = localRoom(FILE, { autoSave: false, policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { person: ANA, review });
    const ben = await join(room, 'Ben', { person: BEN, review });
    await anaComments(room, ana);
    expect(room.room.order).toBe(1);
    room.restart(0);
    await settle();
    ana.handle.session.reconnect();
    ben.handle.session.reconnect();
    await settle(12);
    expect(kinds(room)).toEqual(['comment']);
    converged(room, ana, ben);
    expect(cardsOf(ben.editor)).toEqual(['Ana (Acme): Why 30?/Ask legal. @1:0-1:3']);
  });
});

describe("the editor's own comment writes", () => {
  test("the editor's Add comment is a comment change; its reply, resolve and delete are refused", async () => {
    const { room, olive, people } = await negotiation();
    const [ana] = people as [Person];
    caretAt(ana, 1, 0, 3);
    expect(ana.editor.addComment('From the editor.').ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['comment']);
    converged(room, olive, ana);
    expect(cardsOf(olive.editor)).toEqual(['Ana (Acme): From the editor. @1:0-1:3']);

    // A comment never enters the undo history
    expect(ana.handle.session.canUndo()).toBe(false);

    const [card] = olive.editor
      .getReviewItems()
      .filter((placement) => placement.kind === 'comment');
    const key = card?.key ?? '';
    expect(olive.editor.replyToReviewItem(key, 'From the card.').ok).toBe(false);
    expect(olive.editor.setCommentResolved(key, true).ok).toBe(false);
    expect(olive.editor.deleteReviewItem(key).ok).toBe(false);
    await settle();
    // They name no comment of the room: the editor gives the gate no ops for them
    expect(refused(olive)).toEqual([
      'ordered-opaque-write/opaque-write',
      'ordered-opaque-write/opaque-write',
      'ordered-opaque-write/opaque-write',
    ]);
    expect(kinds(room)).toEqual(['comment']);
    converged(room, olive, ana);
  });
});
