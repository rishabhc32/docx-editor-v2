// Comments travel through the room, with real editors and one room (design D15): a comment, a
// reply and a resolve are changes the room checks, stamps with the person and orders; every page
// shows the Word comments the room holds; save, reopen and a checkpoint keep them with their
// persons; the editor's own Add comment is a comment change, and its reply, resolve and delete
// stay refused (a host replies and resolves through the session).

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
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
  editorTexts,
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

/** What `read` gives from the editor alone on a Word file. */
async function alone<T>(bytes: Uint8Array, read: (editor: Editor) => T | Promise<T>): Promise<T> {
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
    return await read(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The comment cards of the editor alone on a Word file. */
const aloneCards = (bytes: Uint8Array): Promise<string[]> => alone(bytes, cardsOf);

/**
 * The Word file the editor alone saves after it replies each of `replies` to comment 0: the
 * file's, or `comment` when given, which it first adds on "Due" of the second paragraph.
 */
function aloneReplies(
  bytes: Uint8Array,
  replies: readonly string[],
  comment?: string
): Promise<Uint8Array> {
  return alone(bytes, async (editor) => {
    if (comment !== undefined) {
      const paragraphId = editor.surface?.session.paragraphIds()[1] ?? '';
      editor.surface?.setSelection({
        anchor: { paragraphId, offset: 0 },
        head: { paragraphId, offset: 3 },
      });
      expect(editor.addComment(comment).ok).toBe(true);
    }

    for (const text of replies) {
      const card = editor
        .getReviewItems()
        .find((placement) => placement.kind === 'comment' && placement.item.id === '0');
      expect(editor.replyToReviewItem(card?.key ?? '', text).ok).toBe(true);
    }

    return new Uint8Array(await editor.save());
  });
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

/** A Word file's comments, each named by its words. */
function commentWords(files: Record<string, Uint8Array>): Map<string, string> {
  const part = strFromU8(files['word/comments.xml'] ?? new Uint8Array());
  const comments = part.matchAll(/<w:comment [^>]*w:id="(\d+)"[^>]*>(.*?)<\/w:comment>/gs);
  return new Map(
    [...comments].map((match) => {
      const words = [...(match[2] ?? '').matchAll(/<w:t[^>]*>([^<]*)/g)].map((t) => t[1]);
      return [match[1] ?? '', words.join('')];
    })
  );
}

/** The body's comment markers in order, as `RangeStart:words`, each comment named by its words. */
function markersOf(word: Uint8Array): string[] {
  const files = unzipSync(word);
  const words = commentWords(files);
  const body = strFromU8(files['word/document.xml'] ?? new Uint8Array());
  return [...body.matchAll(/<w:comment(RangeStart|RangeEnd|Reference) w:id="(\d+)"/g)].map(
    (match) => `${match[1]}:${words.get(match[2] ?? '')}`
  );
}

/**
 * A thread's markers on one range, as the engine writes them: the starts and the references in
 * the thread's order, the ends the other way.
 */
const nested = (thread: readonly string[]): string[] => [
  ...thread.map((words) => `RangeStart:${words}`),
  ...[...thread].reverse().map((words) => `RangeEnd:${words}`),
  ...thread.map((words) => `Reference:${words}`),
];

/** `word` with its comments in the order of the body's references, as Word's save writes them. */
function inMarkerOrder(word: Uint8Array): Uint8Array {
  const files = unzipSync(word);
  const body = strFromU8(files['word/document.xml'] ?? new Uint8Array());
  const order = [...body.matchAll(/<w:commentReference w:id="(\d+)"/g)].map((match) => match[1]);
  const part = strFromU8(files['word/comments.xml'] ?? new Uint8Array());
  const comments = [...part.matchAll(/<w:comment [^>]*w:id="(\d+)"[^>]*>.*?<\/w:comment>/gs)];
  const from = comments[0]?.index ?? 0;
  const last = comments.at(-1);
  const to = last === undefined ? from : (last.index ?? 0) + last[0].length;
  const sorted = [...comments].sort((a, b) => order.indexOf(a[1]) - order.indexOf(b[1]));
  const xml = sorted.map((match) => match[0]).join('');
  files['word/comments.xml'] = strToU8(part.slice(0, from) + xml + part.slice(to));
  return zipSync(files);
}

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
    expect(await aloneCards(word)).toEqual(['Ana (Acme): Why 30?/Ask legal. @1:0-1:3']);
    expect(roomComments(openOrderedWordFile(word, { file: 2 }).document)).toEqual([]);
    expect(entriesWith(word, ANA.id)).toEqual([]);

    // The checkpoint keeps the person, under its own number and as file 1 of the next room
    expect(entriesWith(checkpoint, ANA.id)).toEqual(['ordered/checkpoint.json']);
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
    expect(await aloneCards(word)).toEqual(cards);
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
    expect(await aloneCards((await saved(room)).word)).toEqual(cards);

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
    const before = await aloneCards(file);

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
    expect(await aloneCards((await saved(room, file)).word)).toEqual(cards);
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

    expect(await aloneCards((await saved(room)).word)).toEqual(['Ana (Acme): Due when? @1:0-1:0']);
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
    expect(await aloneCards((await saved(room)).word)).toEqual(cards);
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
    expect(await aloneCards((await saved(room)).word)).toEqual(cards);
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

describe("a thread's order", () => {
  const THREAD = ['Why 30?', 'Thirty is usual.', 'Agreed.', 'Done.'];

  /** Ana's comment, Ben's and Ana's replies, and Olive's resolve with a note: one thread. */
  async function threadOfThree() {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    const id = await anaComments(room, ana, 'Why 30?');
    expect(ben.handle.session.replyToComment(id, 'Thirty is usual.')).not.toBeNull();
    await settle();
    expect(ana.handle.session.replyToComment(id, 'Agreed.')).not.toBeNull();
    await settle();
    expect(olive.handle.session.resolveComment(id, 'Done.')).toBe(true);
    await settle();
    converged(room, olive, ana, ben);
    return { room, people: [olive, ana, ben] };
  }

  test('the replies follow each other as the editor alone writes them, on every page', async () => {
    const { room, people } = await threadOfThree();
    const { word } = await saved(room);
    expect(markersOf(word)).toEqual(nested(THREAD));

    // A resolve's note is its last reply, and a resolve moves no marker
    expect(markersOf(await aloneReplies(FILE, THREAD.slice(1), THREAD[0]))).toEqual(nested(THREAD));
    for (const person of people) {
      expect(markersOf(new Uint8Array(await person.editor.save()))).toEqual(nested(THREAD));
    }

    // Saved and opened again
    const next = localRoom(word, { policy: suggestionsOnlyPolicy() });
    expect(markersOf((await saved(next, word)).word)).toEqual(nested(THREAD));
  });

  test("Word's save, which writes the comments in the body's order, keeps the dates", async () => {
    const { room } = await threadOfThree();
    const reloaded = openOrderedWordFile(inMarkerOrder((await saved(room)).word), { file: 2 });
    const comments = [...reloaded.document.comments.values()];
    expect(comments.map((comment) => textOf(comment.element))).toEqual(THREAD);
    const dates = comments.map((comment) => comment.date ?? '');
    expect(dates).toEqual([...dates].sort());
  });

  test("a reply to the Word file's thread follows its replies, as the editor alone", async () => {
    const file = commentedFile();
    const { room, olive, people } = await negotiation(file);
    const [ana] = people as [Person];
    expect(ana.handle.session.replyToComment('0', 'Me too.')).not.toBeNull();
    await settle();
    converged(room, olive, ana);
    const markers = markersOf((await saved(room, file)).word);
    expect(markers.slice(0, 4)).toEqual([
      'RangeStart:Remark 0.',
      'RangeStart:Remark 1.',
      'RangeStart:Remark 2.',
      'RangeStart:Me too.',
    ]);
    expect(markers).toEqual(markersOf(await aloneReplies(file, ['Me too.'])));
  });

  test('replies from two pages at once take the room order on every page', async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    const id = await anaComments(room, ana, 'Why 30?');
    expect(ben.handle.session.replyToComment(id, 'From Ben.')).not.toBeNull();
    expect(olive.handle.session.replyToComment(id, 'From Olive.')).not.toBeNull();
    await settle();
    expect(kinds(room)).toEqual(['comment', 'comment-reply', 'comment-reply']);
    converged(room, olive, ana, ben);
    const thread = [...room.room.document.comments.values()].map((comment) =>
      textOf(comment.element)
    );
    expect([...thread].sort()).toEqual(['From Ben.', 'From Olive.', 'Why 30?']);
    expect(markersOf((await saved(room)).word)).toEqual(nested(thread));
    for (const person of [olive, ana, ben]) {
      expect(markersOf(new Uint8Array(await person.editor.save()))).toEqual(nested(thread));
    }
  });

  test("a page's undo of its own typing leaves the replies as they are", async () => {
    const { room, olive, people } = await negotiation(FILE, [ANA, BEN]);
    const [ana, ben] = people as [Person, Person];
    const id = await anaComments(room, ana, 'Why 30?');
    caretAt(ben, 0, 0);
    surface(ben).insertPlainText('Note: ');
    await settle();
    expect(ben.handle.session.replyToComment(id, 'From Ben.')).not.toBeNull();
    await settle();
    expect(olive.handle.session.replyToComment(id, 'From Olive.')).not.toBeNull();
    await settle();
    surface(ben).undo();
    await settle();
    converged(room, olive, ana, ben);
    expect(editorTexts(ben)[0]).toBe('The fee is 100 USD.');
    expect(ben.handle.session.canUndo()).toBe(false);
    const thread = ['Why 30?', 'From Ben.', 'From Olive.'];
    expect(markersOf((await saved(room)).word)).toEqual(nested(thread));
    expect(markersOf(new Uint8Array(await ben.editor.save()))).toEqual(nested(thread));
  });
});
