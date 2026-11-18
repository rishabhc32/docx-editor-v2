// New lists and a change of suggestions' author from the editor in an ordered room (task 9.10).
// A list on a paragraph outside every list makes a new list definition: a `new-list` change
// whose definition every copy numbers alike, a direct edit for the organiser and a suggestion
// for a party; two people who make new lists at once get two lists. No one changes the author of
// a suggestion, which comes from its person (`fixed-author`); a party may change only the date of
// their own (the party rule).

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedPackageBytes } from '../../compose/package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  caretAt,
  closeAll,
  editorXml,
  FILE,
  HeldTransport,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const review = reviewTestModule();

const kinds = (room: ReturnType<typeof localRoom>): string[] =>
  room.room.joinSnapshot(0).map((change) => change.kind);

function numbering(room: ReturnType<typeof localRoom>): string {
  const base = readOoxmlPackage(FILE);
  if (!base.ok) throw new Error('no base');
  const bytes = orderedPackageBytes(room.room.document, base.package);
  openOrderedWordFile(bytes, { file: 2 });
  const pkg = readOoxmlPackage(bytes);
  if (!pkg.ok) throw new Error('the saved file does not open');
  const part = [...pkg.package.parts.values()].find((x) => x.name.endsWith('/numbering.xml'));
  return part ? JSON.stringify(part.root) : '';
}

const checked = (room: ReturnType<typeof localRoom>, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
  }
};

const cards = (person: Person): string[] =>
  person.editor
    .getReviewItems()
    .filter((placement) => placement.kind === 'revision')
    .map((placement) => placement.key);

describe('new lists in an ordered room', () => {
  test("the organiser's numbered list is a direct new list; undo and redo turn it", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(olive, 1, 0);
    expect(olive.editor.exec({ type: 'toggleList', kind: 'ordered' }).ok).toBe(true);
    await settle();
    expect(kinds(room)).toContain('new-list');
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('numPr');
    expect(numbering(room)).toContain('"abstractNum"');
    expect(olive.handle.session.undo()).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorXml(ana)).not.toContain('numPr');
    expect(numbering(room)).not.toContain('"abstractNum"');
    expect(olive.handle.session.redo()).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('numPr');
  });

  test("a party's bullets are a suggestion with Word's change record; a reject takes the list", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ana, 0, 0);
    expect(ana.editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
    await settle();
    expect(kinds(room)).toContain('new-list');
    checked(room, [olive, ana]);
    expect(editorXml(olive)).toContain('pPrChange');
    const [card] = cards(olive);
    if (card === undefined) throw new Error('no card');
    expect(olive.editor.rejectReviewItem(card).ok).toBe(true);
    await settle();
    checked(room, [olive, ana]);
    expect(editorXml(olive)).not.toContain('numPr');
    expect(numbering(room)).not.toContain('"abstractNum"');
  });

  test('two people make new lists at once: two lists, the same on every copy', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
    const ana = await join(room, 'Ana', { review, transport: held });
    held.held = true;
    caretAt(olive, 0, 0);
    olive.editor.exec({ type: 'toggleList', kind: 'ordered' });
    caretAt(ana, 2, 0);
    ana.editor.exec({ type: 'toggleList', kind: 'bullet' });
    await settle();
    held.held = false;
    held.release();
    await settle();
    expect(kinds(room).filter((kind) => kind === 'new-list')).toHaveLength(2);
    checked(room, [olive, ana]);
    const numbers = new Set(
      [...editorXml(olive).matchAll(/numId w:val="(\d+)"/g)].map((m) => m[1])
    );
    expect(numbers.size).toBe(2);
  });
});

describe("a change of suggestions' author in an ordered room", () => {
  test("the organiser may not give a party's suggestions another author: refused by name", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    const done = olive.editor.exec({
      type: 'setReviewChangesAuthor',
      author: 'Ana (Acme Ltd)',
      scope: 'document',
      authors: ['Ana'],
    });
    expect(done.ok).toBe(false);
    await settle();
    expect(olive.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'fixed-author' }),
    ]);
    olive.refusals.length = 0;
    expect(kinds(room)).toEqual(['insert']);
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('w:author="Ana"');
    expect(editorXml(ana)).not.toContain('Acme Ltd');
  });

  test("the organiser changes the date of a party's suggestions, for good", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    const done = olive.editor.exec({
      type: 'setReviewChangesAuthor',
      author: 'Ana',
      date: '2026-10-29T10:00:00Z',
      scope: 'document',
      authors: ['Ana'],
    });
    expect(done.ok).toBe(true);
    await settle();
    expect(kinds(room)).toEqual(['insert', 'attribution']);
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('w:author="Ana" w:date="2026-10-29T10:00:00Z"');
    // The room writes it for good, as a decision: its undo is refused by name
    expect(olive.handle.session.undo()).toBe(false);
    expect(olive.refusals).toEqual([expect.objectContaining({ reason: 'no-inverse' })]);
  });

  test("a party may change only the date of their own suggestions (the host's earlier check)", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, review });
    const ana = await join(room, 'Ana', { review });
    caretAt(ana, 0, 4);
    surface(ana).insertPlainText('new ');
    await settle();
    ana.editor.exec({
      type: 'setReviewChangesAuthor',
      author: 'Ana',
      date: '2026-10-29T10:00:00Z',
      scope: 'document',
      authors: ['Ana'],
    });
    await settle();
    expect(ana.refusals).toEqual([]);
    expect(kinds(room)).toEqual(['insert', 'attribution']);
    ana.editor.exec({
      type: 'setReviewChangesAuthor',
      author: 'Olive',
      scope: 'document',
      authors: ['Ana'],
    });
    await settle();
    expect(ana.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'fixed-author' }),
    ]);
    expect(kinds(room)).toEqual(['insert', 'attribution']);
    for (const person of [olive, ana]) expect(sameAsRoom(person, room)).toBe(true);
  });
});
