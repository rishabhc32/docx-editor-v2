// Repeating section items with real editors and one ordered room (task 9.6): the organiser
// adds an item (a copy of the one before it, its control back to its prompt) and removes one;
// each equals the editor alone, every copy follows, a party types in the new item, and the
// organiser's undo and redo are exact. A party's own add is a direct edit: refused by the host's
// party rule (`text`).

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { stubReviewModule } from '../../../../editor/__tests__/review-test-module.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OrderedChange } from '../../change/types.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedViewParagraph } from '../../compose/compose.ts';
import { orderedPlaceAt } from '../../edits/places.ts';
import { encodeOrderedClientMessage } from '../../protocol/codec.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { snapshot } from '../snapshot.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  CODE,
  editorTexts,
  editorXml,
  HeldTransport,
  join,
  listTexts,
  localRoom,
  ORGANISER,
  PARTY,
  ROOM,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const W15 = 'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const FILE = docx(
  p(t('Parties:')) +
    sdt(
      `<w:id w:val="20"/><w15:repeatingSection ${W15}/>`,
      sdt(
        `<w:id w:val="21"/><w15:repeatingSectionItem ${W15}/>`,
        p(t('Party: ') + sdt('<w:id w:val="22"/><w:text/>', t('Acme')) + t(' (seller)'))
      )
    ) +
    p(t('End.'))
);

const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/<w:id w:val="\d+"\/>/g, '<w:id/>');

type Edit = (person: Person) => void;

const ADD: Edit = (person) => {
  caretAt(person, 1, 2);
  person.editor.exec({ type: 'addRepeatingSectionItem' } as never);
};

function alone(...edits: Edit[]): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: FILE,
    author: 'Olive',
    mode: 'edit',
    modules: [stubReviewModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    for (const edit of edits) edit(person);
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
    expect(editorTexts(person)).toEqual(listTexts(person));
  }
}

describe('repeating section items', () => {
  test('the organiser adds an item; a party types in it; undo and redo', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const before = bodyOf(editorXml(olive));
    ADD(olive);
    await settle();
    converged(room, olive, ana);
    const after = bodyOf(editorXml(olive));
    expect(after).toBe(alone(ADD));
    expect(editorTexts(ana)).toEqual([
      'Parties:',
      'Party: Acme (seller)',
      'Party: Click or tap here to enter text. (seller)',
      'End.',
    ]);
    // A party types over the new item's prompt: a suggestion
    caretAt(ana, 2, 9);
    surface(ana).insertPlainText('Bolt');
    await settle();
    converged(room, olive, ana);
    expect(editorTexts(olive)[2]).toBe('Party: Bolt (seller)');
    const saved = new Uint8Array(await olive.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    expect(openOrderedWordFile(saved, { file: 2 }).document.suggestions.size).toBe(1);
    // The organiser's undo takes the item back, the party's words with it, and redo returns both
    surface(olive).undo();
    await settle();
    converged(room, olive, ana);
    expect(bodyOf(editorXml(olive))).toBe(before);
    surface(olive).redo();
    await settle();
    converged(room, olive, ana);
    expect(editorTexts(ana)[2]).toBe('Party: Bolt (seller)');
  });

  test('the organiser removes an item, then undo', async () => {
    const room = localRoom(FILE);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    const ben = await join(room, 'Ben', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    ADD(olive);
    await settle();
    const two = bodyOf(editorXml(olive));
    const remove: Edit = (person) =>
      void person.editor.exec({ type: 'removeRepeatingSectionItem', index: 0 } as never);
    caretAt(olive, 1, 2);
    remove(olive);
    await settle();
    converged(room, olive, ben);
    expect(bodyOf(editorXml(olive))).toBe(
      alone(ADD, (person) => {
        caretAt(person, 1, 2);
        remove(person);
      })
    );
    expect(editorTexts(ben)).toEqual([
      'Parties:',
      'Party: Click or tap here to enter text. (seller)',
      'End.',
    ]);
    surface(olive).undo();
    await settle();
    converged(room, olive, ben);
    expect(bodyOf(editorXml(olive))).toBe(two);
  });

  test("a party's new item is a direct edit: refused with the party rule's reason", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    ADD(ana);
    await settle();
    expect(ana.refusals.map((r) => r.reason)).toEqual(['text']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(screenFollowsList(ana)).toBe(true);
    expect(editorTexts(ana)).toHaveLength(3);
  });

  // The new-contracts acceptance run xdemo-short, seed 1014: a copy of words that hold a party's
  // pending move, taken back and applied again by the page's rebase, left the move's range ids
  // in the copy's revision ids, so its later Word ids differed from the room's
  test("an item that repeats a party's pending move: the organiser's rebased copy is the room's", async () => {
    const room = localRoom(FILE);
    const link = new HeldTransport(room.transport({ person: testPerson('Olive'), ...ORGANISER }));
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: FILE,
      transport: link,
    });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    // Bob's page sends his changes as the room's wire has them
    const bob = room.transport({ person: testPerson('Bob'), ...PARTY }).connect({
      open: () => {},
      close: () => {},
      message: () => {},
    });
    bob.send(
      encodeOrderedClientMessage({
        v: 3,
        type: 'hello',
        room: ROOM,
        file: 1,
        last: 0,
        code: CODE,
        namespace: 'bob-1',
      })
    );
    await settle(4);
    const at = (index: number, offset: number) => {
      const view = orderedViewParagraph(room.room.document, BODY, index);
      if (!view) throw new Error(`no paragraph ${index}`);
      return orderedPlaceAt(view, offset);
    };
    const send = (change: Record<string, unknown>) =>
      bob.send(
        encodeOrderedClientMessage({
          v: 3,
          type: 'changes',
          changes: [{ v: 1, part: BODY, mode: 'suggest', ...change } as unknown as OrderedChange],
        })
      );
    // "ies" of "Parties:" moved into the item's words: "Party: Acme (ies|seller)"
    send({ id: 'bob-1:1', kind: 'move', from: at(0, 4), to: at(0, 7), place: at(1, 13) });
    await settle();
    expect(editorTexts(olive)[1]).toBe('Party: Acme (iesseller)');
    // Bob's next change reaches the room before the organiser's new item, and reaches her after
    link.held = true;
    send({ id: 'bob-1:2', kind: 'insert', place: at(2, 3), text: '!' });
    await settle();
    ADD(olive);
    await settle();
    link.held = false;
    link.release();
    await settle();
    converged(room, olive, ana);
    expect(editorTexts(ana)[2]).toBe('Party: Click or tap here to enter text. (iesseller)');
    expect(snapshot(olive.handle.session.orderedDocument(), false)).toEqual(
      snapshot(room.room.document, false)
    );
    bob.close();
  });
});
