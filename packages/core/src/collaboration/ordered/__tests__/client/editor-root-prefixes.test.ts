// A Word file that declares a shape's prefixes only on its root, loaded in an ordered room (group
// 9.10). Word writes `wps`, `v`, `o` and the rest on `w:document`, and an `mc:Choice` whose
// `Requires` names `wps` with no declaration of its own. Typing beside the shape (the
// organiser's direct edit and a party's insertion), a whole-story publish (a table) and a
// decision keep the shape, every copy follows the room, and the saved file reopens alike.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { storyParagraphs } from '../../../../store/package/story-blocks.ts';
import { textboxStoriesInPart } from '../../../../store/package/textbox-stories.ts';
import { checkSaveAndLoad } from '../sessions/session.ts';
import { packageOf } from '../word-fixtures.ts';
import {
  caretAt,
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
import { shapeSource } from './paste-harness.ts';
import { reviewTestModule } from './review-module.ts';

const review = reviewTestModule();

afterEach(closeAll);

/** A line with a text box as Word writes it: its prefixes declared on the document's root. */
const FILE = shapeSource('textbox');

const checked = (room: ReturnType<typeof localRoom>, people: readonly Person[]) => {
  for (const person of people) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(screenFollowsList(person)).toBe(true);
    // The shape and its text box stay, with the choice's prefixes resolved
    const xml = editorXml(person);
    expect(xml).toContain('<mc:Choice Requires="wps">');
    expect(xml).toContain('<wps:wsp>');
    expect(xml).toContain('Box text');
  }
};

describe("a shape whose prefixes only the file's root declares, in an ordered room", () => {
  test('typing beside it, a table and a decision keep the shape in every copy', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      review,
      wordFile: FILE,
    });
    const ana = await join(room, 'Ana', { review, wordFile: FILE });
    checked(room, [olive, ana]);

    // The organiser's direct edit just before the shape
    caretAt(olive, 0, 'Seal '.length);
    surface(olive).insertPlainText('Big ');
    await settle(16);
    checked(room, [olive, ana]);
    // A party's insertion just after it
    caretAt(ana, 0, 'Seal Big ￼'.length);
    surface(ana).insertPlainText(' red');
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorTexts(olive)[0]).toBe('Seal Big ￼ red here.');

    // A table after the line: the publisher composes the whole story again
    caretAt(olive, 0, 'Seal Big ￼ red here.'.length);
    expect(surface(olive).insertTable(1, 1)).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('<w:tbl>');

    // The organiser accepts the party's insertion
    for (const key of olive.editor
      .getReviewItems()
      .filter((placement) => placement.kind === 'revision')
      .map((placement) => placement.key))
      expect(olive.editor.acceptReviewItem(key).ok).toBe(true);
    await settle(16);
    checked(room, [olive, ana]);
    expect(editorXml(ana)).not.toContain('<w:ins ');
    checkSaveAndLoad(room.room.document, 'root prefixes', packageOf(FILE));
  });

  test('an editor that leaves and rejoins while the room moves takes the room again', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      review,
      wordFile: FILE,
    });
    const ana = await join(room, 'Ana', { review, wordFile: FILE });
    checked(room, [olive, ana]);

    // Ana leaves; the organiser types in the text box while she is away, so the line that
    // holds the box is composed again and compared through the shape's choice
    room.disconnect(ana.handle.session.namespace);
    await settle();
    expect(ana.handle.session.status()).toBe('disconnected');
    const match = olive.editor.findMatches('Box text')[0];
    if (!match?.scope || match.scope.kind !== 'frame') throw new Error('no text box match');
    expect(surface(olive).setActiveScope(match.scope)).toBe(true);
    const [box] = textboxStoriesInPart(surface(olive).session.part());
    const paragraphId = box ? storyParagraphs(box.root)[0]?.id : undefined;
    if (paragraphId === undefined) throw new Error('no text box paragraph');
    const end = { paragraphId, offset: 'Box text'.length };
    surface(olive).setSelection({ anchor: end, head: end });
    surface(olive).insertPlainText(' now');
    await settle(16);
    checked(room, [olive]);
    // and adds a table after the line: Ana's copy composes the whole story again on return
    surface(olive).setActiveScope({ kind: 'body' });
    caretAt(olive, 0, 'Seal \uFFFC here.'.length);
    expect(surface(olive).insertTable(1, 1)).toBe(true);
    await settle(16);
    checked(room, [olive]);

    // She comes back: her copy takes the room's line, the shape's paragraph compared again
    ana.handle.session.reconnect();
    await settle(16);
    expect(ana.handle.session.status()).toBe('ready');
    checked(room, [olive, ana]);
    expect(editorXml(ana)).toContain('<w:t xml:space="preserve"> now</w:t>');
    expect(editorXml(ana)).toContain('<w:tbl>');
  });
});
