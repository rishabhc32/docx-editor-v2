// Enter, then typing, while the host reads the page geometry on every change, with real editors
// and one ordered room. The room's version of an Enter names the new paragraph by its room
// address, so the editor's caret moves off the engine's own new paragraph (`#new:0`). A host
// read during that change lays the pages out before the session has said where the caret goes;
// the caret then stayed on `#new:0`, which the room does not know, and every later key was
// refused as `out-of-step`. The room's echo of the person's own Enter changes nothing here (one
// clock for the room and the people), so no later change moved the caret either.

import { afterEach, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import {
  BODY,
  caretAt,
  caretOf,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const AT = Date.UTC(2026, 9, 28, 9, 0, 0);

async function seats() {
  const room = localRoom(undefined, { now: () => AT });
  const ana = await join(room, 'Ana', { now: () => AT, review: reviewTestModule() });
  const org = await join(room, 'Org', {
    mode: 'edit',
    login: ORGANISER,
    now: () => AT,
    review: reviewTestModule(),
  });
  return { room, ana, org };
}

/** The host reads the page geometry on every change, as a page placing its cards does. */
function readsGeometry(person: Person): void {
  person.editor.on('change', () => {
    person.editor.getPageGeometry();
  });
}

for (const who of ['party', 'organiser'] as const) {
  test(`Enter at a paragraph end, then typing, under a host read (${who})`, async () => {
    const { room, ana, org } = await seats();
    const person = who === 'party' ? ana : org;
    readsGeometry(person);
    caretAt(person, 0, editorTexts(person)[0]!.length);
    surface(person).splitParagraph();
    await settle();
    expect(caretOf(person)).toEqual({ index: 1, offset: 0, text: '' });
    surface(person).insertPlainText('Q');
    await settle();
    expect(person.refusals).toEqual([]);
    for (const one of [ana, org]) {
      expect(editorTexts(one).slice(0, 3)).toEqual(['The fee is 100 USD.', 'Q', 'Due in 30 days.']);
      expect(sameAsRoom(one, room)).toBe(true);
      expect(screenFollowsList(one)).toBe(true);
    }

    // The saved file reopens with the new paragraph and its letter
    const saved = new Uint8Array(await org.editor.save());
    const again = openOrderedWordFile(saved, { file: 2 }).document;
    expect(serializeOoxmlPart(composeOrderedPart(again, BODY).part)).toContain('>Q<');
  });
}
