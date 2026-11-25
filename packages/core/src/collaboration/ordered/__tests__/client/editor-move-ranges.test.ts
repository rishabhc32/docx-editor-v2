// A move made in a room has its range markers on every copy, as a move a file holds has them:
// typing at the end of the moved words goes inside the move's range, as in the editor alone and
// after a fold (seed 78). Before, every copy put the typed words after the range.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { orderedViewParagraph } from '../../compose/compose.ts';
import { orderedPlaceAt } from '../../edits/places.ts';
import { encodeOrderedClientMessage } from '../../protocol/codec.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { testPerson } from '../test-persons.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  CODE,
  join,
  localRoom,
  PARTY,
  ROOM,
  sameAsRoom,
  settle,
  surface,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const FILE = docx(p(t('Cloud and serverless.')) + p(t('Read online.')));

/** The move markers, moved words and insertions of a saved file's second paragraph, in order. */
async function secondParagraph(save: () => Promise<ArrayBuffer>): Promise<string[]> {
  const xml = strFromU8(unzipSync(new Uint8Array(await save()))['word/document.xml']!);
  const second = [...xml.matchAll(/<w:p[ >](?:(?!<\/w:p>).)*<\/w:p>/gs)][1]?.[0] ?? '';
  return [...second.matchAll(/<w:(moveTo(?:Range(?:Start|End))?|ins)[ >]|<w:t(?: [^>]*)?>([^<]*)/g)]
    .map((m) => m[1] ?? m[2] ?? '')
    .filter((part) => part !== '');
}

describe("typing at the end of another person's moved words, in a room", () => {
  test("goes inside the move's range on every copy", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const review = reviewTestModule();
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    const ben = await join(room, 'Ben', { wordFile: FILE, review });

    // Cara's page sends a move of "Cloud " to the end of the second paragraph, as the room's wire
    // has it (no editor gesture makes one)
    const cara = room.transport({ person: testPerson('Cara'), ...PARTY }).connect({
      open: () => {},
      close: () => {},
      message: () => {},
    });
    const hello = { v: 3, type: 'hello', room: ROOM, file: 1, last: 0, code: CODE } as const;
    cara.send(encodeOrderedClientMessage({ ...hello, namespace: 'cara-1' }));
    await settle(4);
    const at = (index: number, offset: number) => {
      const view = orderedViewParagraph(room.room.document, BODY, index);
      if (!view) {
        throw new Error(`no paragraph ${index}`);
      }

      return orderedPlaceAt(view, offset);
    };
    const move = { v: 1, id: 'cara-1:1', kind: 'move', part: BODY, mode: 'suggest' } as const;
    const where = { from: at(0, 0), to: at(0, 6), place: at(1, 12) };
    cara.send(
      encodeOrderedClientMessage({ v: 3, type: 'changes', changes: [{ ...move, ...where }] })
    );
    await settle();

    // Ben types at the paragraph's end, right after Cara's moved words
    caretAt(ben, 1, 18);
    surface(ben).insertPlainText('X');
    await settle();
    for (const person of [ana, ben]) {
      expect(person.refusals).toEqual([]);
      expect(sameAsRoom(person, room)).toBe(true);
    }

    for (const person of [ana, ben]) {
      expect(await secondParagraph(() => person.editor.save())).toEqual([
        'Read online.',
        'moveToRangeStart',
        'moveTo',
        'Cloud ',
        'ins',
        'X',
        'moveToRangeEnd',
      ]);
    }
  });
});
