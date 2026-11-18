// The organiser's direct paragraph or run change over another person's pending one, then a
// reject of that change, with real editors and one room (audit R1 row 27).
//
// Word keeps one change record, the properties before the first pending change (`w:pPrChange`,
// `w:rPrChange`), so a direct change made after it goes with a reject: in Word, in the editor
// alone, and in the saved file reopened. The room kept the direct change live, so the live room
// and the saved file differed. Each reject is compared with the editor alone, live and on the
// saved file reopened, by the editor's own Reject All and by the host's card decision.

import { afterEach, describe, expect, test } from 'bun:test';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { alone, caretInSurface, comparedBody, converged, type Surface } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

const TEXT = 'The fee is 100 USD.';
const FILE = docx(p(t(TEXT)) + p(t('Due in 30 days.')));

async function seats(file = FILE) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

type Via = 'the editor' | 'the host';

/** Reject every pending suggestion: by the editor's Reject All, or each card by the host. */
async function rejectAll(room: OrderedLocalRoom, person: Person, via: Via): Promise<void> {
  if (via === 'the editor') {
    const command = { type: 'resolveAllReviewChanges', action: 'reject', scope: 'document' };
    expect(person.editor.exec(command as never).ok).toBe(true);
  } else
    for (const card of orderedCards(room.room.document))
      expect(person.handle.session.requestDecision(card.suggestions[0] ?? '', 'rejected')).not.toBe(
        null
      );
  await settle();
}

interface Case {
  readonly name: string;
  /** Ana's suggestion, made on her editor. */
  readonly ana: (s: Surface) => void;
  /** Olive's direct change after it, on her editor or on the editor alone. */
  readonly olive: (s: Surface) => void;
}

const CASES: readonly Case[] = [
  {
    name: 'Ana centres the paragraph, Olive indents it',
    ana: (s) => s.setParagraphProperty('jc', { val: 'center' }),
    olive: (s) => {
      caretInSurface(s, TEXT, 3);
      s.setParagraphProperty('ind', { left: '720' });
    },
  },
  {
    name: 'Ana bolds the words, Olive italics them',
    ana: (s) => s.setRunProperty('b'),
    olive: (s) => {
      const id = s.session.paragraphIds()[0] ?? '';
      s.setSelection({
        anchor: { paragraphId: id, offset: 4 },
        head: { paragraphId: id, offset: 7 },
      });
      s.setRunProperty('i');
    },
  },
];

describe("a direct change over another person's pending one goes with its reject, as alone", () => {
  for (const { name, ana: suggest, olive: direct } of CASES)
    for (const via of ['the editor', 'the host'] as const)
      test(`${name}, rejected by ${via}`, async () => {
        const live = await seats();
        caretAt(live.ana, 0, 4, 7);
        suggest(surface(live.ana));
        await settle();
        const suggested = new Uint8Array(await live.olive.editor.save());
        direct(surface(live.olive));
        await settle();
        converged(live.room, live.ana, live.olive);
        const saved = new Uint8Array(await live.olive.editor.save());
        const expected = alone(
          { author: 'Olive', mode: 'edit', review, file: suggested, decide: 'reject' },
          direct
        );
        await rejectAll(live.room, live.olive, via);
        converged(live.room, live.ana, live.olive);
        expect(comparedBody(surface(live.olive).session.part())).toBe(expected);
        closeAll();
        // The saved file, reopened in a room, rejected the same way
        const reopened = await seats(saved);
        await rejectAll(reopened.room, reopened.olive, via);
        converged(reopened.room, reopened.ana, reopened.olive);
        expect(comparedBody(surface(reopened.olive).session.part())).toBe(expected);
      });
});
