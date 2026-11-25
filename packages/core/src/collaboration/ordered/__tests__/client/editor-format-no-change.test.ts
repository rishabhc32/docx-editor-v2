// A format press that changes no value, with real editors and one room (audit R1 row 26). A
// party's bold made before the organiser's direct bold reached the party arrives at words that
// already have the value. The editor alone writes no change record for such a press: the party's
// change leaves no record, so the organiser's later direct italic does not become part of it
// (which a reject would then take away).

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  HeldTransport,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  settle,
  type Person,
} from './editor-harness.ts';
import {
  alone,
  exec,
  PLAIN_FILE,
  revisionAuthors,
  roomBody,
  select,
  type Step,
} from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** A toggle press on the word "Bold". */
const onBold =
  (mark: 'bold' | 'italic'): Step =>
  (editor) => {
    select(editor, 6, 10);
    exec(editor, { type: 'toggleMark', mark });
  };

/** `first` presses Bold, then `second` the same before the first press reached them. */
async function samePress(first: Person, second: Person, held: HeldTransport): Promise<void> {
  held.held = true;
  onBold('bold')(first.editor);
  await settle();
  onBold('bold')(second.editor);
  await settle();
  held.held = false;
  held.release();
  await settle();
}

describe('a format the words already have when it arrives (audit R1 row 26)', () => {
  test("a party's bold over a direct bold leaves no record, so a later direct italic is the organiser's alone", async () => {
    const room = localRoom(PLAIN_FILE, { policy: suggestionsOnlyPolicy() });
    const held = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
    const ana = await join(room, 'Ana', { wordFile: PLAIN_FILE, transport: held, review });
    const olive = await join(room, 'Olive', {
      wordFile: PLAIN_FILE,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    await samePress(olive, ana, held);
    converged(room, ana, olive);
    expect(revisionAuthors(olive)).toEqual([]);
    onBold('italic')(olive.editor);
    await settle();
    converged(room, ana, olive);
    expect(revisionAuthors(olive)).toEqual([]);
    expect(await roomBody(olive)).toBe(
      await alone(PLAIN_FILE, review, [
        { author: 'Olive', mode: 'edit', step: onBold('bold') },
        { author: 'Olive', mode: 'edit', step: onBold('italic') },
      ])
    );
  });

  // Ana's bold on her own new words writes no change record. Olive's suggested bold, made before
  // Ana's reached her, arrives at words that show it: it makes no record, as the editor alone's
  // press over those words would not
  test("an organiser's bold over a party's own bold on her new words leaves no record", async () => {
    const room = localRoom(PLAIN_FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: PLAIN_FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Olive'), ...ORGANISER }));
    const olive = await join(room, 'Olive', {
      wordFile: PLAIN_FILE,
      transport: held,
      login: ORGANISER,
      review,
    });
    select(ana.editor, 6, 6);
    exec(ana.editor, { type: 'insertText', text: 'net ' });
    await settle();
    held.held = true;
    select(ana.editor, 6, 10);
    exec(ana.editor, { type: 'toggleMark', mark: 'bold' });
    await settle();
    select(olive.editor, 6, 10);
    exec(olive.editor, { type: 'toggleMark', mark: 'bold' });
    await settle();
    held.held = false;
    held.release();
    await settle();
    converged(room, ana, olive);
    const pending = [...room.room.document.suggestions.values()].filter(
      (s) => s.status === 'pending'
    );
    expect(pending.map((s) => s.attribution.author)).toEqual(['Ana']);
  });
});
