// Random sessions with real editors (task 6.6, "Collaboration verification"): two or three
// people type, delete, press Enter and Backspace, format, undo and redo all over the document,
// while the room's frames wait and arrive in random amounts, so people edit on copies that are
// behind. Some sessions hold the organiser, who edits directly and undoes and redoes direct
// edits (task 6.7); a paragraph holds a link, which Enter inside splits (task 3.5); and some
// sessions add links, bookmarks and page fields (group 9.5). After each
// round, when every frame has arrived: every copy equals the room, every screen shows its
// list, and every refusal is one of the client's named reasons.

import { afterEach, describe, expect, test } from 'bun:test';
import { seededRandom } from '../random.ts';
import {
  caretAt,
  closeAll,
  drain,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  surface,
  type Person,
} from './editor-harness.ts';
import { docx, p, t } from '../word-fixtures.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('The fee is 100 USD per month.')) +
    p(t('Payment is due in 30 days.')) +
    p(
      t('Late payment adds ') +
        `<w:hyperlink r:id="rIdLink">${t('interest')}</w:hyperlink>` +
        t('.')
    ) +
    p(t('Signed by both parties.'))
);

/** Reasons a refusal may have in a random session: each names a real conflict. */
const NAMED = new Set([
  'pending-refused',
  'no-inverse',
  'redo-refused',
  'hidden-target',
  'empty-range',
  'no-paragraph-before',
  'text-differs',
  'out-of-step',
  'unknown-address',
  'unknown-suggestion',
  'invalid-range',
  'stale-revert',
  'break-in-container',
]);

function act(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const length = texts[index]?.length ?? 0;
  const at = Math.floor(rand() * (length + 1));
  const roll = rand();
  if (roll < 0.35) {
    caretAt(person, index, at);
    s.insertPlainText(['a', 'bc', 'X ', 'yes'][Math.floor(rand() * 4)] ?? 'a');
  } else if (roll < 0.5) {
    caretAt(person, index, at);
    s.deleteBackward();
  } else if (roll < 0.6) {
    const end = Math.min(length, at + 1 + Math.floor(rand() * 5));
    caretAt(person, index, at, end);
    s.deleteSelection();
  } else if (roll < 0.68) {
    caretAt(person, index, at);
    s.splitParagraph();
  } else if (roll < 0.76) {
    const end = Math.min(length, at + 1 + Math.floor(rand() * 6));
    if (end > at) {
      caretAt(person, index, at, end);
      s.toggleRunProperty(rand() < 0.5 ? 'b' : 'i');
    }
  } else if (roll < 0.88) s.undo();
  else s.redo();
}

/**
 * Group 9.5's edits (links, bookmarks, page fields): a person links words or unlinks, types a
 * page field through the editing API's lane, and the organiser adds, moves or removes a
 * bookmark (Word has no revision for a bookmark, so a party never does).
 */
function linkOrField(person: Person, rand: () => number, organiser: boolean): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const length = texts[index]?.length ?? 0;
  const at = Math.floor(rand() * (length + 1));
  const end = Math.min(length, at + 1 + Math.floor(rand() * 5));
  const roll = rand();
  if (roll < 0.3 && end > at) {
    caretAt(person, index, at, end);
    person.editor.exec({ type: 'insertHyperlink', href: 'https://example.com/r' });
  } else if (roll < 0.45) {
    caretAt(person, index, at);
    person.editor.exec({ type: 'removeHyperlink' });
  } else if (roll < 0.7 || !organiser) {
    const paragraphId = s.session.paragraphIds()[index] ?? '';
    s.applyAutomationOps(() => [{ op: 'insertPageField', paragraphId, offset: at, field: 'PAGE' }]);
  } else if (roll < 0.9) {
    caretAt(person, index, at, rand() < 0.5 ? at : end);
    person.editor.exec({ type: 'insertBookmark', name: rand() < 0.5 ? 'Alpha' : 'Beta' });
  } else person.editor.exec({ type: 'deleteBookmark', name: rand() < 0.5 ? 'Alpha' : 'Beta' });
}

async function session(
  seed: number,
  names: readonly string[],
  steps: number,
  fields = false
): Promise<void> {
  const rand = seededRandom(seed);
  const queue: (() => void)[] = [];
  const room = localRoom(FILE, { deliver: (run) => queue.push(run) });
  const people: Person[] = [];
  for (const name of names) {
    // Olga is the organiser: she edits directly
    const joining = join(room, name, {
      wordFile: FILE,
      ...(name === 'Olga' ? { mode: 'edit' as const, login: ORGANISER } : {}),
    });
    await drain(queue);
    people.push(await joining);
  }
  for (let round = 0; round < 3; round += 1) {
    for (let step = 0; step < steps; step += 1) {
      const person = people[Math.floor(rand() * people.length)];
      // A separate draw, so the sessions without them keep their own sequence
      if (person && fields && rand() < 0.25) linkOrField(person, rand, person.name === 'Olga');
      else if (person) act(person, rand);
      await Promise.resolve();
      await Promise.resolve();
      // Some frames arrive now, the rest wait: people edit on copies that are behind
      for (let k = Math.floor(rand() * 4); k > 0; k -= 1) queue.shift()?.();
      if (rand() < 0.2) await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await drain(queue);
    const label = `seed ${seed} round ${round}`;
    for (const person of people) {
      const unnamed = person.refusals.filter((r) => !NAMED.has(r.reason));
      expect({ label, who: person.name, unnamed }).toEqual({
        label,
        who: person.name,
        unnamed: [],
      });
      expect({ label, who: person.name, status: person.handle.session.status() }).toEqual({
        label,
        who: person.name,
        status: 'ready',
      });
      expect({ label, who: person.name, room: sameAsRoom(person, room) }).toEqual({
        label,
        who: person.name,
        room: true,
      });
      expect({ label, who: person.name, screen: screenFollowsList(person) }).toEqual({
        label,
        who: person.name,
        screen: true,
      });
    }
  }
  expect(room.room.order).toBeGreaterThan(steps);
}

/** More seeds for a local sweep: ORDERED_SWEEP=150 ORDERED_SWEEP_FROM=200. */
const SWEEP = Number(process.env.ORDERED_SWEEP ?? 0);
const FROM = Number(process.env.ORDERED_SWEEP_FROM ?? 100);

describe('random sessions with real editors', () => {
  const PEOPLE = [
    ['Ana', 'Ben'],
    ['Ana', 'Ben', 'Cy'],
    ['Ana', 'Olga', 'Ben'],
  ] as const;
  for (let seed = FROM; seed < FROM + SWEEP; seed += 1)
    test(`sweep, seed ${seed}`, async () => {
      await session(seed, PEOPLE[seed % 3] ?? PEOPLE[0], 30);
    });
  for (const seed of [1, 2, 3])
    test(`two people, seed ${seed}`, async () => {
      await session(seed, ['Ana', 'Ben'], 30);
    });
  test('three people, seed 7', async () => {
    await session(7, ['Ana', 'Ben', 'Cy'], 30);
  });
  test('two parties and the organiser, seed 11', async () => {
    await session(11, ['Ana', 'Olga', 'Ben'], 30);
  });
  for (const seed of [21, 22, 23])
    test(`links, bookmarks and page fields, two parties and the organiser, seed ${seed}`, async () => {
      await session(seed, ['Ana', 'Olga', 'Ben'], 30, true);
    });
});
