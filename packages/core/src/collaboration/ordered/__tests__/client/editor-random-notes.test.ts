// Random sessions with notes, equations and text boxes (group 9.9, task 9.11): the organiser and
// two parties type and delete in the body, in notes and in a text box, add footnotes and
// endnotes, delete and convert notes, set the notes' properties, rewrite and remove an equation,
// with undo and redo, while the room's frames arrive in random amounts. After each round: every
// copy equals the room, every screen shows its list (the notes parts too), and every refusal is
// one of the client's named reasons or a party rule.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { storyParagraphs } from '../../../../store/package/story-blocks.ts';
import { textboxStoriesInPart } from '../../../../store/package/textbox-stories.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { seededRandom } from '../random.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  listXml,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const item = (key: string) => {
  const found = MATRIX_ITEMS.find((candidate) => candidate.key === key);
  if (!found) throw new Error(`no ${key} item`);
  return found;
};

/** A footnote, an endnote, an equation and a text box, each in a paragraph of its own. */
const FILE = docx(
  p(t('The fee') + (item('fnref').xml ?? '') + t(' is 100 USD.')) +
    p(t('Payment is due ') + (item('math').xml ?? '') + t(' in 30 days.')) +
    bodyOf(item('textbox')) +
    p(t('Signed') + (item('enref').xml ?? '') + t(' by both parties.')),
  MATRIX_EXTRAS
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
  // The party rules that the host's earlier check also applies (the room runs the suggestions-only
  // policy)
  'nested',
  'text',
  'package',
  'others_suggestions',
]);

async function drain(queue: (() => void)[]): Promise<void> {
  for (let guard = 0; guard < 10_000; guard += 1) {
    await Promise.resolve();
    await Promise.resolve();
    const next = queue.shift();
    if (next) {
      next();
      continue;
    }
    await new Promise((resolve) => setTimeout(resolve, 0));
    if (queue.length === 0) return;
  }
  throw new Error('the session does not settle');
}

/** The ids of the notes of a kind the editor shows. */
function noteIds(person: Person, kind: 'footnote' | 'endnote'): number[] {
  const xml = serializeOoxmlPart(surface(person).session.part());
  return [...xml.matchAll(new RegExp(`<w:${kind}Reference w:id="(\\d+)"/>`, 'g'))].map((m) =>
    Number(m[1])
  );
}

const pick = <T>(items: readonly T[], rand: () => number): T | undefined =>
  items[Math.floor(rand() * items.length)];

/** Type in a note the editor shows, and come back to the body. */
function typeInNote(person: Person, rand: () => number): void {
  const kind = rand() < 0.5 ? 'footnote' : 'endnote';
  const id = pick(noteIds(person, kind), rand);
  const s = surface(person);
  if (id === undefined || !s.enterNote(`${kind}:${id}`)) return;
  s.insertPlainText(['n', 'note ', 'Z'][Math.floor(rand() * 3)] ?? 'n');
  s.setActiveScope({ kind: 'body' });
}

/** Type in the text box, at its start, and come back to the body. */
function typeInTextBox(person: Person, rand: () => number): void {
  const s = surface(person);
  const [story] = textboxStoriesInPart(s.session.part());
  const paragraph = story ? storyParagraphs(story.root)[0] : undefined;
  if (!story || !paragraph) return;
  const scope = {
    kind: 'frame' as const,
    id: story.root.id,
    drawingNodeId: story.drawingNodeId,
    hostParagraphId: story.hostParagraphId,
  };
  if (!s.setActiveScope(scope)) return;
  const point = { paragraphId: paragraph.id, offset: 0 };
  s.setSelection({ anchor: point, head: point });
  if (rand() < 0.7) s.insertPlainText('b');
  else s.deleteForward();
  s.setActiveScope({ kind: 'body' });
}

/** Rewrite or remove the equation the editor shows. */
function equation(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = texts.findIndex((text) => text.startsWith('Payment is due'));
  const at = texts[index]?.indexOf('￼') ?? -1;
  if (index < 0 || at < 0) return;
  caretAt(person, index, at + 1);
  const found = s.equations.equationAtCaret();
  if (!found) return;
  if (rand() < 0.8) s.equations.applyEquation(found.id, `y=${Math.floor(rand() * 9)}`);
  else s.equations.removeEquation(found.id);
}

function act(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const at = Math.floor(rand() * ((texts[index]?.length ?? 0) + 1));
  const roll = rand();
  if (roll < 0.18) {
    caretAt(person, index, at);
    s.insertPlainText(['a', 'bc', 'X '][Math.floor(rand() * 3)] ?? 'a');
  } else if (roll < 0.26) {
    caretAt(person, index, at);
    s.deleteBackward();
  } else if (roll < 0.3) {
    caretAt(person, index, at);
    s.splitParagraph();
  } else if (roll < 0.4) {
    caretAt(person, index, at);
    if (s.insertNote(rand() < 0.6 ? 'footnote' : 'endnote') && rand() < 0.7)
      s.insertPlainText('New note.');
    s.setActiveScope({ kind: 'body' });
  } else if (roll < 0.5) typeInNote(person, rand);
  else if (roll < 0.56) {
    const kind = rand() < 0.5 ? 'footnote' : 'endnote';
    const id = pick(noteIds(person, kind), rand);
    if (id !== undefined) person.editor.exec({ type: 'deleteNote', noteKind: kind, noteId: id });
  } else if (roll < 0.61) {
    const kind = rand() < 0.5 ? 'footnote' : 'endnote';
    const id = pick(noteIds(person, kind), rand);
    if (id !== undefined) person.editor.exec({ type: 'convertNote', fromKind: kind, noteId: id });
  } else if (roll < 0.64)
    person.editor.exec({
      type: 'setNoteProperties',
      scope: 'document',
      footnote: { numFmt: rand() < 0.5 ? 'lowerRoman' : 'decimal', numStart: 1 + at },
    });
  else if (roll < 0.7) equation(person, rand);
  else if (roll < 0.8) typeInTextBox(person, rand);
  else if (roll < 0.91) s.undo();
  else s.redo();
}

/** Every notes part a copy's editor shows is what its list composes. */
function notesFollowList(person: Person): boolean {
  const doc = person.handle.session.orderedDocument();
  const shown = surface(person).session.currentPackage();
  for (const name of doc.stories.keys()) {
    if (!/\/(foot|end)notes/.test(name)) continue;
    const part = shown.parts.get(name);
    if (part && serializeOoxmlPart(part) !== listXml(person, name)) return false;
  }
  return true;
}

async function session(seed: number, steps: number): Promise<void> {
  const rand = seededRandom(seed);
  const queue: (() => void)[] = [];
  const room = localRoom(FILE, {
    deliver: (run) => queue.push(run),
    policy: suggestionsOnlyPolicy(),
  });
  const people: Person[] = [];
  for (const name of ['Olga', 'Ana', 'Ben']) {
    // Olga is the organiser: she edits directly
    const joining = join(room, name, {
      wordFile: FILE,
      ...(name === 'Olga' ? { mode: 'edit' as const, login: ORGANISER } : { login: PARTY }),
    });
    await drain(queue);
    people.push(await joining);
  }
  for (let round = 0; round < 3; round += 1) {
    for (let step = 0; step < steps; step += 1) {
      const person = people[Math.floor(rand() * people.length)];
      if (person) act(person, rand);
      await Promise.resolve();
      await Promise.resolve();
      for (let k = Math.floor(rand() * 4); k > 0; k -= 1) queue.shift()?.();
      if (rand() < 0.3) await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await drain(queue);
    await new Promise((resolve) => setTimeout(resolve, 0));
    await drain(queue);
    const label = `seed ${seed} round ${round}`;
    for (const person of people) {
      const unnamed = person.refusals.filter((r) => !NAMED.has(r.reason));
      expect({
        label,
        who: person.name,
        unnamed,
        status: person.handle.session.status(),
        room: sameAsRoom(person, room),
        screen: screenFollowsList(person),
        notes: notesFollowList(person),
      }).toEqual({
        label,
        who: person.name,
        unnamed: [],
        status: 'ready',
        room: true,
        screen: true,
        notes: true,
      });
    }
  }
  expect(room.room.order).toBeGreaterThan(steps);
}

/** More seeds for a local sweep: ORDERED_SWEEP=50 ORDERED_SWEEP_FROM=200. */
const SWEEP = Number(process.env.ORDERED_SWEEP ?? 0);
const FROM = Number(process.env.ORDERED_SWEEP_FROM ?? 100);

describe('random sessions with notes, equations and text boxes', () => {
  for (let seed = FROM; seed < FROM + SWEEP; seed += 1)
    test(`sweep, seed ${seed}`, async () => {
      await session(seed, 30);
    });
  for (const seed of [1, 2, 3, 4])
    test(`the organiser and two parties, seed ${seed}`, async () => {
      await session(seed, 30);
    });
});
