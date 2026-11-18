// Random sessions with pictures (group 9.4, task 9.11): the organiser and two parties type,
// delete, press Enter, insert pictures and text boxes, delete pictures, and the organiser
// resizes, wraps, turns, describes and replaces them, with undo and redo, while the room's
// frames arrive in random amounts. After each round: every copy equals the room, every screen
// shows its list, every picture a copy shows has its bytes, and every refusal is one of the
// client's named reasons.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  validateRasterHeader,
  type ImageDecodePort,
} from '../../../../store/package/image-resources.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { seededRandom } from '../random.ts';
import { docx, MATRIX_EXTRAS, p, pictureRun, RED_PNG, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  surface,
  type Person,
} from './editor-harness.ts';

afterEach(closeAll);

const FILE = docx(
  p(t('The fee is ') + pictureRun(5) + t(' 100 USD per month.')) +
    p(t('Payment is due in 30 days.')) +
    p(t('Signed by both parties.')),
  MATRIX_EXTRAS
);

const GIF = Uint8Array.from(atob('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'), (c) =>
  c.charCodeAt(0)
);

const decodePort = (): ImageDecodePort => ({
  async decode(bytes, mime) {
    const header = validateRasterHeader(bytes, mime);
    if (!header) throw new Error('not an image');
    return { pixelWidth: header.pixelWidth, pixelHeight: header.pixelHeight, dpiX: 96, dpiY: 96 };
  },
});

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

/** Select a picture the editor shows, as a click on it does; false when it shows none. */
function selectPicture(person: Person, rand: () => number): boolean {
  const texts = editorTexts(person);
  const places: [number, number][] = [];
  texts.forEach((text, index) => {
    for (let k = 0; k < text.length; k += 1) if (text[k] === '￼') places.push([index, k]);
  });
  const place = places[Math.floor(rand() * places.length)];
  if (!place) return false;
  caretAt(person, place[0], place[1] + 1);
  return person.editor.snapshot().image !== null;
}

/** A picture command runs after its decode: a refusal is the editor's own answer. */
function background(run: Promise<unknown>): void {
  void run.catch(() => undefined);
}

function act(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const length = texts[index]?.length ?? 0;
  const at = Math.floor(rand() * (length + 1));
  const roll = rand();
  if (roll < 0.22) {
    caretAt(person, index, at);
    s.insertPlainText(['a', 'bc', 'X '][Math.floor(rand() * 3)] ?? 'a');
  } else if (roll < 0.32) {
    caretAt(person, index, at);
    s.deleteBackward();
  } else if (roll < 0.38) {
    caretAt(person, index, at);
    s.splitParagraph();
  } else if (roll < 0.52) {
    caretAt(person, index, at);
    if (rand() < 0.2) {
      // Insert Text Box puts the caret in the new box; a click in the body takes it back
      person.editor.exec({ type: 'insertTextBox' });
      s.setActiveScope({ kind: 'body' });
      return;
    }
    const gif = rand() < 0.3;
    background(
      person.editor.executeImageCommand({
        type: 'insertImage',
        data: gif ? GIF : RED_PNG,
        mime: gif ? 'image/gif' : 'image/png',
        widthPoints: 12 + Math.floor(rand() * 24),
        heightPoints: 12,
      })
    );
  } else if (roll < 0.62) {
    if (selectPicture(person, rand)) person.editor.exec({ type: 'deleteImage' });
  } else if (roll < 0.8) {
    if (!selectPicture(person, rand)) return;
    const which = Math.floor(rand() * 5);
    if (which === 0)
      person.editor.exec({
        type: 'setImageProperties',
        widthEmu: 127_000 * (1 + Math.floor(rand() * 4)),
        heightEmu: 127_000,
      });
    else if (which === 1)
      person.editor.exec({ type: 'setImageWrapType', target: rand() < 0.5 ? 'square' : 'inline' });
    else if (which === 2) person.editor.exec({ type: 'transformImage', action: 'flipH' });
    else if (which === 3) person.editor.exec({ type: 'setImageProperties', alt: `alt ${at}` });
    else
      background(
        person.editor.executeImageCommand({ type: 'replaceImage', data: GIF, mime: 'image/gif' })
      );
  } else if (roll < 0.9) s.undo();
  else s.redo();
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
      imageDecodePort: decodePort(),
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
      const doc = person.handle.session.orderedDocument();
      const media = [...doc.resources.added.values()].filter((x) => x.kind === 'media');
      const missing = media.filter((x) => x.kind === 'media' && room.blob(x.sha256) === undefined);
      expect({
        label,
        who: person.name,
        unnamed,
        status: person.handle.session.status(),
        room: sameAsRoom(person, room),
        screen: screenFollowsList(person),
        missing: missing.length,
      }).toEqual({
        label,
        who: person.name,
        unnamed: [],
        status: 'ready',
        room: true,
        screen: true,
        missing: 0,
      });
    }
  }
  expect(room.room.order).toBeGreaterThan(steps);
}

/** More seeds for a local sweep: ORDERED_SWEEP=50 ORDERED_SWEEP_FROM=200. */
const SWEEP = Number(process.env.ORDERED_SWEEP ?? 0);
const FROM = Number(process.env.ORDERED_SWEEP_FROM ?? 100);

describe('random sessions with pictures', () => {
  for (let seed = FROM; seed < FROM + SWEEP; seed += 1)
    test(`sweep, seed ${seed}`, async () => {
      await session(seed, 30);
    });
  for (const seed of [1, 2, 3, 4])
    test(`the organiser and two parties, seed ${seed}`, async () => {
      await session(seed, 30);
    });
});
