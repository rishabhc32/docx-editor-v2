// Random sessions with headers, footers and sections (task 9.8, 9.11), with real editors and the
// suggestions-only policy: parties type, delete, press Enter and Backspace, insert section breaks,
// change page setup, undo and redo, and try the organiser's header, footer and furniture option
// edits, which the host's party rules refuse (and a break that retypes the next section); the
// organiser edits text and changes page setup, furniture options, headers, footers, links between
// sections and the watermark, directly or as suggestions, and accepts and rejects suggestions, all
// while the room's frames wait and arrive in random amounts. After each round, when every frame has
// arrived: every copy equals the room, every screen shows its list in every story, and every
// refusal names a real conflict or a party rule.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { seededRandom } from '../random.ts';
import { docx, p, R, t, WML_CT } from '../word-fixtures.ts';
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
import { reviewTestModule } from './review-module.ts';
import { lifecycle, showsTheList } from './sections-fixtures.ts';

afterEach(closeAll);

const HEADER_PART =
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r><w:t>Header text</w:t></w:r></w:p></w:hdr>';
const FOOTER_PART =
  '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:p><w:r><w:t>Footer text</w:t></w:r></w:p></w:ftr>';

/** Two sections: the first names a header, the second a footer and takes the header too. */
const FILE = docx(
  p(t('The fee is 100 USD per month.')) +
    p(
      t('Payment is due in 30 days.'),
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader1"/>' +
        '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>'
    ) +
    p(t('Signed by both parties.')) +
    '<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  {
    parts: { 'word/header1.xml': HEADER_PART, 'word/footer1.xml': FOOTER_PART },
    relationships:
      `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdFooter1" Type="${R}/footer" Target="footer1.xml"/>`,
    contentTypes:
      `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>` +
      `<Override PartName="/word/footer1.xml" ContentType="${WML_CT}.footer+xml"/>`,
  }
);

/**
 * Reasons a refusal may have here: a real conflict between people's edits, or a party rule (a
 * party's header, footer or furniture options, or a break that retypes the next section; an
 * edit inside another person's suggestion).
 */
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
  // A decision on a card that changed before it arrived (the seen check, D6)
  'stale-decision',
  'not-pending',
  'not-a-section',
  'furniture',
  'package',
  'nested',
  'others_suggestions',
]);

/**
 * The engine's refusals of a header or footer edit whose premise another person's earlier edit
 * changed (both unlinked the same header, one deleted what the other links).
 */
const LIFECYCLE =
  /^(createHeaderFooter|deleteHeaderFooter|linkToPrevious|unlinkFromPrevious): (slot-occupied|already-declared|not-declared|nothing-to-unlink|first-section)$/;

const named = (report: Person['refusals'][number]): boolean =>
  NAMED.has(report.reason) ||
  (report.reason === 'engine-refused' && LIFECYCLE.test(report.detail ?? '')) ||
  // A change that adds to the person's own suggestion, which the organiser decided first
  (report.reason === 'text' && / is decided$/.test(report.detail ?? ''));

/** One text edit at a random place in the body. */
function editText(person: Person, rand: () => number): void {
  const s = surface(person);
  const texts = editorTexts(person);
  const index = Math.floor(rand() * texts.length);
  const length = texts[index]?.length ?? 0;
  const at = Math.floor(rand() * (length + 1));
  const roll = rand();
  caretAt(person, index, at);
  if (roll < 0.45) s.insertPlainText(['a', 'bc', 'X '][Math.floor(rand() * 3)] ?? 'a');
  else if (roll < 0.65) s.deleteBackward();
  else if (roll < 0.8) s.splitParagraph();
  else s.insertSectionBreak(rand() < 0.5 ? 'nextPage' : 'continuous');
}

const pick = <T>(rand: () => number, items: readonly T[]): T | undefined =>
  items[Math.floor(rand() * items.length)];

/** One section edit: page setup, furniture options, a header or footer, a link, the watermark. */
function editSections(person: Person, rand: () => number): void {
  const s = surface(person);
  const sections = s.session.headerFooterResolutionBySection().length;
  const sectionIndex = Math.floor(rand() * sections);
  const ids = s.session.paragraphIds();
  const roll = rand();
  if (roll < 0.2) s.setSectionProperties({ orientation: rand() < 0.5 ? 'landscape' : 'portrait' });
  else if (roll < 0.35)
    s.setSectionProperties({
      marginLeftTwips: 720 + 360 * Math.floor(rand() * 4),
      ...(pick(rand, ids) === undefined ? {} : { anchorParagraphId: pick(rand, ids) }),
    });
  else if (roll < 0.45)
    lifecycle(person, { op: 'setSectionFurnitureOptions', sectionIndex, titlePage: rand() < 0.5 });
  else if (roll < 0.6)
    lifecycle(person, {
      op: rand() < 0.5 ? 'createHeaderFooter' : 'deleteHeaderFooter',
      sectionIndex,
      kind: rand() < 0.5 ? 'header' : 'footer',
      variant: 'default',
    });
  else if (roll < 0.8)
    lifecycle(person, {
      op: rand() < 0.5 ? 'unlinkFromPrevious' : 'linkToPrevious',
      sectionIndex,
      kind: rand() < 0.5 ? 'header' : 'footer',
      variant: 'default',
    });
  else
    person.editor.exec({
      type: 'setWatermark',
      watermark: rand() < 0.7 ? { text: 'DRAFT' } : null,
    });
}

/** The organisers: Olga edits directly, Oscar suggests. */
const ORGANISERS: Readonly<Record<string, 'edit' | 'suggesting'>> = {
  Olga: 'edit',
  Oscar: 'suggesting',
};

/** An organiser accepts or rejects a random card. */
function decideOne(person: Person, rand: () => number): void {
  const cards = person.editor.getReviewItems().filter((item) => item.kind === 'revision');
  const card = pick(rand, cards);
  if (!card) return;
  if (rand() < 0.5) person.editor.acceptReviewItem(card.key);
  else person.editor.rejectReviewItem(card.key);
}

function act(person: Person, rand: () => number): void {
  const roll = rand();
  if (roll < 0.55) editText(person, rand);
  else if (roll < 0.8) editSections(person, rand);
  else if (roll < 0.9) surface(person).undo();
  else if (roll < 0.95 || ORGANISERS[person.name] === undefined) surface(person).redo();
  else decideOne(person, rand);
}

async function session(seed: number, names: readonly string[], steps: number): Promise<void> {
  const rand = seededRandom(seed);
  const queue: (() => void)[] = [];
  const room = localRoom(FILE, {
    deliver: (run) => queue.push(run),
    policy: suggestionsOnlyPolicy(),
  });
  const people: Person[] = [];
  for (const name of names) {
    const mode = ORGANISERS[name];
    const joining = join(room, name, {
      wordFile: FILE,
      ...(mode === undefined ? {} : { mode, login: ORGANISER, review: reviewTestModule() }),
    });
    await drain(queue);
    people.push(await joining);
  }
  for (let round = 0; round < 3; round += 1) {
    for (let step = 0; step < steps; step += 1) {
      const person = pick(rand, people);
      if (person) act(person, rand);
      await Promise.resolve();
      await Promise.resolve();
      // Some frames arrive now, the rest wait: people edit on copies that are behind
      for (let k = Math.floor(rand() * 4); k > 0; k -= 1) queue.shift()?.();
      if (rand() < 0.2) await new Promise((resolve) => setTimeout(resolve, 0));
    }
    await drain(queue);
    const label = `seed ${seed} round ${round}`;
    for (const person of people) {
      const who = { label, who: person.name };
      const unnamed = person.refusals.filter((r) => !named(r));
      expect({ ...who, unnamed }).toEqual({ ...who, unnamed: [] });
      expect({ ...who, status: person.handle.session.status() }).toEqual({
        ...who,
        status: 'ready',
      });
      expect({ ...who, room: sameAsRoom(person, room) }).toEqual({ ...who, room: true });
      expect({ ...who, screen: screenFollowsList(person) }).toEqual({ ...who, screen: true });
      showsTheList(person);
    }
  }
  expect(room.room.order).toBeGreaterThan(steps);
}

/** More seeds for a local sweep: ORDERED_SWEEP=50 ORDERED_SWEEP_FROM=200. */
const SWEEP = Number(process.env.ORDERED_SWEEP ?? 0);
const FROM = Number(process.env.ORDERED_SWEEP_FROM ?? 100);

describe('random sessions with headers, footers and sections', () => {
  for (let seed = FROM; seed < FROM + SWEEP; seed += 1)
    test(`sweep, seed ${seed}`, async () => {
      await session(seed, [seed % 2 === 0 ? 'Olga' : 'Oscar', 'Ana', 'Ben'], 30);
    });
  for (const seed of [1, 2, 3])
    test(`two parties and the organiser, seed ${seed}`, async () => {
      await session(seed, ['Ana', 'Olga', 'Ben'], 30);
    });
  test('the organiser alone with a party, seed 5', async () => {
    await session(5, ['Olga', 'Ana'], 40);
  });
  for (const seed of [6, 7])
    test(`the organiser suggesting with two parties, seed ${seed}`, async () => {
      await session(seed, ['Oscar', 'Ana', 'Ben'], 30);
    });
});
