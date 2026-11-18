// A different first page and the header and footer distances in an ordered room (task 9.8).
// They are section formatting: with tracking on, Word records them in the section's
// `w:sectPrChange`, as it records page setup (Word for Mac, 2026-10-30,
// `/tmp/word-evidence-sections/~$hdrdist-tracked.docx` and `~$titlepg-tracked.docx`). So a
// party's is a suggestion that the organiser accepts or rejects, and the organiser's is direct,
// or a suggestion in suggesting mode.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { bodySection, converged, FILE, lifecycle, saved, TWO } from './sections-fixtures.ts';

afterEach(closeAll);

/** Each option, and what the section states once it is accepted. */
const OPTIONS = {
  'a different first page': {
    options: { titlePage: true },
    shows: '<w:titlePg/>',
    was: '<w:titlePg w:val="0"/>',
  },
  'the header distance': {
    options: { headerDistanceTwips: 1080 },
    shows: 'w:header="1080"',
    was: 'w:header="720"',
  },
  'the footer distance': {
    options: { footerDistanceTwips: 540 },
    shows: 'w:footer="540"',
    was: 'w:footer="720"',
  },
} as const;

/** A party, Ana, and the organiser, Olive, who decides, in one room on `file`. */
async function reviewRoom(file = FILE, organiserMode: 'edit' | 'suggesting' = 'edit') {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    mode: organiserMode,
    login: ORGANISER,
    wordFile: file,
    review,
  });
  return { room, ana, olive };
}

/** The body's section of a save that holds no change record. */
const sectionOf = async (person: Person): Promise<string> =>
  bodySection((await saved(person))['word/document.xml'] ?? '');

/** The main part of a save. */
const documentOf = async (person: Person): Promise<string> =>
  (await saved(person))['word/document.xml'] ?? '';

/** Decide the one revision card, as the organiser does. */
function decide(olive: Person, decision: 'accept' | 'reject'): void {
  const cards = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
  expect(cards).toHaveLength(1);
  const [card] = cards;
  if (!card) throw new Error('no card');
  const done =
    decision === 'accept'
      ? olive.editor.acceptReviewItem(card.key)
      : olive.editor.rejectReviewItem(card.key);
  expect(done.ok).toBe(true);
}

describe("a party's furniture options are suggestions", () => {
  for (const [name, { options, shows, was }] of Object.entries(OPTIONS))
    for (const outcome of ['accept', 'reject', 'undo'] as const)
      test(`${name}: Word's w:sectPrChange, then ${outcome}`, async () => {
        const { room, ana, olive } = await reviewRoom();
        const before = await sectionOf(olive);
        expect(
          lifecycle(ana, { op: 'setSectionFurnitureOptions', sectionIndex: 0, ...options })?.ok
        ).toBe(true);
        await settle();
        converged(room, ana, olive);
        // The section as it is now, and under Ana the section before it
        const pending = await documentOf(olive);
        const record = /<w:sectPrChange [^>]*w:author="Ana"[^>]*><w:sectPr>(.*?)<\/w:sectPr>/.exec(
          pending
        );
        expect(record?.[1]).toBeDefined();
        expect(record?.[1]).not.toContain(shows);
        expect(record?.[1]).toContain(was);
        expect(pending.slice(0, record?.index)).toContain(shows);
        if (outcome === 'undo') surface(ana).undo();
        else decide(olive, outcome);
        await settle();
        converged(room, ana, olive);
        const after = await sectionOf(ana);
        expect(after).not.toContain('sectPrChange');
        if (outcome === 'accept') expect(after).toContain(shows);
        // The host's earlier check: rejecting a party's suggestion gives back the file
        else expect(after).toBe(before);
      });

  test("on a section that holds another person's pending record it is refused", async () => {
    const room = localRoom(TWO, { policy: suggestionsOnlyPolicy() });
    const review = reviewTestModule();
    const ana = await join(room, 'Ana', { wordFile: TWO, review });
    const ben = await join(room, 'Ben', { wordFile: TWO, review });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: TWO,
      review,
    });
    const all = [ana, ben, olive];
    // Ben suggests a different first page for the first section
    expect(
      lifecycle(ben, { op: 'setSectionFurnitureOptions', sectionIndex: 0, titlePage: true })?.ok
    ).toBe(true);
    await settle();
    converged(room, ...all);
    // Ana's header distance there would take Ben's record and his name
    lifecycle(ana, {
      op: 'setSectionFurnitureOptions',
      sectionIndex: 0,
      headerDistanceTwips: 1080,
    });
    await settle();
    expect(ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-refused:others_suggestions',
    ]);
    // Her page setup there too: one record for each section
    const [first] = surface(ana).session.paragraphIds();
    surface(ana).setSectionProperties({
      marginTopTwips: 720,
      ...(first === undefined ? {} : { anchorParagraphId: first }),
    });
    await settle();
    expect(ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-refused:others_suggestions',
    ]);
    // On the second section it is a suggestion of her own
    expect(
      lifecycle(ana, {
        op: 'setSectionFurnitureOptions',
        sectionIndex: 1,
        footerDistanceTwips: 540,
      })?.ok
    ).toBe(true);
    await settle();
    converged(room, ...all);
    const xml = (await saved(olive))['word/document.xml'] ?? '';
    const authors = [...xml.matchAll(/<w:sectPrChange [^>]*w:author="([^"]*)"/g)].map((m) => m[1]);
    expect(authors).toEqual(['Ben', 'Ana']);
  });
});

describe("the organiser's furniture options", () => {
  test('in suggesting mode a suggestion the organiser decides; in editing mode direct', async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const { room, ana, olive } = await reviewRoom(FILE, 'suggesting');
      const before = await sectionOf(olive);
      expect(
        lifecycle(olive, { op: 'setSectionFurnitureOptions', sectionIndex: 0, titlePage: true })?.ok
      ).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(await documentOf(ana)).toMatch(/<w:sectPrChange [^>]*w:author="Olive"/);
      surface(olive).setEditingMode('edit');
      decide(olive, decision);
      await settle();
      converged(room, ana, olive);
      const after = await sectionOf(ana);
      expect(after).not.toContain('sectPrChange');
      if (decision === 'accept') expect(after).toContain('<w:titlePg/>');
      else expect(after).toBe(before);
      // Direct now: no record, and undo takes it back
      expect(
        lifecycle(olive, {
          op: 'setSectionFurnitureOptions',
          sectionIndex: 0,
          headerDistanceTwips: 1080,
        })?.ok
      ).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(await sectionOf(ana)).not.toContain('sectPrChange');
      expect(await sectionOf(ana)).toContain('w:header="1080"');
      surface(olive).undo();
      await settle();
      converged(room, ana, olive);
      expect(await sectionOf(ana)).not.toContain('w:header="1080"');
      closeAll();
    }
  });
});
