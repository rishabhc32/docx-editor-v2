// Section edits that meet other people's edits in an ordered room (task 9.8, 9.11), with real
// editors and one room. Each case runs with the two edits arriving in both orders: every copy
// ends as the room, and the result is the one Word gives when the room's order is the order the
// edits were made in.

import { afterEach, describe, expect, test } from 'bun:test';
import { collectSectionPropertyNodes } from '../../../../store/package/hf-references.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  listTexts,
  localRoom,
  ORGANISER,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged, FILE, HEADER, headers, inHeader, lifecycle, TWO } from './sections-fixtures.ts';

afterEach(closeAll);

/** A party and two organisers (one person on two devices, or two organisers) in one room. */
async function room3(file: Uint8Array) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const review = reviewTestModule();
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: file,
    review,
  });
  const oscar = await join(room, 'Oscar', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: file,
    review,
  });
  return { room, ana, olive, oscar, all: [ana, olive, oscar] };
}

/** Each section of a person's save, in order: its orientation, and whether it holds a record. */
async function sectionsOf(person: Person): Promise<string[]> {
  const read = readOoxmlPackage(new Uint8Array(await person.editor.save()), {});
  if (!read.ok) throw new Error(read.reason);
  const main = read.package.parts.get(read.package.mainDocumentPart);
  if (!main) throw new Error('no main part');
  return collectSectionPropertyNodes(main.root).map((section) => {
    const children = section?.children ?? [];
    const size = children.find((c) => c.kind !== 'textValue' && c.localName === 'pgSz');
    const orient =
      size && size.kind !== 'textValue'
        ? (size.attributes.find((a) => a.localName === 'orient')?.value ?? 'portrait')
        : 'portrait';
    const record = children.some((c) => c.kind !== 'textValue' && c.localName === 'sectPrChange');
    return record ? `${orient} (pending)` : orient;
  });
}

/** Run two edits as concurrent: both made before any frame reaches the room. */
async function together(first: () => void, second: () => void): Promise<void> {
  first();
  second();
  await settle();
}

const ORDERS = ['first', 'second'] as const;

describe('a whole-document page setup and a section break made at the same time', () => {
  for (const order of ORDERS)
    test(`the new section takes the page setup too (the break ${order})`, async () => {
      const { room, ana, olive, all } = await room3(FILE);
      const breakIt = () => {
        caretAt(ana, 0, 5);
        expect(surface(ana).insertSectionBreak('nextPage')).toBe(true);
      };
      const landscape = () =>
        expect(surface(olive).setSectionProperties({ orientation: 'landscape' })).toBe(true);
      await (order === 'first' ? together(breakIt, landscape) : together(landscape, breakIt));
      converged(room, ...all);
      for (const person of all)
        expect(await sectionsOf(person)).toEqual(['landscape', 'landscape']);
    });
});

describe("a party's page setup and the organiser's section break made at the same time", () => {
  for (const order of ORDERS)
    for (const decision of ['accept', 'reject'] as const)
      test(`both sections show the suggestion until it is ${decision}ed (the break ${order})`, async () => {
        const { room, ana, olive, all } = await room3(FILE);
        const breakIt = () => {
          caretAt(olive, 0, 5);
          expect(surface(olive).insertSectionBreak('nextPage')).toBe(true);
        };
        const landscape = () =>
          expect(surface(ana).setSectionProperties({ orientation: 'landscape' })).toBe(true);
        await (order === 'first' ? together(breakIt, landscape) : together(landscape, breakIt));
        converged(room, ...all);
        // Made after the break, it goes to both sections; made before, the break's copy keeps it
        for (const person of all)
          expect(await sectionsOf(person)).toEqual(['landscape (pending)', 'landscape (pending)']);
        const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
        if (!card) throw new Error('no card');
        const done =
          decision === 'accept'
            ? olive.editor.acceptReviewItem(card.key)
            : olive.editor.rejectReviewItem(card.key);
        expect(done.ok).toBe(true);
        await settle();
        converged(room, ...all);
        const page = decision === 'accept' ? 'landscape' : 'portrait';
        for (const person of all) expect(await sectionsOf(person)).toEqual([page, page]);
      });
});

describe('a section break deleted while its section changes', () => {
  for (const order of ORDERS)
    test(`two organisers: the join and the page setup (the join ${order})`, async () => {
      const { room, olive, oscar, all } = await room3(TWO);
      const joinThem = () => {
        caretAt(olive, 1, 0);
        surface(olive).deleteBackward();
      };
      const [first = ''] = surface(oscar).session.paragraphIds();
      const landscape = () =>
        expect(
          surface(oscar).setSectionProperties({
            orientation: 'landscape',
            anchorParagraphId: first,
          })
        ).toBe(true);
      await (order === 'first' ? together(joinThem, landscape) : together(landscape, joinThem));
      converged(room, ...all);
      expect(editorTexts(olive)).toEqual(['Section one.Section two.']);
      // One after the other in the room's order, as in Word: a page setup made after the join
      // goes to the section the joined paragraph is in; one made before goes with the break
      expect(await sectionsOf(olive)).toEqual([order === 'first' ? 'landscape' : 'portrait']);
      for (const person of all) expect(await sectionsOf(person)).toEqual(await sectionsOf(olive));
    });

  test("a party's deletion of the break is a suggestion; the organiser's page setup applies to its section until it is accepted", async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const { room, ana, olive, all } = await room3(TWO);
      const [first = ''] = surface(olive).session.paragraphIds();
      await together(
        () => {
          caretAt(ana, 1, 0);
          surface(ana).deleteBackward();
        },
        () =>
          expect(
            surface(olive).setSectionProperties({
              orientation: 'landscape',
              anchorParagraphId: first,
            })
          ).toBe(true)
      );
      converged(room, ...all);
      expect(await sectionsOf(ana)).toEqual(['landscape', 'portrait']);
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ...all);
      // Accepted, the break goes and its text takes the next section's properties, as in Word
      expect(await sectionsOf(ana)).toEqual(
        decision === 'accept' ? ['portrait'] : ['landscape', 'portrait']
      );
      closeAll();
    }
  });
});

describe('a header unlinked while someone types in it', () => {
  for (const order of ORDERS)
    test(`the copy holds the typing when the room has it first (the typing called ${order})`, async () => {
      const { room, olive, oscar, all } = await room3(TWO);
      const type = () => {
        inHeader(oscar, 6);
        surface(oscar).insertPlainText(' new');
      };
      const unlink = () =>
        expect(
          lifecycle(olive, {
            op: 'unlinkFromPrevious',
            sectionIndex: 1,
            kind: 'header',
            variant: 'default',
          })?.ok
        ).toBe(true);
      await (order === 'first' ? together(type, unlink) : together(unlink, type));
      converged(room, ...all);
      expect(listTexts(olive, HEADER)).toEqual(['Header new text']);
      // The room's order decides, as one edit after the other would in Word
      const kinds = room.room.joinSnapshot(0).map((change) => change.kind);
      const typedFirst = kinds.indexOf('insert') < kinds.indexOf('header-footer');
      const texts = Object.values(await headers(olive)).sort();
      expect(texts).toEqual(
        typedFirst ? ['Header new text', 'Header new text'] : ['Header new text', 'Header text']
      );
      for (const person of all) expect(await headers(person)).toEqual(await headers(olive));
    });
});

describe('Enter in a paragraph that ends a section the organiser changed', () => {
  // Found by the random sessions: Enter copied the section's records to the first half, which
  // then ended a section too, and a section break there was refused
  test('the break stays with the second half, and a new break there is a suggestion', async () => {
    const { room, ana, olive, all } = await room3(TWO);
    const [first = ''] = surface(olive).session.paragraphIds();
    expect(
      surface(olive).setSectionProperties({ orientation: 'landscape', anchorParagraphId: first })
    ).toBe(true);
    await settle();
    caretAt(ana, 0, 7);
    surface(ana).splitParagraph();
    await settle();
    converged(room, ...all);
    expect(editorTexts(ana)).toEqual(['Section', ' one.', 'Section two.']);
    expect(await sectionsOf(ana)).toEqual(['landscape', 'portrait']);
    caretAt(ana, 1, 2);
    expect(surface(ana).insertSectionBreak('nextPage')).toBe(true);
    await settle();
    converged(room, ...all);
    for (const person of all)
      expect(await sectionsOf(person)).toEqual(['landscape', 'landscape', 'portrait']);
  });
});

describe("a party's section break taken back", () => {
  test('undo takes it back on every copy, and redo makes it again', async () => {
    const { room, ana, all } = await room3(FILE);
    caretAt(ana, 0, 5);
    expect(surface(ana).insertSectionBreak('nextPage')).toBe(true);
    await settle();
    converged(room, ...all);
    expect(await sectionsOf(ana)).toEqual(['portrait', 'portrait']);
    surface(ana).undo();
    await settle();
    converged(room, ...all);
    for (const person of all) expect(await sectionsOf(person)).toEqual(['portrait']);
    expect(editorTexts(ana)).toEqual(['First page text.', 'More text.']);
    surface(ana).redo();
    await settle();
    converged(room, ...all);
    expect(editorTexts(ana)).toEqual(['First', ' page text.', 'More text.']);
  });
});
