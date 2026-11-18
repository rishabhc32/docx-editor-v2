// Headers, footers and sections in an ordered room (task 9.8), with real editors and one room.

import { afterEach, describe, expect, test } from 'bun:test';
import { openOrderedWordFile } from '../../client/open.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  join,
  listTexts,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import {
  bodySection,
  converged,
  FILE,
  HEADER,
  headers,
  inHeader,
  lifecycle,
  listTextsOf,
  room2,
  saved,
  TWO,
} from './sections-fixtures.ts';

afterEach(closeAll);

describe('typing in headers and footers', () => {
  test('the organiser types in the header: every copy shows it', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: FILE });
    inHeader(olive, 6);
    surface(olive).insertPlainText(' new');
    await settle();
    expect(olive.refusals).toEqual([]);
    for (const person of [ana, olive]) {
      expect(sameAsRoom(person, room)).toBe(true);
      expect(listTexts(person, HEADER)).toEqual(['Header new text']);
    }
  });

  test("a party's typing in the header is refused by the party rule", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE });
    inHeader(ana, 6);
    surface(ana).insertPlainText(' new');
    await settle();
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual(['ordered-refused:furniture']);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(listTexts(ana, HEADER)).toEqual(['Header text']);
  });
});

describe("the organiser's section edits", () => {
  test('page setup on the whole document: undo and redo, and the save', async () => {
    const { room, ana, olive } = await room2();
    expect(surface(olive).setSectionProperties({ orientation: 'landscape' })).toBe(true);
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive])
      expect(bodySection((await saved(person))['word/document.xml'] ?? '')).toContain(
        'w:orient="landscape"'
      );
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(bodySection((await saved(ana))['word/document.xml'] ?? '')).not.toContain('landscape');
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(bodySection((await saved(ana))['word/document.xml'] ?? '')).toContain('landscape');
  });

  test('a new footer: one part and one relationship on every copy, none after undo', async () => {
    const { room, ana, olive } = await room2();
    const made = surface(olive).applyHeaderFooterLifecycle?.({
      op: 'createHeaderFooter',
      sectionIndex: 0,
      kind: 'footer',
      variant: 'default',
    });
    expect(made?.ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    const footers = async (person: Person) =>
      Object.keys(await saved(person)).filter(
        (name) => /footer/.test(name) && !name.includes('_rels')
      );
    const [footer] = await footers(ana);
    expect(await footers(olive)).toEqual(await footers(ana));
    expect(footer).toBeDefined();
    const parts = await saved(ana);
    const target = (footer ?? '').replace(/^word\//, '');
    expect(parts['word/_rels/document.xml.rels']).toContain(`Target="${target}"`);
    expect(bodySection(parts['word/document.xml'] ?? '')).toContain('w:footerReference');
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(await footers(ana)).toEqual([]);
    expect((await saved(ana))['word/_rels/document.xml.rels']).not.toContain('footer');
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(await footers(ana)).toEqual([footer]);
  });

  test('the header removed: its part goes from the save, and undo brings it back', async () => {
    const { room, ana, olive } = await room2();
    const removed = surface(olive).applyHeaderFooterLifecycle?.({
      op: 'deleteHeaderFooter',
      sectionIndex: 0,
      kind: 'header',
      variant: 'default',
    });
    expect(removed?.ok).toBe(true);
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      const parts = await saved(person);
      expect(parts['word/header1.xml']).toBeUndefined();
      expect(parts['word/_rels/document.xml.rels']).not.toContain('header1.xml');
      expect(parts['[Content_Types].xml']).not.toContain('header1.xml');
    }
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect((await saved(ana))['word/header1.xml']).toContain('Header text');
  });

  test('a section break: two sections on every copy, one after undo, and the save reopens', async () => {
    const { room, ana, olive } = await room2();
    caretAt(olive, 0, 5);
    expect(surface(olive).insertSectionBreak('nextPage')).toBe(true);
    await settle();
    converged(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorTexts(person)).toEqual(['First', ' page text.', 'More text.']);
      expect(surface(person).session.headerFooterResolutionBySection().length).toBe(2);
    }
    const bytes = new Uint8Array(await ana.editor.save());
    const reopened = openOrderedWordFile(bytes, { file: 2 });
    expect(listTextsOf(reopened.document)).toEqual(['First', ' page text.', 'More text.']);
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(surface(ana).session.headerFooterResolutionBySection().length).toBe(1);
    expect(editorTexts(ana)).toEqual(['First page text.', 'More text.']);
  });
});

describe('links between sections', () => {
  test('unlink: the second section gets its own copy, which the organiser edits alone', async () => {
    const { room, ana, olive } = await room2(TWO);
    const slot = () =>
      surface(ana).session.headerFooterResolutionBySection()[1]?.headers.get('default');
    expect(slot()?.inherited).toBe(true);
    expect(
      lifecycle(olive, {
        op: 'unlinkFromPrevious',
        sectionIndex: 1,
        kind: 'header',
        variant: 'default',
      })?.ok
    ).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(slot()?.inherited).toBe(false);
    const copies = await headers(ana);
    expect(Object.values(copies)).toEqual(['Header text', 'Header text']);
    expect(await headers(olive)).toEqual(copies);
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    expect(Object.keys(await headers(ana))).toEqual(['word/header1.xml']);
    surface(olive).redo();
    await settle();
    converged(room, ana, olive);
    expect(Object.keys(await headers(ana))).toEqual(Object.keys(copies));
  });

  test('link to previous after an unlink: the copy goes, and the section shows the first header', async () => {
    const { room, ana, olive } = await room2(TWO);
    lifecycle(olive, {
      op: 'unlinkFromPrevious',
      sectionIndex: 1,
      kind: 'header',
      variant: 'default',
    });
    await settle();
    expect(
      lifecycle(olive, {
        op: 'linkToPrevious',
        sectionIndex: 1,
        kind: 'header',
        variant: 'default',
      })?.ok
    ).toBe(true);
    await settle();
    converged(room, ana, olive);
    expect(Object.keys(await headers(ana))).toEqual(['word/header1.xml']);
    expect(
      surface(ana).session.headerFooterResolutionBySection()[1]?.headers.get('default')?.inherited
    ).toBe(true);
  });

  test('a different first page and different odd and even pages: section and settings, with undo', async () => {
    const { room, ana, olive } = await room2();
    expect(
      lifecycle(olive, {
        op: 'setSectionFurnitureOptions',
        sectionIndex: 0,
        titlePage: true,
        evenAndOddHeaders: true,
      })?.ok
    ).toBe(true);
    await settle();
    converged(room, ana, olive);
    const parts = await saved(ana);
    expect(bodySection(parts['word/document.xml'] ?? '')).toContain('<w:titlePg/>');
    expect(parts['word/settings.xml']).toContain('evenAndOddHeaders');
    const section = surface(ana).session.headerFooterResolutionBySection()[0];
    expect([section?.titlePage, section?.evenAndOddHeaders]).toEqual([true, true]);
    surface(olive).undo();
    await settle();
    converged(room, ana, olive);
    const after = await saved(ana);
    expect(bodySection(after['word/document.xml'] ?? '')).not.toContain('titlePg');
    expect(after['word/settings.xml'] ?? '').not.toContain('evenAndOddHeaders');
  });
});

describe("a party's section edits", () => {
  test('a section break is a suggestion: the organiser accepts it, or rejects it and the paragraphs join', async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const review = reviewTestModule();
      const ana = await join(room, 'Ana', { wordFile: FILE, review });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: FILE,
        review,
      });
      caretAt(ana, 0, 5);
      expect(surface(ana).insertSectionBreak('nextPage')).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(surface(olive).session.headerFooterResolutionBySection().length).toBe(2);
      const xml = (await saved(olive))['word/document.xml'] ?? '';
      expect(xml).toMatch(/<w:rPr><w:ins [^>]*w:author="Ana"[^>]*\/><\/w:rPr><w:sectPr>/);
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      expect(surface(ana).session.headerFooterResolutionBySection().length).toBe(
        decision === 'accept' ? 2 : 1
      );
      expect(editorTexts(ana)).toEqual(
        decision === 'accept'
          ? ['First', ' page text.', 'More text.']
          : ['First page text.', 'More text.']
      );
      closeAll();
    }
  });

  test("a party's page setup is a suggestion, Word's w:sectPrChange: accepted, or rejected back to the file", async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
      const review = reviewTestModule();
      const ana = await join(room, 'Ana', { wordFile: FILE, review });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: FILE,
        review,
      });
      const before = bodySection((await saved(olive))['word/document.xml'] ?? '');
      expect(surface(ana).setSectionProperties({ orientation: 'landscape' })).toBe(true);
      await settle();
      converged(room, ana, olive);
      // Word's form: the section as it is now, and its earlier properties in full under Ana
      const pending = (await saved(olive))['word/document.xml'] ?? '';
      expect(pending).toMatch(/<w:pgSz [^>]*w:orient="landscape"/);
      expect(pending).toMatch(
        /<w:sectPrChange [^>]*w:author="Ana"[^>]*><w:sectPr><w:pgSz [^>]*w:orient="portrait"/
      );
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      // A section change, at the document's last paragraph mark, as Word lists it, on every copy
      const last = surface(olive).session.paragraphIds().at(-1) ?? '';
      for (const person of [ana, olive]) {
        const [own] = person.editor.getReviewItems().filter((item) => item.kind === 'revision');
        expect(
          own?.kind === 'revision' ? [own.item.formattingKind, own.item.ranges] : null
        ).toEqual([
          'sectPrChange',
          [
            {
              partName: '/word/document.xml',
              start: { paragraphId: last, offset: 'More text.'.length },
              end: { paragraphId: last, offset: 'More text.'.length },
            },
          ],
        ]);
      }
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      const after = bodySection((await saved(ana))['word/document.xml'] ?? '');
      expect(after).not.toContain('sectPrChange');
      if (decision === 'accept') expect(after).toContain('w:orient="landscape"');
      else expect(after).toBe(before);
      closeAll();
    }
  });

  test("a party's page setup over another person's pending one is refused: one record for each section", async () => {
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
    const firstOf = (person: typeof ana) => surface(person).session.paragraphIds()[0] ?? '';
    const secondOf = (person: typeof ana) => surface(person).session.paragraphIds()[1] ?? '';
    // Ben suggests landscape for the first section
    expect(
      surface(ben).setSectionProperties({
        orientation: 'landscape',
        anchorParagraphId: firstOf(ben),
      })
    ).toBe(true);
    await settle();
    converged(room, ...all);
    // Ana's page setup on the second section is a suggestion of her own
    expect(
      surface(ana).setSectionProperties({ marginLeftTwips: 2160, anchorParagraphId: secondOf(ana) })
    ).toBe(true);
    await settle();
    converged(room, ...all);
    // On the first section, or on every section, it would take Ben's record and his name
    for (const anchor of [firstOf(ana), undefined]) {
      surface(ana).setSectionProperties({
        marginTopTwips: 720,
        ...(anchor === undefined ? {} : { anchorParagraphId: anchor }),
      });
      await settle();
      expect(ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`)).toEqual([
        'ordered-refused:others_suggestions',
      ]);
    }
    converged(room, ...all);
    // The organiser's direct page setup there is a direct edit, as before
    expect(
      surface(olive).setSectionProperties({
        marginTopTwips: 720,
        anchorParagraphId: firstOf(olive),
      })
    ).toBe(true);
    await settle();
    converged(room, ...all);
    // Each suggestion keeps its author: Ben's on the first section, Ana's on the second
    const xml = (await saved(olive))['word/document.xml'] ?? '';
    const authors = [...xml.matchAll(/<w:sectPrChange [^>]*w:author="([^"]*)"/g)].map((m) => m[1]);
    expect(authors).toEqual(['Ben', 'Ana']);
    expect(xml).toContain('w:top="720"');
  });

  test("a party's section break that would copy another person's pending section record is refused", async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const review = reviewTestModule();
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    const cy = await join(room, 'Cy', { wordFile: FILE, review });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: FILE,
      review,
    });
    const all = [ana, cy, olive];
    // Cy suggests a page setup for the whole document: one record on its one section
    expect(surface(cy).setSectionProperties({ marginTopTwips: 2880 })).toBe(true);
    await settle();
    converged(room, ...all);
    // Ana's break would copy Cy's record into her new section, inside her own suggestion:
    // rejecting her break would take Cy's record back with it
    caretAt(ana, 0, 5);
    surface(ana).insertSectionBreak('nextPage');
    await settle();
    expect(ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-refused:nested',
    ]);
    converged(room, ...all);
    expect(surface(ana).session.headerFooterResolutionBySection().length).toBe(1);
    // The organiser's break there is a direct edit: both sections keep Cy's suggestion
    caretAt(olive, 0, 5);
    expect(surface(olive).insertSectionBreak('nextPage')).toBe(true);
    await settle();
    converged(room, ...all);
    const xml = (await saved(cy))['word/document.xml'] ?? '';
    const authors = [...xml.matchAll(/<w:sectPrChange [^>]*w:author="([^"]*)"/g)].map((m) => m[1]);
    expect(authors).toEqual(['Cy', 'Cy']);
  });

  test("a party's header and odd and even pages are refused by the host's party rules", async () => {
    const { room, ana, olive } = await room2();
    const reasons: string[] = [];
    const attempt = async (run: () => void) => {
      run();
      await settle();
      reasons.push(...ana.refusals.splice(0).map((r) => `${r.code}:${r.reason}`));
    };
    await attempt(() =>
      lifecycle(ana, {
        op: 'createHeaderFooter',
        sectionIndex: 0,
        kind: 'footer',
        variant: 'default',
      })
    );
    await attempt(() =>
      lifecycle(ana, {
        op: 'deleteHeaderFooter',
        sectionIndex: 0,
        kind: 'header',
        variant: 'default',
      })
    );
    // A setting has no tracked form: the editor writes it untracked in suggesting mode too
    await attempt(() =>
      lifecycle(ana, { op: 'setSectionFurnitureOptions', evenAndOddHeaders: true })
    );
    expect(reasons).toEqual([
      'ordered-refused:furniture',
      'ordered-refused:furniture',
      'ordered-refused:package',
    ]);
    converged(room, olive);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(Object.keys(await headers(ana))).toEqual(['word/header1.xml']);
    expect((await saved(ana))['word/settings.xml'] ?? '').not.toContain('evenAndOddHeaders');
  });
});

describe("a file's own section change record", () => {
  const RECORDED = docx(
    p(t('Text.')) +
      '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>' +
      '<w:sectPrChange w:id="7" w:author="Cara" w:date="2026-01-01T00:00:00Z">' +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>'
  );

  test('the organiser accepts or rejects it through the room', async () => {
    for (const decision of ['accept', 'reject'] as const) {
      const room = localRoom(RECORDED, { policy: suggestionsOnlyPolicy() });
      const review = reviewTestModule();
      const ana = await join(room, 'Ana', { wordFile: RECORDED, review });
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: RECORDED,
        review,
      });
      expect(bodySection((await saved(olive))['word/document.xml'] ?? '')).toContain(
        'w:sectPrChange'
      );
      const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
      if (!card) throw new Error('no card');
      const done =
        decision === 'accept'
          ? olive.editor.acceptReviewItem(card.key)
          : olive.editor.rejectReviewItem(card.key);
      expect(done.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      const section = bodySection((await saved(ana))['word/document.xml'] ?? '');
      expect(section).not.toContain('sectPrChange');
      // The engine's reject lays the recorded attributes over the current ones, as Word reads
      // the record (relative); the page is the recorded size
      expect(section).toContain(decision === 'accept' ? 'w:w="15840"' : 'w:w="12240"');
      closeAll();
    }
  });
});
