// Bookmarks and a table of contents in an ordered room (group 9.5), with real editors and one
// room. The organiser adds, moves and removes bookmarks, inserts a table of contents and
// refreshes it, as the editor alone does; the organiser's undo and redo are exact; a party's
// bookmark or refresh is refused by the host's party rule; edits at the same time converge; and a
// saved file reopens in a new room, where a refresh still works.

import { afterEach, describe, expect, test } from 'bun:test';
import { writeOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import type { OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, R, t, W, WML_CT } from '../word-fixtures.ts';
import {
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Editor,
} from './editor-harness.ts';
import {
  act,
  alone,
  converged,
  roomFile,
  sameAsAlone,
  select,
  storyAsRead,
  type Act,
} from './links-harness.ts';

afterEach(closeAll);

const start = (id: number, name: string) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/>`;
const end = (id: number) => `<w:bookmarkEnd w:id="${id}"/>`;

/** Three bookmarks: one in a paragraph, one across two, and a copy of the first's name. */
const MARKED = docx(
  p(t('The ') + start(1, 'Fee') + t('fee') + end(1) + t(' is 100 USD.')) +
    p(t('Due in ') + start(2, 'Days') + t('30 days.')) +
    p(t('Signed') + end(2) + t(' here.')) +
    p(start(3, 'fee') + t('Copy') + end(3))
);

const HEADER_PART = '/word/header1.xml';

/** A bookmark in a header the file has, and the body's paragraphs. */
const MARKED_HEADER = docx(
  p(t('The fee is 100 USD.')) +
    p(t('Due in 30 days.')) +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/></w:sectPr>',
  {
    parts: {
      'word/header1.xml': `<w:hdr xmlns:w="${W}" xmlns:r="${R}">${p(start(7, 'Head') + t('Acme') + end(7))}</w:hdr>`,
    },
    contentTypes: `<Override PartName="/word/header1.xml" ContentType="${WML_CT}.header+xml"/>`,
    relationships: `<Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/>`,
  }
);

const add =
  (name: string, index: number, from: number, to = from): Act =>
  (e) => (select(e, index, from, to), e.exec({ type: 'insertBookmark', name }));
const remove =
  (name: string): Act =>
  (e) =>
    e.exec({ type: 'deleteBookmark', name });
const keepFirst =
  (name: string): Act =>
  (e) =>
    e.exec({ type: 'keepFirstBookmark', name });

/** Select from one body paragraph's offset to another's. */
function selectAcross(editor: Editor, from: [number, number], to: [number, number]): void {
  const s = editor.surface;
  if (!s) throw new Error('no surface');
  const ids = s.session.paragraphIds();
  s.setSelection({
    anchor: { paragraphId: ids[from[0]] ?? '', offset: from[1] },
    head: { paragraphId: ids[to[0]] ?? '', offset: to[1] },
  });
}

describe('bookmarks (direct edits only: Word has no revision for a bookmark)', () => {
  test('the organiser adds a range and a point bookmark, as the editor alone does', async () => {
    await sameAsAlone(MARKED, 'edit', [add('Terms', 0, 4, 7)]);
    await sameAsAlone(MARKED, 'edit', [add('Here', 2, 7), add('Start', 1, 0)]);
    await sameAsAlone(MARKED, 'edit', [add('Whole', 1, 0, 15)]);
  });

  test('a name in use moves; a removal takes every copy, or keeps the first', async () => {
    await sameAsAlone(MARKED, 'edit', [add('DAYS', 0, 0, 3)]);
    await sameAsAlone(MARKED, 'edit', [remove('Fee')]);
    await sameAsAlone(MARKED, 'edit', [keepFirst('Fee')]);
    await sameAsAlone(MARKED, 'edit', [remove('Days')]);
  });

  test("a name the header uses moves to the body, out of the header's story too", async () => {
    const step = add('Head', 0, 4, 7);
    const { room } = await sameAsAlone(MARKED_HEADER, 'edit', [step]);
    const editorAlone = await alone(MARKED_HEADER, 'Ana', 'edit', step);
    const saved = roomFile(room.room.document, MARKED_HEADER);
    expect(storyAsRead(saved, HEADER_PART)).toBe(storyAsRead(editorAlone, HEADER_PART));
    expect(storyAsRead(saved, HEADER_PART)).not.toContain('bookmarkStart');
  });

  test('a deletion across paragraphs takes the bookmarks it encloses, as the editor alone', async () => {
    await sameAsAlone(MARKED, 'edit', [
      (e) => (selectAcross(e, [1, 3], [3, 4]), e.surface?.deleteBackward(), { ok: true }),
    ]);
  });

  test("the organiser's undo and redo of bookmarks are exact", async () => {
    const room = localRoom(MARKED);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: MARKED });
    const story = () => storyAsRead(roomFile(room.room.document, MARKED));
    const original = story();
    await act(olive, add('Fee', 1, 0, 3));
    const moved = story();
    await act(olive, remove('Days'));
    const removed = story();
    await act(olive, (e) => e.surface?.undo());
    expect(story()).toBe(moved);
    await act(olive, (e) => e.surface?.undo());
    expect(story()).toBe(original);
    await act(olive, (e) => e.surface?.redo());
    await act(olive, (e) => e.surface?.redo());
    expect(story()).toBe(removed);
    converged(room, [olive]);
  });

  test("the party rules: a party's bookmark is refused as a package change", async () => {
    const room = localRoom(MARKED, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: MARKED });
    const original = storyAsRead(roomFile(room.room.document, MARKED));
    await act(ana, add('Terms', 0, 4, 7));
    await act(ana, remove('Fee'));
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual([
      'ordered-refused:package',
      'ordered-refused:package',
    ]);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(storyAsRead(roomFile(room.room.document, MARKED))).toBe(original);
  });

  test('the organiser removes a bookmark while a party types inside it', async () => {
    const room = localRoom(MARKED, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: MARKED });
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: MARKED });
    select(ana.editor, 0, 5);
    surface(ana).insertPlainText('X');
    olive.editor.exec({ type: 'deleteBookmark', name: 'Fee' });
    await settle();
    converged(room, [ana, olive]);
    const story = storyAsRead(roomFile(room.room.document, MARKED));
    expect(story).toContain('X');
    expect(story).not.toContain('w:name="Fee"');
  });
});

const heading = (text: string) => p(t(text), '<w:pStyle w:val="Heading1"/>');

/** Two headings with body text, and a paragraph to put a table of contents before. */
const HEADINGS = docx(
  p(t('Contents follow.')) +
    heading('Terms') +
    p(t('The fee is 100 USD.')) +
    heading('Payment') +
    p(t('Due in 30 days.'))
);

/** A table of contents with no rows yet, as a file may hold one before its first update. */
const EMPTY_TOC =
  '<w:sdt><w:sdtPr><w:alias w:val="Contents"/></w:sdtPr><w:sdtContent>' +
  p(
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-1" \\h </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
  ) +
  p('<w:r><w:fldChar w:fldCharType="end"/></w:r>') +
  '</w:sdtContent></w:sdt>';

/** Each editor mints its own `_Toc` names, by its person: the names are not the content. */
const tocAsRead = (pkg: OoxmlPackage) => storyAsRead(pkg).replace(/_Toc\d+/g, '_Toc');

const insertToc: Act = (e) => (select(e, 0, 0), e.exec({ type: 'insertToc' }));
const refresh =
  (mode: 'entire' | 'pageNumbers'): Act =>
  (e) =>
    e.exec({ type: 'refreshToc', mode });
/** A page break before the second heading moves it to page two. */
const breakBeforePayment: Act = (e) => (
  select(e, 4, 0),
  e.surface?.insertPageBreak(),
  { ok: true }
);

/** Steps by the organiser in a room and by the editor alone: the saved bodies agree. */
async function tocSameAsAlone(steps: readonly Act[]) {
  const room = localRoom(HEADINGS);
  const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: HEADINGS });
  const ana = await join(room, 'Ana', { wordFile: HEADINGS });
  for (const step of steps) expect(await act(olive, step)).toMatchObject({ ok: true });
  converged(room, [olive, ana]);
  const editorAlone = await alone(HEADINGS, 'Olive', 'edit', (editor) => {
    for (const step of steps) step(editor);
  });
  const saved = roomFile(room.room.document, HEADINGS);
  expect(tocAsRead(saved)).toBe(tocAsRead(editorAlone));
  return { room, olive, ana, saved };
}

/** A complex field's runs, with its result. */
const fieldRuns = (code: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  t(result) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** A contents row as Word writes it: a link to the heading, a tab, and a PAGEREF field. */
const wordRow = (text: string, bookmark: string, opening = '') =>
  p(
    opening +
      `<w:hyperlink w:anchor="${bookmark}">` +
      t(text) +
      '<w:r><w:tab/></w:r>' +
      fieldRuns(`PAGEREF ${bookmark} \\h`, '1') +
      '</w:hyperlink>',
    '<w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
  );

/**
 * A table of contents as Word writes it, before the two headings it lists. Both rows show page
 * 1, but the second heading is on page 2.
 */
const WORD_TOC = docx(
  '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    wordRow(
      'Terms',
      '_Toc1',
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    ) +
    wordRow('Payment', '_Toc2') +
    p('<w:r><w:fldChar w:fldCharType="end"/></w:r>') +
    '</w:sdtContent></w:sdt>' +
    p(start(1, '_Toc1') + t('Terms') + end(1), '<w:pStyle w:val="Heading1"/>') +
    p(t('The fee is 100 USD.')) +
    p('<w:r><w:br w:type="page"/></w:r>') +
    p(start(2, '_Toc2') + t('Payment') + end(2), '<w:pStyle w:val="Heading1"/>') +
    p(t('Due in 30 days.'))
);

describe('a table of contents (direct edits only: the engine has no tracked form)', () => {
  test('inserted by the organiser, it saves as the editor alone inserts it', async () => {
    const { saved } = await tocSameAsAlone([insertToc]);
    const story = tocAsRead(saved);
    expect(story).toContain('TOC \\o');
    expect(story).toContain('Payment');
  });

  test('refreshed whole and by page numbers, as the editor alone refreshes it', async () => {
    const whole = await tocSameAsAlone([
      insertToc,
      (e) => (select(e, 4, 0), e.surface?.insertPlainText('New '), { ok: true }),
      refresh('entire'),
    ]);
    expect(tocAsRead(whole.saved)).toContain('<w:t>New Terms</w:t><w:ptab');
    const { saved } = await tocSameAsAlone([insertToc, breakBeforePayment, refresh('pageNumbers')]);
    expect(tocAsRead(saved)).toContain('>2<');
  });

  test("the organiser's undo of a refresh shows the old rows, and undo of the insert none", async () => {
    const room = localRoom(HEADINGS);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: HEADINGS });
    const story = () => tocAsRead(roomFile(room.room.document, HEADINGS));
    /** Undo until the story is `target`, in at most `steps` presses. */
    const undoTo = async (target: string, steps: number) => {
      for (let k = 0; k < steps && story() !== target; k += 1)
        await act(olive, (e) => e.surface?.undo());
      expect(story()).toBe(target);
    };
    const original = story();
    await act(olive, insertToc);
    const inserted = story();
    await act(olive, (e) => (select(e, 4, 0), e.surface?.insertPlainText('New ')));
    const typed = story();
    await act(olive, refresh('entire'));
    expect(story()).not.toBe(typed);
    // Each command is one step with its page pass, as Word's Update Table is
    await undoTo(typed, 1);
    await undoTo(inserted, 1);
    await undoTo(original, 1);
    converged(room, [olive]);
  });

  test('a refresh whose rows push a heading to the next page is one undo step', async () => {
    // The nine rows push the last heading from page 2 to page 3 after they were planned with
    // page 2, so the page pass writes: the rows and the pass are one step, as in the editor
    // alone (`toc-refresh.test.ts`) and in Word.
    const lines = Array.from({ length: 34 }, (_, k) => p(t(`Line ${k}`))).join('');
    const file = docx(EMPTY_TOC + [...'ABCDEFGH'].map(heading).join('') + lines + heading('Last'));
    const room = localRoom(file);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
    const story = () => tocAsRead(roomFile(room.room.document, file));
    const original = story();
    expect(await act(olive, refresh('entire'))).toMatchObject({ ok: true });
    const refreshed = story();
    expect(refreshed).toContain('<w:t>3</w:t>');
    expect(refreshed).toBe(tocAsRead(await alone(file, 'Olive', 'edit', refresh('entire'))));
    await act(olive, (e) => e.surface?.undo());
    expect(story()).toBe(original);
    await act(olive, (e) => e.surface?.redo());
    expect(story()).toBe(refreshed);
    converged(room, [olive]);
  });

  test("Word's rows: page numbers in PAGEREF fields, and a whole refresh, as the editor alone", async () => {
    const cases: readonly { steps: readonly Act[]; shows: RegExp }[] = [
      {
        steps: [refresh('pageNumbers')],
        // The second row's PAGEREF field shows the heading's new page
        shows:
          /PAGEREF _Toc \\h <\/w:instrText><w:fldChar w:fldCharType="separate"\/><w:t>2<\/w:t>/,
      },
      {
        steps: [
          (e) => (select(e, 3, 0), e.surface?.insertPlainText('New '), { ok: true }),
          refresh('entire'),
        ],
        shows: /<w:t>New Terms<\/w:t>/,
      },
    ];
    for (const { steps, shows } of cases) {
      const room = localRoom(WORD_TOC);
      const olive = await join(room, 'Olive', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: WORD_TOC,
      });
      for (const step of steps) expect(await act(olive, step)).toMatchObject({ ok: true });
      converged(room, [olive]);
      const editorAlone = await alone(WORD_TOC, 'Olive', 'edit', (editor) => {
        for (const step of steps) step(editor);
      });
      const saved = tocAsRead(roomFile(room.room.document, WORD_TOC));
      expect(saved).toBe(tocAsRead(editorAlone));
      expect(saved).toMatch(shows);
    }
  });

  test("the party rules: a party's table of contents is refused as a text change", async () => {
    const room = localRoom(HEADINGS, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: HEADINGS });
    const original = tocAsRead(roomFile(room.room.document, HEADINGS));
    await act(ana, insertToc);
    expect(ana.refusals.map((r) => `${r.code}:${r.reason}`)).toContain('ordered-refused:text');
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(tocAsRead(roomFile(room.room.document, HEADINGS))).toBe(original);
  });

  test('two organisers refresh it at the same time: the copies converge', async () => {
    const room = localRoom(HEADINGS);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: HEADINGS });
    const oscar = await join(room, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: HEADINGS });
    await act(olive, insertToc);
    await act(oscar, (e) => (select(e, 4, 0), e.surface?.insertPlainText('New ')));
    olive.editor.exec({ type: 'refreshToc', mode: 'entire' });
    oscar.editor.exec({ type: 'refreshToc', mode: 'entire' });
    await settle();
    converged(room, [olive, oscar]);
    expect(tocAsRead(roomFile(room.room.document, HEADINGS))).toContain('New Terms');
  });

  test('a saved file reopens in a new room, where a refresh still works', async () => {
    const first = localRoom(HEADINGS);
    const olive = await join(first, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: HEADINGS,
    });
    await act(olive, insertToc);
    const bytes = writeOoxmlPackage(roomFile(first.room.document, HEADINGS));
    const second = localRoom(bytes);
    const again = await join(second, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: bytes });
    await act(again, (e) => (select(e, 4, 0), e.surface?.insertPlainText('New ')));
    expect(await act(again, refresh('entire'))).toMatchObject({ ok: true });
    converged(second, [again]);
    expect(tocAsRead(roomFile(second.room.document, bytes))).toContain('New Terms');
  });
});

/** A contents row of a pasted table of contents: plain text, a tab and a page, in a link. */
const pastedRow = (text: string, bookmark: string, page: string) =>
  p(
    `<w:hyperlink w:anchor="${bookmark}" w:history="1">${t(text)}<w:r><w:tab/></w:r>${t(page)}</w:hyperlink>`,
    '<w:pStyle w:val="TOC3Pasted"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9734"/></w:tabs>'
  );

/**
 * A table of contents control with no TOC field, as a pasted copy holds it (the s25 acceptance
 * contract has two), and a paragraph after it.
 */
const PASTED_TOC = docx(
  '<w:sdt><w:sdtPr><w:id w:val="1808430988"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    pastedRow('Fees', '_bookmark0', '4') +
    pastedRow('Payment shall not', '_bookmark1', '5') +
    '</w:sdtContent></w:sdt>' +
    p(t('The fee is 1.200,00 per month.'))
);

describe('a table of contents with no field', () => {
  const READ_ONLY = 'the table of contents is generated and read-only';

  /** Each key in a row is the editor's own refusal: no change, nothing for the room to refuse. */
  const refusedInRows = async (person: Awaited<ReturnType<typeof join>>) => {
    for (const [row, offset] of [
      [0, 4],
      [1, 5],
    ] as const) {
      select(person.editor, row, offset);
      expect(await act(person, (e) => e.exec({ type: 'insertText', text: 'X' }))).toMatchObject({
        ok: false,
        reason: READ_ONLY,
      });
      await act(person, () => surface(person).insertPlainText('Y'));
      expect(surface(person).state().lastRejection).toBe(READ_ONLY);
    }
    expect(person.refusals).toEqual([]);
  };

  test('the organiser and a party are refused in its rows, and the next key elsewhere goes', async () => {
    const room = localRoom(PASTED_TOC, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: PASTED_TOC,
    });
    const ana = await join(room, 'Ana', { wordFile: PASTED_TOC });
    const original = storyAsRead(roomFile(room.room.document, PASTED_TOC));
    for (const person of [olive, ana]) await refusedInRows(person);
    converged(room, [olive, ana]);
    expect(storyAsRead(roomFile(room.room.document, PASTED_TOC))).toBe(original);
    // The paragraph after the control takes each person's next key, a direct edit and a suggestion
    for (const person of [olive, ana]) {
      select(person.editor, 2, 0);
      expect(
        await act(person, (e) => e.exec({ type: 'insertText', text: `${person.name}: ` }))
      ).toEqual({ ok: true, changed: true });
    }
    converged(room, [olive, ana]);
    const story = storyAsRead(roomFile(room.room.document, PASTED_TOC));
    expect(story).toContain('Olive: ');
    expect(story).toContain('Ana: ');
    // The saved file reopens in a new room, where the rows are still refused
    const bytes = writeOoxmlPackage(roomFile(room.room.document, PASTED_TOC));
    const again = localRoom(bytes);
    const oscar = await join(again, 'Oscar', { mode: 'edit', login: ORGANISER, wordFile: bytes });
    await refusedInRows(oscar);
    converged(again, [oscar]);
  });
});
