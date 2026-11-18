// A long session keeps a bounded number of numbering indexes alive, in a room and without one.
//
// Each list edit gives layout a new numbering index. The undo history keeps the parts of earlier
// revisions, and layout memos keyed by those parts or their first paragraphs kept the index of
// their own pass: an editor in a long session kept one whole index per header edit until it ran
// out of memory, when list edits came between the header edits.

import { afterEach, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { numberingIndexTestRecorder } from '../../../../layout/numbering-index.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { bodyOf, docx, MATRIX_EXTRAS, MATRIX_ITEMS, p, R, t, W } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Editor,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const listItem = (text: string) =>
  p(t(text), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
const box = MATRIX_ITEMS.find((item) => item.key === 'textbox');
if (!box) throw new Error('no text box item');
const FILE = docx(
  p(t('Plain one.')) +
    p(t('Plain gap.')) +
    listItem('First item.') +
    bodyOf(box) +
    listItem('Second item.') +
    p(t('Plain two.')) +
    `<w:sectPr xmlns:r="${R}"><w:headerReference w:type="default" r:id="rIdHeader1"/></w:sectPr>`,
  {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/header1.xml': `<w:hdr xmlns:w="${W}">${listItem('Header item.')}${p(t('Header.'))}</w:hdr>`,
    },
    relationships:
      `${MATRIX_EXTRAS.relationships}` +
      `<Relationship Id="rIdHeader1" Type="${R}/header" Target="header1.xml"/>`,
    contentTypes:
      `${MATRIX_EXTRAS.contentTypes}` +
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
  }
);

/** Type at the start of the header's first paragraph, and come back to the body. */
function typeInHeader(s: NonNullable<Editor['surface']>, text: string): void {
  expect(s.enterHeaderFooter({ rId: 'rIdHeader1', kind: 'header' })).toBe(true);
  const first = s.session.paragraphIdsIn({ kind: 'headerFooter', rId: 'rIdHeader1' })[0];
  if (first === undefined) throw new Error('no header paragraph');
  s.setSelection({
    anchor: { paragraphId: first, offset: 0 },
    head: { paragraphId: first, offset: 0 },
  });
  s.insertPlainText(text);
  s.setActiveScope({ kind: 'body' });
}

/**
 * One round: the organiser takes the first list off and makes it again (a new list definition)
 * and types in the header, where a party may not write; each person types in the body, undoes
 * and redoes; the organiser accepts a suggestion.
 */
async function round(olive: Person, ana: Person): Promise<void> {
  for (let k = 0; k < 2; k += 1) {
    caretAt(olive, 0, 0);
    expect(olive.editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
    await settle();
  }
  typeInHeader(surface(olive), 'O');
  await settle();
  for (const person of [ana, olive]) {
    caretAt(person, 5, 0);
    surface(person).insertPlainText(person.name[0]!);
    await settle();
    expect(person.handle.session.undo()).toBe(true);
    await settle();
    expect(person.handle.session.redo()).toBe(true);
    await settle();
  }
  const card = olive.editor.getReviewItems().find((item) => item.kind === 'revision');
  if (card) expect(olive.editor.acceptReviewItem(card.key).ok).toBe(true);
  await settle();
}

test('a long session keeps a bounded number of numbering indexes', async () => {
  const review = reviewTestModule();
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  caretAt(olive, 0, 0);
  expect(olive.editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
  await settle();
  expect(room.room.joinSnapshot(0).map((change) => change.kind)).toContain('new-list');
  const recorder = numberingIndexTestRecorder();
  try {
    const liveAfter = async (rounds: number): Promise<number> => {
      for (let k = 0; k < rounds; k += 1) await round(olive, ana);
      await settle();
      Bun.gc(true);
      return recorder.live;
    };
    const early = await liveAfter(4);
    const builtEarly = recorder.built;
    const late = await liveAfter(8);
    // Each round builds new indexes; the live ones stay as many as the open revisions need
    expect(recorder.built - builtEarly).toBeGreaterThanOrEqual(8);
    expect(late - early).toBeLessThanOrEqual(2);
  } finally {
    recorder.dispose();
  }
  for (const person of [olive, ana]) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
  }
});

test('one editor without a room keeps a bounded number of numbering indexes', async () => {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({ container, document: FILE, mode: 'edit' });
  const recorder = numberingIndexTestRecorder();
  try {
    await settle();
    const s = editor.surface;
    if (!s) throw new Error('no surface');
    const first = s.session.paragraphIds()[0]!;
    const liveAfter = async (rounds: number): Promise<number> => {
      for (let k = 0; k < rounds; k += 1) {
        typeInHeader(s, 'h');
        await settle();
        // A list on and off again: a new list definition, so a new numbering part
        s.setSelection({
          anchor: { paragraphId: first, offset: 0 },
          head: { paragraphId: first, offset: 0 },
        });
        expect(editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
        await settle();
        expect(editor.exec({ type: 'toggleList', kind: 'bullet' }).ok).toBe(true);
        await settle();
      }
      Bun.gc(true);
      return recorder.live;
    };
    const early = await liveAfter(4);
    const builtEarly = recorder.built;
    const late = await liveAfter(8);
    expect(recorder.built - builtEarly).toBeGreaterThanOrEqual(8);
    expect(late - early).toBeLessThanOrEqual(2);
  } finally {
    recorder.dispose();
    editor.destroy();
    container.remove();
  }
});
