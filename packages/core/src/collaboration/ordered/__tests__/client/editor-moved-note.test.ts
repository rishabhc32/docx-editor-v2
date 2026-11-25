// A Word file with a pending move of words that hold a footnote reference, as a room writes it:
// both halves name the one note. Real editors and one ordered room against the editor alone, the
// move accepted and rejected live, and after one and two folds. The room once read the note as a
// new note of the move on every load: a reject removed the note, where the editor alone keeps it.

import { afterEach, describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { openOrderedWordFile, type OrderedWordFile } from '../../client/open.ts';
import { orderedPackageMedia } from '../../compose/checkpoint.ts';
import { checkpointBytes, folded, mediaStore } from '../fold-check.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import { closeAll, join, localRoom, ORGANISER, sameAsRoom, settle } from './editor-harness.ts';
import { select } from './links-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const STAMP = 'w:author="Olive" w:date="2026-11-24T12:00:00Z"';
const MOVED =
  '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>' +
  t(' Terms');
const BODY =
  p(
    t('Payment') +
      `<w:moveFromRangeStart w:id="11" ${STAMP} w:name="move1"/>` +
      `<w:moveFrom w:id="12" ${STAMP}>${MOVED}</w:moveFrom><w:moveFromRangeEnd w:id="11"/>` +
      t(' apply here.')
  ) +
  p(
    t('The') +
      `<w:moveToRangeStart w:id="13" ${STAMP} w:name="move1"/>` +
      `<w:moveTo w:id="14" ${STAMP}>${MOVED}</w:moveTo><w:moveToRangeEnd w:id="13"/>` +
      t(' second paragraph.')
  );
const FILE = docx(BODY, MATRIX_EXTRAS);

/** The Word file's media, which the room's storage holds, as a host stores them. */
const MEDIA = new Map(
  orderedPackageMedia(openOrderedWordFile(FILE, { file: 1 }).base).map((m) => [m.sha256, m.bytes])
);

type Decision = 'accept' | 'reject';
type Editor = ReturnType<typeof createDocxEditor>;

/** The paragraphs with a footnote reference, and the footnotes with their text, as saved. */
async function notesOf(editor: Editor): Promise<string[]> {
  const files = unzipSync(new Uint8Array(await editor.save()));
  const body = strFromU8(files['word/document.xml']!);
  const notes = strFromU8(files['word/footnotes.xml']!);
  return [
    ...[...body.matchAll(/<w:p[ >](?:(?!<\/w:p>).)*<\/w:p>/gs)]
      .filter((m) => m[0].includes('<w:footnoteReference w:id="1"/>'))
      .map((m) => `reference in "${m[0].replace(/<[^>]+>/g, '')}"`),
    ...[...notes.matchAll(/<w:footnote w:id="(\d+)">(.*?)<\/w:footnote>/gs)].map(
      (m) => `note ${m[1]}:${m[2]!.replace(/<[^>]+>/g, '')}`
    ),
  ];
}

/** Decide every card of an editor: the move's two halves. */
function decide(editor: Editor, decision: Decision): void {
  const result = editor.exec({
    type: 'resolveAllReviewChanges',
    action: decision,
    scope: 'document',
  });
  expect(result.ok).toBe(true);
}

/** The organiser's edit before the decision, if any. */
type Before = (editor: Editor) => void;

/** What the editor alone saves after `before` and the decision. */
async function alone(file: Uint8Array, decision: Decision, before?: Before): Promise<string[]> {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olga',
    modules: [reviewTestModule()],
  });
  try {
    before?.(editor);
    decide(editor, decision);
    return await notesOf(editor);
  } finally {
    editor.destroy();
    container.remove();
  }
}

/**
 * What the organiser saves after `before` and the decision, in a room on `file` with a party
 * beside.
 */
async function inRoom(
  file: Uint8Array,
  number: number,
  decision: Decision,
  before?: Before
): Promise<string[]> {
  const review = reviewTestModule();
  const room = localRoom(file, { file: number, blobs: MEDIA });
  const olive = await join(room, 'Olive', { login: ORGANISER, wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  if (before) {
    before(olive.editor);
    await settle();
  }

  decide(olive.editor, decision);
  await settle();
  for (const person of [olive, ben]) {
    expect({
      person: person.name,
      refusals: person.refusals,
      same: sameAsRoom(person, room),
    }).toEqual({ person: person.name, refusals: [], same: true });
  }

  return notesOf(olive.editor);
}

/** The room's checkpoint on the Word file, folded `times` times before it is written. */
async function foldedBytes(times: number): Promise<Uint8Array> {
  let copy: OrderedWordFile = openOrderedWordFile(FILE, { file: 1 });
  const blobs = mediaStore(copy.base);
  for (let fold = 1; fold < times; fold += 1) {
    copy = await folded(copy.document, copy.base, fold + 1, blobs);
  }

  return checkpointBytes(copy.document, copy.base, blobs);
}

describe('a pending move of words that hold a footnote reference, as the editor alone', () => {
  for (const decision of ['accept', 'reject'] as const) {
    test(`${decision}ed live and after one and two folds`, async () => {
      const expected = await alone(FILE, decision);
      expect(expected).toEqual([
        decision === 'accept'
          ? 'reference in "The Terms second paragraph."'
          : 'reference in "Payment Terms apply here."',
        'note 1: Fees are due net of tax.',
      ]);
      const copies = [
        ['live', FILE],
        ['folded once', await foldedBytes(1)],
        ['folded twice', await foldedBytes(2)],
      ] as const;
      for (const [copy, bytes] of copies) {
        expect({ copy, notes: await inRoom(bytes, 2, decision) }).toEqual({
          copy,
          notes: expected,
        });
      }
    });
  }
});

// Bob's pending deletion in the note, and the organiser's direct deletion of the copy's
// reference: the struck source still names the note, so the editor alone keeps the note whole,
// and removes it with the source's reference on an accept. The room once struck the note's text
// as gone but for Bob's words, which then stood with no paragraph mark: a save refused the notes
// story, or a fold put them in the next note
describe("the organiser's direct deletion of the copy's reference, as the editor alone", () => {
  const notes = MATRIX_EXTRAS.parts?.['word/footnotes.xml'];
  const bobs =
    t(' Fees are due') +
    '<w:del w:id="30" w:author="Bob" w:date="2026-11-24T12:00:00Z"><w:r>' +
    '<w:delText xml:space="preserve"> net</w:delText></w:r></w:del>' +
    t(' of tax.');
  const file = docx(BODY, {
    ...MATRIX_EXTRAS,
    parts: {
      ...MATRIX_EXTRAS.parts,
      'word/footnotes.xml': String(notes).replace(t(' Fees are due net of tax.'), bobs),
    },
  });
  const deleteCopy = (editor: Editor) => {
    editor.setMode('edit');
    select(editor, 1, 3, 4);
    editor.surface?.deleteBackward();
  };
  for (const decision of ['accept', 'reject'] as const) {
    test(`${decision}ed after it`, async () => {
      const expected = await alone(file, decision, deleteCopy);
      expect(expected.some((line) => line.startsWith('note 1'))).toBe(decision === 'reject');
      expect(await inRoom(file, 2, decision, deleteCopy)).toEqual(expected);
    });
  }
});
