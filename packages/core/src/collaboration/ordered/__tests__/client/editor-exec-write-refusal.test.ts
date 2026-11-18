// A write the room refuses, through `exec`, with two real editors and one room.
//
// Edits inside a table of contents' result are refused. `insertText` reported that refusal since
// row 32; every other write verb answered `{ ok: true, changed: false }`, so a host could not
// tell a refused delete from one that had nothing to delete. Each refused write now answers
// `ok: false` with its reason, for the organiser and for a party, and nothing reaches the room or
// the other person. (The room once refused the contents' heading too, where the editor alone
// takes edits; it refuses no more than the editor alone there now.) A write with
// truly nothing to change still answers `{ ok: true, changed: false }`.

import { afterEach, describe, expect, test } from 'bun:test';
import type { EditorCommand } from '../../../../contracts/editor.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  join,
  localRoom,
  ORGANISER,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

/** A Word table of contents (its heading, then the field over one row), then body text. */
const FILE = docx(
  '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>' +
    '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    p(t('Contents')) +
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>${t('Terms\t1')}</w:p>` +
    '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
    '</w:sdtContent></w:sdt>' +
    p(t('Body text.'))
);

/** Write verbs that answered `ok: true` for a refused write, with whether each needs a range. */
const REFUSED: readonly (readonly [EditorCommand, boolean])[] = [
  [{ type: 'deleteText' }, true],
  [{ type: 'cut' }, true],
  [{ type: 'paste', text: 'P', html: '<b>P</b>' }, false],
  [{ type: 'pasteWithoutFormatting', text: 'P' }, false],
  [{ type: 'insertBreak', kind: 'line' }, false],
  [{ type: 'insertBreak', kind: 'page' }, false],
  [{ type: 'toggleMark', mark: 'bold' }, true],
];

const label = (command: EditorCommand) =>
  'kind' in command ? `${command.type} ${String(command.kind)}` : command.type;

async function people() {
  const room = localRoom(FILE);
  const olive = await join(room, 'Olive', { wordFile: FILE, login: ORGANISER, mode: 'edit' });
  const ana = await join(room, 'Ana', { wordFile: FILE });
  return { room, olive, ana };
}

const bodyText = (person: Person) => surface(person).session.storyText({ kind: 'body' });

describe('a write the room refuses is refused through exec', () => {
  for (const who of ['Olive (organiser)', 'Ana (party)'] as const) {
    for (const [command, ranged] of REFUSED) {
      test(`${who}: ${label(command)} in the table of contents`, async () => {
        const { room, olive, ana } = await people();
        const person = who.startsWith('Olive') ? olive : ana;
        const before = bodyText(person);
        // In the result's row, "Terms"
        caretAt(person, 1, 1, ranged ? 4 : 1);

        expect(person.editor.exec(command)).toMatchObject({ ok: false, code: 'unsupported' });
        await settle();

        expect(person.refusals).toEqual([]);
        expect(bodyText(olive)).toBe(before);
        expect(bodyText(ana)).toBe(before);
        expect(sameAsRoom(olive, room)).toBe(true);
        expect(sameAsRoom(ana, room)).toBe(true);
      });
    }
  }
});

describe('a write with nothing to change in the room', () => {
  test('stays ok: true, changed: false for the organiser and a party', async () => {
    const { room, olive, ana } = await people();
    const body = 3;
    caretAt(olive, body, 2);
    expect(olive.editor.exec({ type: 'setAlignment', align: 'center' })).toEqual({
      ok: true,
      changed: true,
    });
    await settle();
    expect(sameAsRoom(ana, room)).toBe(true);

    for (const person of [olive, ana]) {
      caretAt(person, body, 2);
      expect(person.editor.exec({ type: 'setAlignment', align: 'center' })).toEqual({
        ok: true,
        changed: false,
      });
      await settle();
    }
    expect(olive.refusals).toEqual([]);
    expect(ana.refusals).toEqual([]);
    expect(sameAsRoom(olive, room)).toBe(true);
    expect(sameAsRoom(ana, room)).toBe(true);
    expect(bodyText(olive)).toBe(bodyText(ana));
  });
});

describe('an undo step the room refuses is refused through exec', () => {
  const review = reviewTestModule();
  const UNDO_REFUSAL = {
    ok: false,
    code: 'unsupported',
    reason: 'the collaboration session refused this undo',
  } as const;

  async function suggested() {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const olive = await join(room, 'Olive', {
      wordFile: FILE,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    caretAt(ana, 3, 5);
    surface(ana).insertPlainText('new ');
    await settle();
    return { room, olive, ana };
  }

  test('Ana (party): undo of a suggestion the organiser accepted', async () => {
    const { room, olive, ana } = await suggested();
    const accepted = olive.editor.exec({
      type: 'resolveAllReviewChanges',
      action: 'accept',
      scope: 'document',
    });
    expect(accepted.ok).toBe(true);
    await settle();
    const text = bodyText(ana);
    expect(text).toContain('Body new text.');

    expect(ana.editor.exec({ type: 'undo' })).toEqual(UNDO_REFUSAL);
    await settle();
    expect(ana.refusals.at(-1)).toMatchObject({ code: 'ordered-refused', reason: 'no-inverse' });
    expect(bodyText(ana)).toBe(text);
    expect(bodyText(olive)).toBe(text);
    expect(sameAsRoom(olive, room)).toBe(true);
    expect(sameAsRoom(ana, room)).toBe(true);
  });

  test('Olive (organiser): undo of a decision the room writes for good', async () => {
    const { room, olive, ana } = await suggested();
    const redated = olive.editor.exec({
      type: 'setReviewChangesAuthor',
      author: 'Ana',
      date: '2026-10-29T10:00:00Z',
      scope: 'document',
      authors: ['Ana'],
    });
    expect(redated.ok).toBe(true);
    await settle();

    expect(olive.editor.exec({ type: 'undo' })).toEqual(UNDO_REFUSAL);
    await settle();
    expect(olive.refusals.at(-1)).toMatchObject({ code: 'ordered-refused', reason: 'no-inverse' });
    expect(sameAsRoom(olive, room)).toBe(true);
    expect(sameAsRoom(ana, room)).toBe(true);
  });
});
