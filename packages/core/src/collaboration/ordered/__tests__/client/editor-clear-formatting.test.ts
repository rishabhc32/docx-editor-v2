// Clear All Formatting with real editors and one room (audit R1 row 7). The eraser takes off
// every direct run property of the selected words (a font, a size, a colour, shading, a
// language) and every direct paragraph property of the paragraphs it touches (alignment,
// indents, a list), as the editor alone does. The organiser's press is direct; a party's is a
// suggestion whose accept and reject each give the editor alone's result, also after a save and
// reopen.

import { afterEach, describe, expect, test } from 'bun:test';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import { closeAll, join, localRoom, ORGANISER, settle, type Person } from './editor-harness.ts';
import { alone, exec, roomBody, select, type Step } from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

const RED_ARIAL =
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:color w:val="FF0000"/><w:sz w:val="32"/>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/><w:lang w:val="fr-FR"/>';

/** A plain paragraph, then red Arial words in a centred, indented list item. */
const FILE = docx(
  p(t('Plain words.')) +
    p(
      t('Big red words.', RED_ARIAL),
      '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:jc w:val="center"/><w:ind w:left="1440"/>'
    ),
  MATRIX_EXTRAS
);

/** The eraser over "Big red" in the second paragraph. */
const clear: Step = (editor) => {
  select(editor, 0, 7, 1);
  exec(editor, { type: 'clearFormatting' });
};

async function people() {
  const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: FILE, review });
  const olive = await join(room, 'Olive', {
    wordFile: FILE,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, olive };
}

async function press(person: Person, step: Step): Promise<void> {
  step(person.editor);
  await settle();
}

describe('Clear All Formatting in a room (audit R1 row 7)', () => {
  test("the organiser's press takes every direct format off, as in the editor alone", async () => {
    const { room, ana, olive } = await people();
    await press(olive, clear);
    converged(room, ana, olive);
    const body = await roomBody(olive);
    expect(body).not.toContain('<w:jc');
    expect(body).not.toContain('<w:numPr');
    expect(body).toBe(await alone(FILE, review, [{ author: 'Olive', mode: 'edit', step: clear }]));
  });

  test("a party's press is a suggestion, as in the editor alone", async () => {
    const { room, ana, olive } = await people();
    await press(ana, clear);
    converged(room, ana, olive);
    expect(await roomBody(olive)).toBe(
      await alone(FILE, review, [{ author: 'Ana', mode: 'suggesting', step: clear }])
    );
  });

  for (const decide of ['accept', 'reject'] as const)
    test(`${decide} all of a party's press gives the editor alone's result`, async () => {
      const { room, ana, olive } = await people();
      await press(ana, clear);
      exec(olive.editor, { type: 'resolveAllReviewChanges', action: decide, scope: 'document' });
      await settle();
      converged(room, ana, olive);
      expect(await roomBody(olive)).toBe(
        await alone(FILE, review, [{ author: 'Ana', mode: 'suggesting', step: clear }], decide)
      );
    });
});
