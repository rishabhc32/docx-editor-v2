// Increase Indent to a list level that the list's definition does not declare, with real editors
// and one room (audit R1 row 1). The editor alone declares the level in `numbering.xml` first, in
// Word's default format for that depth, so the item keeps a number at its new level. The room
// declares it on every copy and in the saved file, for the organiser's direct press and a
// party's suggested one, also when two people press at the same time.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import {
  closeAll,
  HeldTransport,
  join,
  localRoom,
  ORGANISER,
  settle,
  type Editor,
  type Person,
} from './editor-harness.ts';
import { alone, exec, roomBody, select, type Step } from './format-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { converged } from './tables-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const review = reviewTestModule();

/** A plain paragraph and three items of a list whose definition declares only level 0. */
const item = (text: string) =>
  p(t(text), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
const FILE = docx(
  p(t('Plain.')) + item('First item.') + item('Second item.') + item('Third item.'),
  MATRIX_EXTRAS
);

/** Increase Indent with the caret in the paragraph at `index`. */
const indent =
  (index: number): Step =>
  (editor) => {
    select(editor, 3, 3, index);
    exec(editor, { type: 'adjustIndent', direction: 'increase' });
  };

/** The `w:lvl` elements of the file's first list definition, as XML. */
function levelsOf(xml: string): string[] {
  const abstract = /<w:abstractNum w:abstractNumId="0"[\s\S]*?<\/w:abstractNum>/.exec(xml)?.[0];
  return abstract?.match(/<w:lvl [\s\S]*?<\/w:lvl>/g) ?? [];
}

/** The list levels an editor's package declares, and the same after a save and reopen. */
async function levels(editor: Editor): Promise<string[]> {
  const s = editor.surface;
  if (!s) throw new Error('no surface');
  const live = s.session.currentPackage().parts.get('/word/numbering.xml');
  const read = readOoxmlPackage(new Uint8Array(await editor.save()), {});
  if (!read.ok || !live) throw new Error('no numbering');
  const saved = read.package.parts.get('/word/numbering.xml');
  if (!saved) throw new Error('no saved numbering');
  expect(levelsOf(serializeOoxmlPart(saved))).toEqual(levelsOf(serializeOoxmlPart(live)));
  return levelsOf(serializeOoxmlPart(live));
}

/** The body and the list levels the editor alone makes of the steps. */
async function aloneWith(
  steps: readonly { author: string; mode: 'suggesting' | 'edit'; step: Step }[],
  decide?: 'accept' | 'reject'
): Promise<{ body: string; levels: string[] }> {
  let found: string[] = [];
  const last = steps.at(-1);
  if (!last) throw new Error('no steps');
  const body = await alone(
    FILE,
    review,
    [
      ...steps.slice(0, -1),
      {
        ...last,
        step: (editor) => {
          last.step(editor);
          const numbering = editor.surface?.session
            .currentPackage()
            .parts.get('/word/numbering.xml');
          found = numbering ? levelsOf(serializeOoxmlPart(numbering)) : [];
        },
      },
    ],
    decide
  );
  return { body, levels: found };
}

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

describe('Increase Indent to a list level the list does not declare (audit R1 row 1)', () => {
  for (const who of ['Olive', 'Ana'] as const)
    test(`${who}'s press declares the level on every copy and in the file, as in the editor alone`, async () => {
      const { room, ana, olive } = await people();
      const person = who === 'Olive' ? olive : ana;
      await press(person, indent(2));
      converged(room, ana, olive);
      const expected = await aloneWith([
        { author: who, mode: who === 'Olive' ? 'edit' : 'suggesting', step: indent(2) },
      ]);
      expect(expected.levels.length).toBe(2);
      for (const copy of [ana, olive]) expect(await levels(copy.editor)).toEqual(expected.levels);
      expect(await roomBody(olive)).toBe(expected.body);
    });

  for (const decide of ['accept', 'reject'] as const)
    test(`${decide} all of a party's press gives the editor alone's paragraphs`, async () => {
      const { room, ana, olive } = await people();
      await press(ana, indent(2));
      exec(olive.editor, { type: 'resolveAllReviewChanges', action: decide, scope: 'document' });
      await settle();
      converged(room, ana, olive);
      const expected = await aloneWith(
        [{ author: 'Ana', mode: 'suggesting', step: indent(2) }],
        decide
      );
      expect(await roomBody(olive)).toBe(expected.body);
      // A rejected press's level goes with its suggestion; the editor alone keeps it unused
      const declared = await levels(olive.editor);
      expect(declared.length).toBe(decide === 'accept' ? 2 : 1);
    });

  test('one press over two items moves both, as in the editor alone', async () => {
    const { room, ana, olive } = await people();
    const both: Step = (editor) => {
      const s = editor.surface;
      if (!s) throw new Error('no surface');
      const [, , second, third] = s.session.paragraphIds();
      if (!second || !third) throw new Error('no items');
      s.setSelection({
        anchor: { paragraphId: second, offset: 0 },
        head: { paragraphId: third, offset: 3 },
      });
      exec(editor, { type: 'adjustIndent', direction: 'increase' });
    };
    await press(ana, both);
    converged(room, ana, olive);
    const expected = await aloneWith([{ author: 'Ana', mode: 'suggesting', step: both }]);
    expect(await levels(olive.editor)).toEqual(expected.levels);
    const body = await roomBody(olive);
    expect(body.split('<w:ilvl w:val="1"/>').length - 1).toBe(2);
    expect(body).toBe(expected.body);
  });

  test('two people who indent at the same time both reach the level', async () => {
    const room = localRoom(FILE, { policy: suggestionsOnlyPolicy() });
    const ana = await join(room, 'Ana', { wordFile: FILE, review });
    const held = new HeldTransport(room.transport({ person: testPerson('Olive'), ...ORGANISER }));
    const olive = await join(room, 'Olive', {
      wordFile: FILE,
      mode: 'edit',
      login: ORGANISER,
      transport: held,
      review,
    });
    held.held = true;
    await press(ana, indent(2));
    await press(olive, indent(3));
    held.held = false;
    held.release();
    await settle();
    converged(room, ana, olive);
    for (const copy of [ana, olive]) expect((await levels(copy.editor)).length).toBe(2);
    expect((await roomBody(olive)).split('<w:ilvl w:val="1"/>').length - 1).toBe(2);
  });

  test('undo and redo of the press, as in the editor alone', async () => {
    const { room, ana, olive } = await people();
    const undone: Step = (editor) => {
      indent(2)(editor);
      exec(editor, { type: 'undo' });
    };
    await press(ana, indent(2));
    exec(ana.editor, { type: 'undo' });
    await settle();
    converged(room, ana, olive);
    expect(await roomBody(olive)).toBe(
      (await aloneWith([{ author: 'Ana', mode: 'suggesting', step: undone }])).body
    );
    exec(ana.editor, { type: 'redo' });
    await settle();
    converged(room, ana, olive);
    expect((await levels(olive.editor)).length).toBe(2);
    expect(await roomBody(olive)).toBe(
      (await aloneWith([{ author: 'Ana', mode: 'suggesting', step: indent(2) }])).body
    );
  });
});
