// Typing over a control's prompt while another person types after the control, with two real
// editors and one ordered room (task 9.6). Yjs rooms lost the text after the control in this
// case (fork fix 4): the same edit must keep it on every copy, and after a save and reopen.
// Inline and block controls, the other person's text right after the control and at the end of
// the paragraph that follows it. Also Backspace and Delete at a prompt's edge: inline prompts at
// two paragraphs' edges join as other text does, and the mark inside a block control's prompt
// stays, as the prompt's text does (Word changes a prompt only in Design Mode).

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedParagraphCount, orderedViewParagraph } from '../../compose/compose.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
  listTexts,
  localRoom,
  ORGANISER,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';

afterEach(closeAll);

const PROMPT = 'Click here to enter text.';
const PROMPT_FACE = '<w:rStyle w:val="PlaceholderText"/>';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/** "Name: <prompt> end." with an inline plain text control showing its prompt. */
const INLINE = docx(
  p(t('Name: ') + sdt('<w:showingPlcHdr/><w:text/>', t(PROMPT, PROMPT_FACE)) + t(' end.')) +
    p(t('Signed here.'))
);

/** A block control showing its prompt, then "end." in the paragraph after it. */
const BLOCK = docx(
  p(t('Name:')) +
    sdt('<w:showingPlcHdr/>', p(t(PROMPT, PROMPT_FACE))) +
    p(t('end.')) +
    p(t('Signed here.'))
);

/** Two paragraphs whose edges are inline controls showing their prompts. */
const EDGES = docx(
  p(t('A ') + sdt('<w:showingPlcHdr/><w:text/>', t('First.', PROMPT_FACE))) +
    p(sdt('<w:showingPlcHdr/><w:text/>', t('Second.', PROMPT_FACE)) + t(' tail'))
);

/** A block control whose prompt has two paragraphs, then "end.". */
const TWO_LINES = docx(
  sdt('<w:showingPlcHdr/>', p(t('First line.', PROMPT_FACE)) + p(t('Second line.', PROMPT_FACE))) +
    p(t('end.'))
);

function textsOf(doc: OrderedDocument): string[] {
  const out: string[] = [];
  for (let k = 0; k < orderedParagraphCount(doc, BODY); k += 1)
    out.push(
      orderedViewParagraph(doc, BODY, k)
        ?.chars.map((c) => c.ch)
        .join('') ?? ''
    );
  return out;
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
    expect({ person: person.name, refusals: person.refusals }).toEqual({
      person: person.name,
      refusals: [],
    });
    expect({ person: person.name, room: sameAsRoom(person, room) }).toEqual({
      person: person.name,
      room: true,
    });
    expect({ person: person.name, screen: screenFollowsList(person) }).toEqual({
      person: person.name,
      screen: true,
    });
    expect(editorTexts(person)).toEqual(listTexts(person));
  }
}

/** Every copy, and each editor's save reopened, shows `expected`. */
async function everywhere(
  room: ReturnType<typeof localRoom>,
  people: readonly Person[],
  expected: readonly string[]
): Promise<void> {
  expect(textsOf(room.room.document)).toEqual([...expected]);
  for (const person of people) {
    expect(editorTexts(person)).toEqual([...expected]);
    const saved = new Uint8Array(await person.editor.save());
    expect(readOoxmlPackage(saved).ok).toBe(true);
    expect(textsOf(openOrderedWordFile(saved, { file: 2 }).document)).toEqual([...expected]);
  }
}

describe("typing over a control's prompt while another person types after it", () => {
  test('inline control: the text right after the control survives', async () => {
    const room = localRoom(INLINE);
    const ana = await join(room, 'Ana', { wordFile: INLINE });
    const ben = await join(room, 'Ben', { wordFile: INLINE });
    expect(editorTexts(ana)[0]).toBe(`Name: ${PROMPT} end.`);
    caretAt(ana, 0, 8);
    caretAt(ben, 0, 6 + PROMPT.length);
    // Both keystrokes go out before either copy hears of the other
    surface(ana).insertPlainText('Bobby');
    surface(ben).insertPlainText(' Jr');
    await settle();
    converged(room, ana, ben);
    await everywhere(room, [ana, ben], ['Name: Bobby Jr end.', 'Signed here.']);
  });

  test('inline control: the text at the end of the paragraph survives', async () => {
    const room = localRoom(INLINE);
    const ana = await join(room, 'Ana', { wordFile: INLINE });
    const ben = await join(room, 'Ben', { wordFile: INLINE });
    caretAt(ana, 0, 6);
    caretAt(ben, 0, 6 + PROMPT.length + 5);
    surface(ana).insertPlainText('Bobby');
    surface(ben).insertPlainText(' Thanks.');
    await settle();
    converged(room, ana, ben);
    await everywhere(room, [ana, ben], ['Name: Bobby end. Thanks.', 'Signed here.']);
  });

  test('block control: the paragraph after the control keeps its text', async () => {
    const room = localRoom(BLOCK);
    const ana = await join(room, 'Ana', { wordFile: BLOCK });
    const ben = await join(room, 'Ben', { wordFile: BLOCK });
    expect(editorTexts(ana)).toEqual(['Name:', PROMPT, 'end.', 'Signed here.']);
    caretAt(ana, 1, 3);
    caretAt(ben, 2, 0);
    surface(ana).insertPlainText('Bobby');
    surface(ben).insertPlainText('The ');
    await settle();
    converged(room, ana, ben);
    await everywhere(room, [ana, ben], ['Name:', 'Bobby', 'The end.', 'Signed here.']);
    caretAt(ben, 2, 8);
    caretAt(ana, 1, 5);
    surface(ben).insertPlainText(' Done.');
    surface(ana).insertPlainText(' Smith');
    await settle();
    converged(room, ana, ben);
    await everywhere(room, [ana, ben], ['Name:', 'Bobby Smith', 'The end. Done.', 'Signed here.']);
  });
});

describe('the organiser types over a selected prompt, as Word does', () => {
  // Word edits a prompt only in Design Mode: deleting a selected prompt changes nothing, typed
  // words replace it whole (no `w:showingPlcHdr`), and the control's own prompt comes back
  // when the words go
  const cases = [
    { name: 'inline', file: INLINE, from: [0, 6], to: [0, 6 + PROMPT.length], at: [0, 6] },
    { name: 'block', file: BLOCK, from: [1, 0], to: [1, PROMPT.length], at: [1, 0] },
  ] as const;
  for (const { name, file, from, to, at } of cases)
    test(`${name} control`, async () => {
      const room = localRoom(file);
      const olive = await join(room, 'Olive', { wordFile: file, mode: 'edit', login: ORGANISER });
      const ben = await join(room, 'Ben', { wordFile: file });
      const shown = editorTexts(olive);
      const select = () => caretAt(olive, from[0], from[1], to[1], to[0]);
      select();
      surface(olive).deleteSelection();
      await settle();
      converged(room, olive, ben);
      expect(editorTexts(ben)).toEqual(shown);
      select();
      surface(olive).insertPlainText('Bobby');
      await settle();
      converged(room, olive, ben);
      expect(editorXml(olive)).not.toContain('<w:showingPlcHdr/>');
      const typed = editorTexts(olive);
      expect(typed.join('\n')).toContain('Bobby');
      await everywhere(room, [olive, ben], typed);
      caretAt(olive, at[0], at[1], at[1] + 'Bobby'.length);
      surface(olive).deleteSelection();
      await settle();
      converged(room, olive, ben);
      expect(editorXml(olive)).toContain('<w:showingPlcHdr/>');
      await everywhere(room, [olive, ben], shown);
    });
});

/** Backspace at the start of paragraph `index`, or Delete at the end of the one before it. */
function press(person: Person, key: 'Backspace' | 'Delete', index: number): void {
  if (key === 'Backspace') {
    caretAt(person, index, 0);
    surface(person).deleteBackward();
    return;
  }
  caretAt(person, index - 1, editorTexts(person)[index - 1]?.length ?? 0);
  surface(person).deleteForward();
}

describe("Backspace and Delete at a control's prompt", () => {
  const review = reviewTestModule();

  for (const key of ['Backspace', 'Delete'] as const)
    test(`${key}: inline prompts at two paragraphs' edges join as other text does`, async () => {
      const room = localRoom(EDGES);
      const ana = await join(room, 'Ana', { wordFile: EDGES, review });
      const olive = await join(room, 'Olive', {
        wordFile: EDGES,
        mode: 'edit',
        login: ORGANISER,
        review,
      });
      // The party's join is a suggestion: a struck paragraph mark the organiser accepts
      press(ana, key, 1);
      await settle();
      converged(room, ana, olive);
      expect(editorTexts(olive)).toEqual(['A First.', 'Second. tail']);
      expect(editorXml(olive)).toMatch(/<w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/);
      const accepted = olive.editor.exec({
        type: 'resolveAllReviewChanges',
        action: 'accept',
        scope: 'document',
      });
      expect(accepted.ok).toBe(true);
      await settle();
      converged(room, ana, olive);
      await everywhere(room, [ana, olive], ['A First.Second. tail']);
      // The organiser's join is direct, and her undo splits the paragraphs again
      const again = localRoom(EDGES);
      const olga = await join(again, 'Olga', { wordFile: EDGES, mode: 'edit', login: ORGANISER });
      const ben = await join(again, 'Ben', { wordFile: EDGES });
      press(olga, key, 1);
      await settle();
      converged(again, olga, ben);
      await everywhere(again, [olga, ben], ['A First.Second. tail']);
      surface(olga).undo();
      await settle();
      converged(again, olga, ben);
      await everywhere(again, [olga, ben], ['A First.', 'Second. tail']);
    });

  for (const offset of [0, PROMPT.length])
    test(`Enter at the ${offset === 0 ? 'start' : 'end'} of a block control's prompt: every screen shows the room's control`, async () => {
      const room = localRoom(BLOCK);
      const ana = await join(room, 'Ana', { wordFile: BLOCK });
      const olive = await join(room, 'Olive', {
        wordFile: BLOCK,
        mode: 'edit',
        login: ORGANISER,
      });
      caretAt(ana, 1, offset);
      surface(ana).splitParagraph();
      await settle();
      converged(room, ana, olive);
      // Word selects the whole prompt, so Enter at its edge replaces it too, and no glossary
      // block keeps it (Word 16.114, `~$p4-enter-tracked.docx`)
      expect(editorXml(olive)).not.toContain('<w:placeholder>');
      expect(editorXml(olive)).not.toContain(PROMPT);
      expect(editorXml(olive)).not.toContain('showingPlcHdr');
    });

  test("the mark inside a block control's prompt stays, for a party and the organiser", async () => {
    const room = localRoom(TWO_LINES);
    const ana = await join(room, 'Ana', { wordFile: TWO_LINES });
    const olive = await join(room, 'Olive', {
      wordFile: TWO_LINES,
      mode: 'edit',
      login: ORGANISER,
    });
    const shown = ['First line.', 'Second line.', 'end.'];
    expect(editorTexts(ana)).toEqual(shown);
    for (const person of [ana, olive])
      for (const key of ['Backspace', 'Delete'] as const) {
        press(person, key, 1);
        await settle();
        converged(room, ana, olive);
        expect(editorTexts(person)).toEqual(shown);
      }
    // A deletion across the mark leaves it too: a prompt is no content
    caretAt(ana, 0, 6, 3, 1);
    surface(ana).deleteSelection();
    await settle();
    converged(room, ana, olive);
    expect(editorXml(olive)).not.toContain('<w:del ');
    await everywhere(room, [ana, olive], shown);
  });
});

describe("typing into a block control's prompt of two paragraphs, as the editor alone does", () => {
  // The engine empties the control into the caret's paragraph and drops the prompt's other
  // paragraph (`clearPlaceholder`), as Word replaces a prompt whole; the whole prompt comes back
  // once the control is empty. The room once kept the other paragraph's mark, and saved the
  // prompt's second line as text
  const review = reviewTestModule();
  const shown = ['First line.', 'Second line.', 'end.'];
  const cases = [
    { name: 'a caret in the first line', at: [0, 3, 3, 0] },
    { name: 'a caret in the second line', at: [1, 3, 3, 1] },
    { name: 'the whole prompt selected', at: [0, 0, 12, 1] },
  ] as const;
  for (const who of ['party', 'organiser'] as const)
    for (const { name, at } of cases)
      test(`${who}: ${name}, then the prompt comes back whole`, async () => {
        const room = localRoom(TWO_LINES);
        const ana = await join(
          room,
          'Ana',
          who === 'party'
            ? { wordFile: TWO_LINES, review }
            : { wordFile: TWO_LINES, mode: 'edit', login: ORGANISER, review }
        );
        const olive = await join(room, 'Olive', {
          wordFile: TWO_LINES,
          mode: 'edit',
          login: ORGANISER,
          review,
        });
        caretAt(ana, at[0], at[1], at[2], at[3]);
        surface(ana).insertPlainText('Bobby');
        await settle();
        converged(room, ana, olive);
        expect(editorXml(olive)).not.toContain('PlaceholderText');
        await everywhere(room, [ana, olive], ['Bobby', 'end.']);
        // A party's words rejected, the organiser's undone: the control is empty again
        if (who === 'party')
          expect(
            olive.editor.exec({
              type: 'resolveAllReviewChanges',
              action: 'reject',
              scope: 'document',
            }).ok
          ).toBe(true);
        else surface(ana).undo();
        await settle();
        converged(room, ana, olive);
        expect(editorXml(olive)).toContain('<w:showingPlcHdr/>');
        await everywhere(room, [ana, olive], shown);
      });

  test('the organiser deletes her words: the whole prompt comes back', async () => {
    const room = localRoom(TWO_LINES);
    const olive = await join(room, 'Olive', {
      wordFile: TWO_LINES,
      mode: 'edit',
      login: ORGANISER,
    });
    const ben = await join(room, 'Ben', { wordFile: TWO_LINES });
    caretAt(olive, 1, 0);
    surface(olive).insertPlainText('Bobby');
    await settle();
    converged(room, olive, ben);
    caretAt(olive, 0, 0, 5, 0);
    surface(olive).deleteSelection();
    await settle();
    converged(room, olive, ben);
    await everywhere(room, [olive, ben], shown);
  });

  test('text typed after the control stays there when the prompt comes back', async () => {
    const room = localRoom(TWO_LINES);
    const ana = await join(room, 'Ana', { wordFile: TWO_LINES, review });
    const olive = await join(room, 'Olive', {
      wordFile: TWO_LINES,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    // The prompt's last paragraph goes: the paragraph after the control starts right after it
    caretAt(ana, 0, 0);
    surface(ana).insertPlainText('Bobby');
    await settle();
    converged(room, ana, olive);
    caretAt(olive, 1, 0);
    surface(olive).insertPlainText('The ');
    await settle();
    converged(room, ana, olive);
    await everywhere(room, [ana, olive], ['Bobby', 'The end.']);
    olive.editor.exec({ type: 'resolveAllReviewChanges', action: 'reject', scope: 'document' });
    await settle();
    converged(room, ana, olive);
    await everywhere(room, [ana, olive], ['First line.', 'Second line.', 'The end.']);
  });

  test('two people type into the two lines at once: both words stay in the control', async () => {
    // The first words replace the prompt, so the other line is gone when the second words come:
    // they go to the control's paragraph that stays, never past the control's end
    const room = localRoom(TWO_LINES);
    const ana = await join(room, 'Ana', { wordFile: TWO_LINES });
    const ben = await join(room, 'Ben', { wordFile: TWO_LINES });
    caretAt(ana, 0, 0);
    caretAt(ben, 1, 0);
    surface(ana).insertPlainText('Bobby');
    surface(ben).insertPlainText('Smith');
    await settle();
    converged(room, ana, ben);
    await everywhere(room, [ana, ben], ['BobbySmith', 'end.']);
  });

  test('the organiser removes the control and keeps its words, then undoes it', async () => {
    const room = localRoom(TWO_LINES);
    const olive = await join(room, 'Olive', {
      wordFile: TWO_LINES,
      mode: 'edit',
      login: ORGANISER,
    });
    const ben = await join(room, 'Ben', { wordFile: TWO_LINES });
    caretAt(olive, 0, 0);
    surface(olive).insertPlainText('Bobby');
    await settle();
    caretAt(olive, 0, 2);
    expect(olive.editor.exec({ type: 'removeContentControl' } as never)).toMatchObject({
      ok: true,
    });
    await settle();
    converged(room, olive, ben);
    expect(editorXml(olive)).not.toContain('<w:sdt>');
    await everywhere(room, [olive, ben], ['Bobby', 'end.']);
    surface(olive).undo();
    await settle();
    converged(room, olive, ben);
    expect(editorXml(olive)).toContain('<w:sdt>');
    await everywhere(room, [olive, ben], ['Bobby', 'end.']);
  });
});
