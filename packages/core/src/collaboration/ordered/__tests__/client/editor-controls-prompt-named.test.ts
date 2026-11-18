// A control that names a glossary prompt, once the content that replaced its prompt went, with
// real editors and one ordered room (task 9.6).
//
// Typing over a shown prompt drops the prompt's runs with no revision, as Word does
// (`~$p4-enter-tracked.docx`). When the control is empty again, Word and the editor alone show
// the block its `w:placeholder` names (ECMA-376 §17.5.2.25): after a reject of the typing, an
// accepted deletion of it, or a direct deletion of it. Only an undo gives the replaced runs back.
// The room once showed the replaced runs again, with another person's pending format on them and
// its card, where the editor alone and a file saved before the reject showed the glossary's
// prompt. Each step here is compared with the editor alone, on every copy and in the save.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t, W } from '../word-fixtures.ts';
import {
  caretAt,
  closeAll,
  editorTexts,
  editorXml,
  join,
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

const FACE = '<w:rStyle w:val="PlaceholderText"/>';
/** Bob's pending bold on the prompt's words " here to ". */
const BOBS =
  `${FACE}<w:b/><w:rPrChange w:id="7" w:author="Bob" w:date="2026-10-28T08:00:00Z">` +
  `<w:rPr>${FACE}</w:rPr></w:rPrChange>`;
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const promptRuns = (bob: boolean) =>
  t('Click or tap', FACE) + t(' here to ', bob ? BOBS : FACE) + t('enter text.', FACE);
/** The glossary's prompt differs from the runs the body shows, as after an edit of either. */
const GLOSSARY = {
  parts: {
    'word/glossary/document.xml':
      `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>` +
      '<w:name w:val="PartyName"/><w:category><w:name w:val="General"/>' +
      '<w:gallery w:val="placeholder"/></w:category></w:docPartPr>' +
      `<w:docPartBody><w:p>${t('Enter the party.', FACE)}</w:p></w:docPartBody></w:docPart>` +
      '</w:docParts></w:glossaryDocument>',
  },
  contentTypes:
    '<Override PartName="/word/glossary/document.xml" ContentType="application/' +
    'vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>',
  relationships:
    '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/' +
    '2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
};
const NAMED = '<w:placeholder><w:docPart w:val="PartyName"/></w:placeholder><w:showingPlcHdr/>';
const inline = (bob: boolean) =>
  docx(
    p(t('Name: ') + sdt(`${NAMED}<w:text/>`, promptRuns(bob)) + t(' end.')) + p(t('Signed.')),
    GLOSSARY
  );
const block = (bob: boolean) =>
  docx(p(t('Name:')) + sdt(NAMED, p(promptRuns(bob))) + p(t('Signed.')), GLOSSARY);

/** Each shape: its file, the prompt's line, a caret in Bob's words, the line once filled. */
const SHAPES = [
  { name: 'inline', file: inline, line: 0, at: 21, filled: 'Name: Q end.' },
  { name: 'block', file: block, line: 1, at: 15, filled: 'Q' },
] as const;
type Shape = (typeof SHAPES)[number];

/**
 * One step of the organiser's: typing "Q" over the prompt as a suggestion or directly, a
 * decision on her typing, a direct deletion of it, Enter, undo, or redo.
 */
type Step = 'suggest' | 'type' | 'reject' | 'accept' | 'empty' | 'enter' | 'undo' | 'redo';

const bobsCards = (person: Person) =>
  person.editor
    .getReviewItems()
    .filter((item) => item.kind === 'revision' && 'author' in item && item.author === 'Bob');

function decide(person: Person, action: 'accept' | 'reject'): void {
  const card = person.editor
    .getReviewItems()
    .find((item) => item.kind === 'revision' && 'author' in item && item.author === 'Olga');
  if (card === undefined) throw new Error("no card of Olga's");
  const result =
    action === 'accept'
      ? person.editor.acceptReviewItem(card.key)
      : person.editor.rejectReviewItem(card.key);
  expect(result.ok).toBe(true);
}

function step(person: Person, what: Step, { line, at, name }: Shape): void {
  person.editor.setMode(what === 'suggest' ? 'suggesting' : 'edit');
  if (what === 'accept' || what === 'reject') decide(person, what);
  else if (what === 'suggest' || what === 'type') {
    caretAt(person, line, at);
    surface(person).insertPlainText('Q');
  } else if (what === 'empty') {
    const start = name === 'inline' ? 6 : 0;
    caretAt(person, line, start, start + 1);
    surface(person).deleteBackward();
  } else if (what === 'enter') {
    caretAt(person, line, at);
    surface(person).splitParagraph();
  } else if (what === 'undo') surface(person).undo();
  else surface(person).redo();
}

/** The control without ids, dates and paragraph ids, which the room and the editor mint apart. */
const stable = (xml: string): string =>
  (/<w:sdt>.*<\/w:sdt>/s.exec(xml)?.[0] ?? '<none>')
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:(id|date)="[^"]*"/g, '');

/** The control after each step on the editor alone. */
function aloneSteps(file: Uint8Array, steps: readonly Step[], shape: Shape) {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olga',
    mode: 'edit',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container, refusals: [] } as unknown as Person;
    return steps.map((what) => {
      step(person, what, shape);
      return { control: stable(editorXml(person)), bobs: bobsCards(person).length };
    });
  } finally {
    editor.destroy();
    container.remove();
  }
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
  }
}

async function saved(person: Person): Promise<{ control: string; bobs: number }> {
  const bytes = new Uint8Array(await person.editor.save());
  const again = openOrderedWordFile(bytes, { file: 2 }).document;
  return {
    control: stable(editorXml(person)),
    bobs: orderedCards(again).filter((card) => card.author === 'Bob').length,
  };
}

const SEQUENCES: readonly (readonly Step[])[] = [
  ['suggest', 'reject'],
  ['suggest', 'undo', 'redo', 'reject'],
  ['suggest', 'accept', 'empty', 'undo'],
  ['type', 'empty', 'undo', 'undo'],
  ['type', 'empty', 'undo', 'redo'],
];

for (const bob of [true, false])
  describe(`the organiser's typing over a named prompt ${bob ? "with Bob's format on it" : 'that differs from the glossary'}`, () => {
    for (const shape of SHAPES)
      for (const steps of SEQUENCES)
        test(`${shape.name}: ${steps.join(', ')}, as the editor alone`, async () => {
          const file = shape.file(bob);
          const expected = aloneSteps(file, steps, shape);
          const review = reviewTestModule();
          const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
          const olga = await join(room, 'Olga', {
            mode: 'edit',
            login: ORGANISER,
            wordFile: file,
            review,
          });
          const cy = await join(room, 'Cy', { wordFile: file, review });
          const shown: { control: string; bobs: number }[] = [];
          for (const what of steps) {
            step(olga, what, shape);
            await settle();
            converged(room, olga, cy);
            if (what === 'suggest' || what === 'type')
              expect(editorTexts(cy)[shape.line]).toBe(shape.filled);
            expect(stable(editorXml(olga))).toBe(stable(editorXml(cy)));
            shown.push({ control: stable(editorXml(cy)), bobs: bobsCards(cy).length });
          }
          expect(shown).toEqual(expected);
          expect(await saved(cy)).toEqual(expected.at(-1) ?? { control: '', bobs: 0 });
        });
  });

describe("after a reject, the glossary's prompt is the prompt", () => {
  for (const shape of SHAPES)
    test(`${shape.name}: Bob's format is gone, so a party's typing over it is a suggestion`, async () => {
      const file = shape.file(true);
      const review = reviewTestModule();
      const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
      const olga = await join(room, 'Olga', {
        mode: 'edit',
        login: ORGANISER,
        wordFile: file,
        review,
      });
      const cy = await join(room, 'Cy', { wordFile: file, review });
      step(olga, 'suggest', shape);
      await settle();
      step(olga, 'reject', shape);
      await settle();
      converged(room, olga, cy);
      expect(editorTexts(cy)[shape.line]).toContain('Enter the party.');
      expect(bobsCards(cy)).toHaveLength(0);
      caretAt(cy, shape.line, shape.name === 'inline' ? 6 : 0);
      surface(cy).insertPlainText('Q');
      await settle();
      converged(room, olga, cy);
      expect(editorTexts(olga)[shape.line]).toBe(shape.filled);
      expect(editorXml(olga)).toMatch(/<w:ins [^>]*w:author="Cy"/);
    });
});
