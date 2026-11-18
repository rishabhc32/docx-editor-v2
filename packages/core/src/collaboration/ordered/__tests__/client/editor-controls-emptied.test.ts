// A decision that empties a content control which names no prompt, with real editors and one
// ordered room. Word 16.114 (2026-10-30, samples `~$i6w-*` and `~$i6n-*`): Accept All of a
// deletion of all the control's text, and Reject All of an insertion of all of it, leave the
// control showing Word's blank prompt: `w:showingPlcHdr`, an empty `w:sdtEndPr`, no
// `w:richText`, and one run of five spaces in no style, inline and block. The room gives every
// copy, and the save, what the editor alone gives. Typing over that prompt, its undo and redo,
// and a decision on the typing give what the editor alone gives too. A direct deletion that
// empties the control still shows the type's prompt, as the editor alone does.

import { afterEach, describe, expect, test } from 'bun:test';
import { createDocxEditor } from '../../../../editor/docx-editor.ts';
import { readOoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
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

const REV = 'w:id="91" w:author="Cara" w:date="2026-10-28T09:00:00Z"';
const BLANK = '<w:r><w:t xml:space="preserve">     </w:t></w:r>';
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const deleted = `<w:del ${REV}><w:r><w:delText>Acme</w:delText></w:r></w:del>`;
const inserted = `<w:ins ${REV}>${t('Acme')}</w:ins>`;

/** Word's own input shapes: an inline control in a line, a block control between two lines. */
const inline = (type: string, content: string) =>
  docx(p(t('Name: ') + sdt(`<w:id w:val="601"/>${type}`, content) + t(' end.')) + p(t('After.')));
const block = (type: string, content: string) =>
  docx(p(t('Before.')) + sdt(`<w:id w:val="602"/>${type}`, p(content)) + p(t('After.')));

/** The body XML without paragraph ids, as a reader of the saved file compares it. */
const bodyOf = (xml: string): string =>
  xml
    .slice(xml.indexOf('<w:body>'), xml.indexOf('</w:body>'))
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '');

/** The control's XML in a body. */
const controlOf = (xml: string): string => /<w:sdt>.*<\/w:sdt>/s.exec(xml)?.[0] ?? '<none>';

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

/** The editor alone (no room), the organiser editing: `act` on it, then its body. */
function alone(file: Uint8Array, act: (person: Person) => void): string {
  const container = document.createElement('div');
  document.body.append(container);
  const editor = createDocxEditor({
    container,
    document: file,
    author: 'Olive',
    mode: 'edit',
    modules: [reviewTestModule()],
  });
  try {
    const person = { name: 'alone', editor, container } as unknown as Person;
    act(person);
    return bodyOf(editorXml(person));
  } finally {
    editor.destroy();
    container.remove();
  }
}

/** The organiser decides the file's only revision card. */
function decide(person: Person, action: 'accept' | 'reject'): void {
  const [card] = person.editor.getReviewItems().filter((item) => item.kind === 'revision');
  if (card === undefined) throw new Error('no card');
  const result =
    action === 'accept'
      ? person.editor.acceptReviewItem(card.key)
      : person.editor.rejectReviewItem(card.key);
  expect(result.ok).toBe(true);
}

async function savedBody(person: Person): Promise<string> {
  const saved = new Uint8Array(await person.editor.save());
  expect(readOoxmlPackage(saved).ok).toBe(true);
  const again = openOrderedWordFile(saved, { file: 2 }).document;
  return bodyOf(serializeOoxmlPart(composeOrderedPart(again, BODY).part));
}

const wordInline = (type: string) =>
  `<w:sdt><w:sdtPr><w:id w:val="601"/><w:showingPlcHdr/>${type}</w:sdtPr><w:sdtEndPr/>` +
  `<w:sdtContent>${BLANK}</w:sdtContent></w:sdt>`;
const wordBlock =
  `<w:sdt><w:sdtPr><w:id w:val="602"/><w:showingPlcHdr/></w:sdtPr><w:sdtEndPr/>` +
  `<w:sdtContent><w:p>${BLANK}</w:p></w:sdtContent></w:sdt>`;

const CASES = [
  ['an inline plain text control', inline, '<w:text/>', wordInline('<w:text/>')],
  ['an inline untyped control', inline, '', wordInline('')],
  ['an inline rich text control', inline, '<w:richText/>', wordInline('')],
  ['a block untyped control', block, '', wordBlock],
  ['a block rich text control', block, '<w:richText/>', wordBlock],
] as const;

describe("a decision that empties a control with no prompt: Word's blank prompt", () => {
  for (const [name, shape, type, expected] of CASES) {
    for (const [action, content] of [
      ['accept', deleted],
      ['reject', inserted],
    ] as const) {
      test(`${name}: the organiser's ${action}, on every copy and in the save`, async () => {
        const file = shape(type, content);
        const editorAlone = alone(file, (person) => decide(person, action));
        expect(controlOf(editorAlone)).toBe(expected);
        const review = reviewTestModule();
        const room = localRoom(file);
        const olive = await join(room, 'Olive', {
          mode: 'edit',
          login: ORGANISER,
          wordFile: file,
          review,
        });
        const ana = await join(room, 'Ana', { wordFile: file, review });
        decide(olive, action);
        await settle();
        converged(room, olive, ana);
        expect(bodyOf(editorXml(olive))).toBe(editorAlone);
        expect(bodyOf(editorXml(ana))).toBe(editorAlone);
        expect(controlOf(await savedBody(olive))).toBe(expected);
      });
    }
  }

  test("a party's suggested deletion of the whole value, accepted: the blank prompt", async () => {
    const file = inline('<w:text/>', t('Acme'));
    const review = reviewTestModule();
    const room = localRoom(file);
    const olive = await join(room, 'Olive', {
      mode: 'edit',
      login: ORGANISER,
      wordFile: file,
      review,
    });
    const ana = await join(room, 'Ana', { wordFile: file, review });
    caretAt(ana, 0, 6, 10);
    ana.editor.surface?.deleteBackward();
    await settle();
    converged(room, olive, ana);
    decide(olive, 'accept');
    await settle();
    converged(room, olive, ana);
    expect(controlOf(bodyOf(editorXml(ana)))).toBe(wordInline('<w:text/>'));
    expect(controlOf(await savedBody(ana))).toBe(wordInline('<w:text/>'));
  });
});

/**
 * One step: the organiser's decision or direct deletion of the control's whole value, or the
 * typist's typing over the prompt, undo or redo.
 */
type Step = 'accept' | 'reject' | 'empty' | 'type' | 'undo' | 'redo';

/** A step only the organiser takes. */
const organiserStep = (what: Step): boolean =>
  what === 'accept' || what === 'reject' || what === 'empty';

/** The control without revision ids and dates, which the room and the editor alone mint apart. */
const stable = (xml: string): string => controlOf(xml).replace(/ w:(id|date)="[^"]*"/g, '');

const SHAPES = [
  ['inline', inline, 0, 6, 'Name: Beta end.'],
  ['block', block, 1, 0, 'Beta'],
] as const;

type Shape = (typeof SHAPES)[number];

/** Run a step on a person: the organiser decides the only card; the typist types or undoes. */
function step(person: Person, what: Step, [, , index, offset]: Shape): void {
  if (what === 'accept' || what === 'reject') decide(person, what);
  else if (what === 'empty') {
    caretAt(person, index, offset, offset + 4);
    surface(person).deleteBackward();
  } else if (what === 'type') {
    caretAt(person, index, offset);
    surface(person).insertPlainText('Beta');
  } else if (what === 'undo') surface(person).undo();
  else surface(person).redo();
}

/** The control after each step on the editor alone; a party's steps are suggestions by Ana. */
function aloneSteps(
  file: Uint8Array,
  steps: readonly Step[],
  who: 'party' | 'organiser',
  shape: Shape
) {
  const out: string[] = [];
  alone(file, (person) => {
    for (const what of steps) {
      const typist = !organiserStep(what);
      if (who === 'party') {
        person.editor.setAuthor(typist ? 'Ana' : 'Olive');
        if (typist) person.editor.setMode('suggesting');
      }
      step(person, what, shape);
      out.push(stable(bodyOf(editorXml(person))));
    }
  });
  return out;
}

// The room once refused typing into a control an edit or a decision emptied (`text-differs`):
// its prompt showed only in the drawing, with no units in the list. The emptying change now
// gives the control its prompt as units (`change/settle-emptied.ts`), so typing replaces it as
// in the editor alone, and a suggestion over it keeps it in the glossary. A decision leaves
// Word's blank prompt; a direct deletion the type's prompt, which its undo takes away again
const SOURCES = [
  [
    'the blank prompt a decision left',
    deleted,
    [
      ['party', ['accept', 'type', 'undo', 'redo', 'reject']],
      ['party', ['accept', 'type', 'accept']],
      ['organiser', ['accept', 'type', 'undo', 'redo']],
    ],
  ],
  [
    "the type's prompt a direct deletion left",
    t('Acme'),
    [
      ['party', ['empty', 'type', 'undo', 'redo', 'reject']],
      ['party', ['empty', 'type', 'accept']],
      ['organiser', ['empty', 'type', 'undo', 'redo']],
      ['organiser', ['empty', 'undo', 'redo']],
      ['organiser', ['empty', 'type', 'undo', 'undo', 'redo']],
    ],
  ],
] as const;

for (const [source, content, cases] of SOURCES)
  describe(`typing over ${source}, as the editor alone`, () => {
    for (const shape of SHAPES)
      for (const [who, steps] of cases)
        test(`${shape[0]}: the ${who} ${steps.slice(1).join(', ')}`, async () => {
          const file = shape[1]('<w:text/>', content);
          const expected = aloneSteps(file, steps, who, shape);
          const review = reviewTestModule();
          const room = localRoom(file);
          const olive = await join(room, 'Olive', {
            mode: 'edit',
            login: ORGANISER,
            wordFile: file,
            review,
          });
          const ana = await join(room, 'Ana', { wordFile: file, review });
          const shown: string[] = [];
          for (const what of steps) {
            const typist = !organiserStep(what) && who === 'party' ? ana : olive;
            step(typist, what, shape);
            await settle();
            converged(room, olive, ana);
            expect(bodyOf(editorXml(olive))).toBe(bodyOf(editorXml(ana)));
            if (what === 'type') expect(editorTexts(ana)[shape[2]]).toBe(shape[4]);
            shown.push(stable(bodyOf(editorXml(ana))));
          }
          expect(shown).toEqual(expected);
          expect(stable(await savedBody(ana))).toBe(expected.at(-1) ?? '');
        });
  });

describe('a direct deletion that empties a control: the type prompt, as the editor alone', () => {
  test("the organiser deletes the control's whole value", async () => {
    const file = inline('<w:text/>', t('Acme'));
    const remove = (person: Person) => {
      caretAt(person, 0, 6, 10);
      person.editor.surface?.deleteBackward();
    };
    const editorAlone = alone(file, remove);
    expect(controlOf(editorAlone)).toContain('Click or tap here to enter text.');
    const room = localRoom(file);
    const olive = await join(room, 'Olive', { mode: 'edit', login: ORGANISER, wordFile: file });
    const ana = await join(room, 'Ana', { wordFile: file });
    remove(olive);
    await settle();
    converged(room, olive, ana);
    expect(bodyOf(editorXml(ana))).toBe(editorAlone);
  });
});
