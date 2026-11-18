// Typing over a control's prompt that carries another person's pending suggestion, with real
// editors and one ordered room (task 9.6; acceptance quick pass 101, cy-1:78). Typed words replace
// the prompt's runs, as Word drops them, and the save keeps the prompt as words only. A party's
// typing there is refused as inside the other person's suggestion, as today's check refuses it
// (`others_suggestions`), so every copy and the saved file keep that suggestion. The organiser
// types over it as the editor alone does: the prompt goes, with the suggestion on it, and undo
// brings both back.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-serialize.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { mount } from '../matrix/engine.ts';
import { docx, p, t } from '../word-fixtures.ts';
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

afterEach(closeAll);

const PROMPT = 'Click or tap here to enter text.';
const FACE = '<w:rStyle w:val="PlaceholderText"/>';
/** Bob's pending bold on the prompt's words " here to ". */
const BOBS =
  `${FACE}<w:b/><w:rPrChange w:id="7" w:author="Bob" w:date="2026-10-28T08:00:00Z">` +
  `<w:rPr>${FACE}</w:rPr></w:rPrChange>`;
const sdt = (props: string, content: string) =>
  `<w:sdt><w:sdtPr>${props}</w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
const promptRuns = (bob: boolean) =>
  t('Click or tap', FACE) + t(' here to ', bob ? BOBS : FACE) + t('enter text.', FACE);

const inline = (bob: boolean) =>
  docx(
    p(t('Name: ') + sdt('<w:showingPlcHdr/><w:text/>', promptRuns(bob)) + t(' end.')) +
      p(t('Signed here.'))
  );
const block = (bob: boolean) =>
  docx(p(t('Name:')) + sdt('<w:showingPlcHdr/>', p(promptRuns(bob))) + p(t('Signed here.')));

const CASES = [
  { name: 'inline', file: inline, line: 0, starts: 6, shown: `Name: ${PROMPT} end.` },
  { name: 'block', file: block, line: 1, starts: 0, shown: PROMPT },
] as const;

const bobsCards = (doc: OrderedDocument) => orderedCards(doc).filter((c) => c.author === 'Bob');

async function reopenedCards(person: Person): Promise<number> {
  const saved = new Uint8Array(await person.editor.save());
  return bobsCards(openOrderedWordFile(saved, { file: 2 }).document).length;
}

function converged(room: ReturnType<typeof localRoom>, ...people: Person[]): void {
  for (const person of people) {
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

/** The control the editor alone writes after typing "Q" directly at `offset` of paragraph `line`. */
function aloneControl(bytes: Uint8Array, line: number, offset: number): string {
  const alone = mount(bytes, 'Olga');
  try {
    alone.surface.setEditingMode('edit');
    const paragraphId = alone.surface.session.paragraphIds()[line];
    if (paragraphId === undefined) throw new Error(`no paragraph ${line}`);
    alone.surface.setSelection({ anchor: { paragraphId, offset }, head: { paragraphId, offset } });
    alone.surface.insertPlainText('Q');
    return controlXml(serializeOoxmlPart(alone.surface.session.part()));
  } finally {
    alone.close();
  }
}

/** The control as the editor writes it, ids and dates aside. */
const controlXml = (xml: string) =>
  (/<w:sdt>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? '')
    .replace(/<w:docPart w:val="[^"]*"\/>/g, '<w:docPart/>')
    .replace(/ w:(id|date)="[^"]*"/g, '');

describe("a party's typing over a prompt that carries another person's pending format", () => {
  for (const { name, file, line, starts, shown } of CASES)
    for (const [where, at] of [
      ['at its start', starts],
      ["in the other person's words", starts + 15],
    ] as const)
      test(`${name} control, ${where}: refused as nested, and every copy keeps the format`, async () => {
        const room = localRoom(file(true), { policy: suggestionsOnlyPolicy() });
        const cy = await join(room, 'Cy', { wordFile: file(true) });
        const olga = await join(room, 'Olga', {
          wordFile: file(true),
          mode: 'edit',
          login: ORGANISER,
        });
        expect(editorTexts(cy)[line]).toBe(shown);
        caretAt(cy, line, at);
        surface(cy).insertPlainText('Q');
        await settle();
        expect(cy.refusals.map((r) => `${r.code}:${r.reason}`)).toEqual(['ordered-refused:nested']);
        converged(room, cy, olga);
        expect(editorTexts(cy)[line]).toBe(shown);
        expect(bobsCards(room.room.document)).toHaveLength(1);
        expect(await reopenedCards(cy)).toBe(1);
        expect(await reopenedCards(olga)).toBe(1);
      });

  test('without a suggestion of another person on the prompt, the same typing is a suggestion', async () => {
    for (const { file, line, starts } of CASES) {
      const room = localRoom(file(false), { policy: suggestionsOnlyPolicy() });
      const cy = await join(room, 'Cy', { wordFile: file(false) });
      caretAt(cy, line, starts);
      surface(cy).insertPlainText('Q');
      await settle();
      expect(cy.refusals).toEqual([]);
      converged(room, cy);
      expect(editorXml(cy)).toMatch(/<w:ins [^>]*w:author="Cy"/);
    }
  });
});

describe('the organiser types over it as the editor alone does', () => {
  for (const { name, file, line, starts, shown } of CASES)
    test(`${name} control: the prompt goes with the format, and undo brings both back`, async () => {
      const room = localRoom(file(true), { policy: suggestionsOnlyPolicy() });
      const olga = await join(room, 'Olga', {
        wordFile: file(true),
        mode: 'edit',
        login: ORGANISER,
      });
      const cy = await join(room, 'Cy', { wordFile: file(true) });
      // In Bob's words: the typed text takes the prompt's first words' face, not Bob's bold
      caretAt(olga, line, starts + 15);
      surface(olga).insertPlainText('Q');
      await settle();
      expect(olga.refusals).toEqual([]);
      converged(room, olga, cy);
      expect(editorTexts(cy)[line]).toBe(shown.replace(PROMPT, 'Q'));
      expect(controlXml(editorXml(olga))).toBe(aloneControl(file(true), line, starts + 15));
      expect(bobsCards(room.room.document)).toHaveLength(0);
      expect(await reopenedCards(olga)).toBe(0);
      surface(olga).undo();
      await settle();
      converged(room, olga, cy);
      expect(editorTexts(cy)[line]).toBe(shown);
      expect(bobsCards(room.room.document)).toHaveLength(1);
      expect(await reopenedCards(cy)).toBe(1);
    });
});
