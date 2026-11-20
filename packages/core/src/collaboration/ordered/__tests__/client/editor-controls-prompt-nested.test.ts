// A control inside another control's prompt, with real editors and one ordered room. Ana's
// pending insertion fills the inner control. The organiser's Enter, typing, paste or field in
// the outer prompt takes the outer prompt's own words only: the inner control keeps its content,
// as the room keeps a control it put in a prompt. The room once read a file's inner control as
// part of the outer prompt, so the edit took Ana's insertion away with the prompt, with no
// decision: no card, and an empty control that a save and reopen gave a prompt. Each step is
// checked on every copy, in the save and reopen, and after one and two folds.

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { isOrderedItemHidden } from '../../card-of.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedCards } from '../../decide.ts';
import type { OrderedDocument } from '../../model/types.ts';
import { checkFold, folded, mediaStore } from '../fold-check.ts';
import { docx, p, t, W } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
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
const GLOSSARY = {
  parts: {
    'word/glossary/document.xml':
      `<w:glossaryDocument xmlns:w="${W}"><w:docParts><w:docPart><w:docPartPr>` +
      '<w:name w:val="Inner"/><w:category><w:name w:val="General"/>' +
      '<w:gallery w:val="placeholder"/></w:category></w:docPartPr>' +
      `<w:docPartBody><w:p>${t('Choose an item.', FACE)}</w:p></w:docPartBody></w:docPart>` +
      '</w:docParts></w:glossaryDocument>',
  },
  contentTypes:
    '<Override PartName="/word/glossary/document.xml" ContentType="application/' +
    'vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/>',
  relationships:
    '<Relationship Id="rIdGlossary" Type="http://schemas.openxmlformats.org/officeDocument/' +
    '2006/relationships/glossaryDocument" Target="glossary/document.xml"/>',
};

/** The inner control's type, by its properties. */
const TYPES = [
  ['plain text', '<w:text/>'],
  ['rich text', ''],
  ['combo box', '<w:comboBox><w:listItem w:displayText="data" w:value="data"/></w:comboBox>'],
  [
    'drop-down',
    '<w:dropDownList><w:listItem w:displayText="data" w:value="data"/></w:dropDownList>',
  ],
  ['date', '<w:date><w:dateFormat w:val="M/d/yyyy"/><w:lid w:val="en-US"/></w:date>'],
] as const;

/**
 * A date picker showing its prompt, as a party left it: an inner control in the prompt, which
 * names the glossary's prompt and holds Ana's pending insertion.
 */
const fileWith = (inner: string): Uint8Array =>
  docx(
    p(t('Due:')) +
      '<w:sdt><w:sdtPr><w:id w:val="601"/><w:showingPlcHdr/><w:date>' +
      '<w:dateFormat w:val="MMMM d, yyyy"/></w:date></w:sdtPr><w:sdtContent><w:p>' +
      t('Click to select a date..') +
      '<w:sdt><w:sdtPr><w:id w:val="602"/><w:placeholder><w:docPart w:val="Inner"/>' +
      `</w:placeholder>${inner}</w:sdtPr><w:sdtContent>` +
      '<w:ins w:id="7" w:author="Ana" w:date="2026-10-28T08:00:00Z">' +
      `${t('data ')}</w:ins></w:sdtContent></w:sdt>${t('.')}</w:p></w:sdtContent></w:sdt>` +
      p(t('Signed.')),
    GLOSSARY
  );

/** An edit at the caret in the outer prompt; a field goes in through the editing API's lane. */
const EDITS = [
  ['Enter', (person: Person) => surface(person).splitParagraph()],
  ['typing', (person: Person) => surface(person).insertPlainText('Q')],
  ['a paste', (person: Person) => surface(person).pasteRich('Q', '<p>Q</p>')],
  [
    'a field',
    (person: Person) =>
      surface(person).applyAutomationOps(() => [
        {
          op: 'insertPageField',
          paragraphId: surface(person).session.paragraphIds()[1] ?? '',
          offset: 22,
          field: 'PAGE',
        },
      ]),
  ],
] as const;

/** The inner control as a reader of the file sees it: no ids, paragraph ids or dates. */
const innerOf = (xml: string): string =>
  (/<w:sdt><w:sdtPr><w:id w:val="602"\/>.*?<\/w:sdt>/s.exec(xml)?.[0] ?? '<none>')
    .replace(/ w14:(paraId|textId)="[^"]*"/g, '')
    .replace(/ w:(date|id)="[^"]*"/g, '');

const composed = (doc: OrderedDocument): string =>
  innerOf(serializeOoxmlPart(composeOrderedPart(doc, BODY).part));

/** The inner control after a save and reopen of the room's document. */
async function reopened(person: Person): Promise<string> {
  const saved = new Uint8Array(await person.editor.save());
  return composed(openOrderedWordFile(saved, { file: 2 }).document);
}

/** Whether Ana's insertion shows: her units, and her card. */
function anasShown(doc: OrderedDocument): { units: boolean; cards: number } {
  const ana = [...doc.suggestions.values()].filter((s) => s.attribution.author === 'Ana');
  const units = ana.flatMap((s) =>
    [...(doc.labelled.get(s.id) ?? [])].filter((i) => i.by === s.id)
  );
  return {
    units: units.length > 0 && units.every((item) => !isOrderedItemHidden(doc, item)),
    cards: orderedCards(doc).filter((card) => card.author === 'Ana').length,
  };
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

/** A room with the organiser and Bob, a party, in their editors; Ana's insertion is checked. */
async function seats(file: Uint8Array) {
  const review = reviewTestModule();
  const room = localRoom(file);
  const olive = await join(room, 'Olive', {
    mode: 'edit',
    login: ORGANISER,
    wordFile: file,
    review,
  });
  const bob = await join(room, 'Bob', { wordFile: file, review });
  const filled = composed(room.room.document);
  expect(filled).toContain('<w:ins w:author="Ana"><w:r><w:t xml:space="preserve">data ');
  expect(anasShown(room.room.document)).toEqual({ units: true, cards: 1 });
  return { room, olive, bob, filled };
}

/**
 * One person's edit in the outer prompt: it replaces the prompt, and the inner control keeps
 * Ana's insertion and its card on every copy, after a save and reopen, and after one and two
 * folds.
 */
async function editKeeps(
  file: Uint8Array,
  edit: string,
  act: (person: Person) => unknown,
  who: 'Olive' | 'Bob'
) {
  const seated = await seats(file);
  const { room, olive, bob, filled } = seated;
  const doc = room.room.document;
  const person = who === 'Olive' ? olive : bob;
  caretAt(person, 1, 22);
  act(person);
  await settle();
  converged(room, olive, bob);
  expect(editorXml(bob)).not.toContain('Click to select');
  expect(innerOf(editorXml(bob))).toBe(filled);
  expect(innerOf(editorXml(olive))).toBe(filled);
  expect(composed(doc)).toBe(filled);
  expect(await reopened(bob)).toBe(filled);
  expect(anasShown(doc)).toEqual({ units: true, cards: 1 });
  const { base } = openOrderedWordFile(file, { file: 1 });
  let copy = { document: doc, base };
  for (const fold of [2, 3]) {
    copy = await folded(copy.document, copy.base, fold, mediaStore(base));
    expect({ fold, shown: anasShown(copy.document) }).toEqual({
      fold,
      shown: { units: true, cards: 1 },
    });
  }

  return { ...seated, base };
}

describe("an edit in a prompt keeps a control inside it, with a party's insertion", () => {
  for (const [type, inner] of TYPES) {
    for (const [edit, act] of EDITS) {
      test(`the organiser's ${edit} in the prompt around a ${type} control`, async () => {
        const { room, olive, bob, base } = await editKeeps(fileWith(inner), edit, act, 'Olive');
        await checkFold(room.room.document, base, mediaStore(base), `${edit}, ${type}`);

        // The organiser rejects Ana's insertion: the emptied control shows its glossary prompt,
        // the same after a save and reopen
        const [card] = olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
        expect(olive.editor.rejectReviewItem(card?.key ?? '').ok).toBe(true);
        await settle();
        converged(room, olive, bob);
        const emptied = composed(room.room.document);
        expect(emptied).toContain('<w:showingPlcHdr/>');
        expect(emptied).toContain('Choose an item.');
        expect(innerOf(editorXml(bob))).toBe(emptied);
        expect(await reopened(olive)).toBe(emptied);
      });
    }
  }

  for (const [edit, act] of EDITS) {
    test(`a party's suggested ${edit} in the prompt around a combo box control`, async () => {
      await editKeeps(fileWith(TYPES[2][1]), edit, act, 'Bob');
    });
  }
});
