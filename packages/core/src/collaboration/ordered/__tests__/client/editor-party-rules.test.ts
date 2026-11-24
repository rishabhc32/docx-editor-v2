// A party's edit in an ordered room gets the verdict the host's earlier check gives (it
// takes back the party's own suggestions from the Word file before and after the change, and
// compares the two views), with real editors, two
// parties and the organiser. Where the check allows an edit the room refused before, the edit
// is the party's own tracked suggestion, saved, rejected and accepted. Where it refuses one the
// room accepted before, the party's editor takes it back and the room refuses it by name. The
// differences the room keeps are the check's own faults of Word form (design D9).

import { afterEach, describe, expect, test } from 'bun:test';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPart } from '../../compose/compose.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorTexts,
  join,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  screenFollowsList,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { caretIn, TABLE_FILE, tableCommand } from './tables-harness.ts';

afterEach(closeAll);

const review = reviewTestModule();

const FILE = docx(p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')));

const PROMPT = 'Click or tap here to enter text.';
/** A block control that shows its prompt, between two lines. */
const PROMPTED = docx(
  p(t('Name:')) +
    '<w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
    p(t(PROMPT, '<w:rStyle w:val="PlaceholderText"/>')) +
    '</w:sdtContent></w:sdt>' +
    p(t('After.'))
);

/** An inline control that shows its prompt, between words. */
const INLINE_PROMPTED = docx(
  p(
    t('Name: ') +
      '<w:sdt><w:sdtPr><w:showingPlcHdr/></w:sdtPr><w:sdtContent>' +
      t(PROMPT, '<w:rStyle w:val="PlaceholderText"/>') +
      '</w:sdtContent></w:sdt>' +
      t(' end.')
  )
);

/** Two parties and the organiser, in a room with the suggestions-only policy. */
async function people(file: Uint8Array) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const ana = await join(room, 'Ana', { wordFile: file, review });
  const ben = await join(room, 'Ben', { wordFile: file, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  return { room, ana, ben, olive };
}

/** Every copy is the room's, nothing is refused, and every editor shows its list. */
function converged(room: OrderedLocalRoom, ...everyone: Person[]): void {
  for (const person of everyone) {
    expect({ who: person.name, refusals: person.refusals }).toEqual({
      who: person.name,
      refusals: [],
    });
    expect({ who: person.name, room: sameAsRoom(person, room) }).toEqual({
      who: person.name,
      room: true,
    });
    expect({ who: person.name, screen: screenFollowsList(person) }).toEqual({
      who: person.name,
      screen: true,
    });
  }
}

/** The person's refusals by reason, then none. */
function refused(person: Person): string[] {
  const reasons = person.refusals.map((r) => `${r.code}/${r.reason ?? ''}`);
  person.refusals.length = 0;
  return reasons;
}

const run = (person: Person, command: Parameters<Person['editor']['exec']>[0]) =>
  expect(person.editor.exec(command).ok).toBe(true);

/** The cards the room shows, each as its suggestions' authors. */
const cardsOf = (room: OrderedLocalRoom): string[][] => {
  const doc = room.room.document;
  return orderedCards(doc).map((card) =>
    card.suggestions.map((id) => doc.suggestions.get(id)?.attribution.author ?? '?')
  );
};

/** The organiser decides the card that holds a suggestion of `author`. */
async function decide(
  room: OrderedLocalRoom,
  olive: Person,
  author: string,
  status: 'accepted' | 'rejected'
) {
  const doc = room.room.document;
  const card = orderedCards(doc).find((c) =>
    c.suggestions.some((id) => doc.suggestions.get(id)?.attribution.author === author)
  );
  if (!card) throw new Error(`no card of ${author}`);
  olive.handle.session.requestDecision(card.id, status);
  await settle();
}

/** The room's body as the save writes it, and the body of the saved file opened again. */
async function savedBody(olive: Person): Promise<string> {
  const bytes = new Uint8Array(await olive.editor.save());
  const reopened = openOrderedWordFile(bytes, { file: 9 });
  return serializeOoxmlPart(composeOrderedPart(reopened.document, BODY).part);
}

const revisions = (xml: string): string[] =>
  [
    ...xml.matchAll(
      /<w:(ins|del|rPrChange|pPrChange|tcPrChange|trPrChange) [^>]*w:author="(\w+)"/g
    ),
  ].map((m) => `${m[1]}:${m[2]}`);

describe('a party edit today’s check allows', () => {
  // Today: Ben's `w:ins` between Ana's `w:del` and `w:ins` is taken back alone and leaves the
  // file as it was, so the check allows it. It reads no cards
  test('words between another person’s deletion and insertion: the party’s own card', async () => {
    const { room, ana, ben, olive } = await people(FILE);
    caretAt(ana, 0, 11, 14);
    run(ana, { type: 'insertText', text: '200' });
    await settle();
    // One suggestion: her deletion and her insertion
    expect(cardsOf(room)).toEqual([['Ana']]);
    // In the editor the struck "100" shows, so "200" starts at 14
    caretAt(ben, 0, 14);
    run(ben, { type: 'insertText', text: 'x' });
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ana'], ['Ben']]);
    expect(revisions(await savedBody(olive))).toEqual(['del:Ana', 'ins:Ben', 'ins:Ana']);
    // Rejected, Ben's words go and Ana's replacement stays whole
    await decide(room, olive, 'Ben', 'rejected');
    converged(room, ana, ben, olive);
    expect(revisions(await savedBody(olive))).toEqual(['del:Ana', 'ins:Ana']);
    caretAt(ben, 0, 14);
    run(ben, { type: 'insertText', text: 'y' });
    await settle();
    // Accepted, the words stay as plain text between Ana's two
    await decide(room, olive, 'Ben', 'accepted');
    converged(room, ana, ben, olive);
    expect(revisions(await savedBody(olive))).toEqual(['del:Ana', 'ins:Ana']);
    expect(editorTexts(olive)[0]).toBe('The fee is 100y200 USD.');
  });

  // Today: the editor types at a link's start beside it, as the editor alone does, so Ana's
  // `w:ins` is next to Ben's new link and taking it back leaves his link as it was (r7
  // quick-short:273). Between his words it is inside his insertion
  test('typing at the start of another person’s new link: beside it, the party’s own card', async () => {
    const { room, ana, ben, olive } = await people(FILE);
    caretAt(ben, 0, 4);
    run(ben, { type: 'insertHyperlink', href: 'https://example.com/', text: 'abc' });
    await settle();
    caretAt(ana, 0, 5);
    surface(ana).type('Y');
    await settle();
    expect(refused(ana)).toEqual(['ordered-refused/nested']);
    caretAt(ana, 0, 4);
    surface(ana).type('X');
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ana'], ['Ben']]);
    const body = await savedBody(olive);
    expect(revisions(body)).toEqual(['ins:Ana', 'ins:Ben']);
    expect(body.indexOf('w:author="Ana"')).toBeLessThan(body.indexOf('<w:hyperlink'));

    // Rejected, her words go and his link stays whole
    await decide(room, olive, 'Ana', 'rejected');
    converged(room, ana, ben, olive);
    expect(revisions(await savedBody(olive))).toEqual(['ins:Ben']);
    expect(editorTexts(olive)[0]).toBe('The abcfee is 100 USD.');
  });

  // Today: Enter in a link writes two `w:hyperlink` elements and marks only the new paragraph
  // mark, so Ana's `w:ins` in the part it split off is in the file's link, in no one's insertion
  test('typing in the part of a link that another person’s Enter split off: the party’s own card', async () => {
    const file = docx(p(`<w:hyperlink w:anchor="_Toc1">${t('Form Elements')}</w:hyperlink>`));
    const { room, ana, ben, olive } = await people(file);
    caretAt(ben, 0, 4);
    surface(ben).splitParagraph();
    await settle();
    caretAt(ana, 1, 4);
    surface(ana).type('X');
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ben'], ['Ana']]);
    expect(editorTexts(olive)).toEqual(['Form', ' EleXments']);
    expect(revisions(await savedBody(olive))).toEqual(['ins:Ben', 'ins:Ana']);

    // Her undo takes her words back alone, and his Enter stays
    surface(ana).undo();
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ben']]);
    expect(editorTexts(olive)).toEqual(['Form', ' Elements']);
    surface(ana).redo();
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ben'], ['Ana']]);
    expect(revisions(await savedBody(olive))).toEqual(['ins:Ben', 'ins:Ana']);

    // His undo takes his Enter back alone: one link again, her words in it and hers
    surface(ben).undo();
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ana']]);
    expect(editorTexts(olive)).toEqual(['Form EleXments']);
    const body = await savedBody(olive);
    expect(revisions(body)).toEqual(['ins:Ana']);
    expect(body.split('<w:hyperlink ').length).toBe(2);

    // Rejected, her words go and the link stays whole
    await decide(room, olive, 'Ana', 'rejected');
    converged(room, ana, ben, olive);
    expect(revisions(await savedBody(olive))).toEqual([]);
    expect(editorTexts(olive)).toEqual(['Form Elements']);
  });

  // Today: Word writes Ben's `w:pPrChange` beside Ana's `w:ins` on the paragraph mark, and the
  // take-back of Ben's record alone leaves the file as it was (allowed). It joins Ana's card
  test('a paragraph property on another person’s new paragraph: decided with it', async () => {
    for (const status of ['rejected', 'accepted'] as const) {
      const { room, ana, ben, olive } = await people(FILE);
      caretAt(ana, 0, 7);
      surface(ana).splitParagraph();
      await settle();
      caretAt(ben, 0, 2);
      run(ben, { type: 'setAlignment', align: 'center' });
      await settle();
      converged(room, ana, ben, olive);
      expect(cardsOf(room)).toEqual([['Ana', 'Ben']]);
      expect(revisions(await savedBody(olive))).toEqual(['ins:Ana', 'pPrChange:Ben']);
      await decide(room, olive, 'Ana', status);
      converged(room, ana, ben, olive);
      const body = await savedBody(olive);
      expect(revisions(body)).toEqual([]);
      expect(editorTexts(olive)).toEqual(
        status === 'rejected'
          ? ['The fee is 100 USD.', 'Due in 30 days.']
          : ['The fee', ' is 100 USD.', 'Due in 30 days.']
      );
      expect(body.includes('<w:jc w:val="center"/>')).toBe(status === 'accepted');
      closeAll();
    }
  });

  // Today: the deletion of a prompt changes no word of the file (the prompt is no content), so
  // the check allows the party's direct edit; one that changes words gives `text`
  test('a party’s direct edit that changes nothing is allowed; one that changes words is not', async () => {
    const room = localRoom(PROMPTED, { policy: suggestionsOnlyPolicy() });
    const ben = await join(room, 'Ben', { wordFile: PROMPTED, mode: 'edit', login: PARTY, review });
    caretAt(ben, 1, 0, 5);
    run(ben, { type: 'deleteText' });
    await settle();
    expect(refused(ben)).toEqual([]);
    expect(editorTexts(ben)).toEqual(['Name:', PROMPT, 'After.']);
    expect(room.room.document.suggestions.size).toBe(0);
    caretAt(ben, 0, 5);
    run(ben, { type: 'insertText', text: ' Ben' });
    await settle();
    expect(refused(ben)).toEqual(['ordered-refused/text']);
    expect(editorTexts(ben)).toEqual(['Name:', PROMPT, 'After.']);
  });

  // Today: the engine's new row copies the row's shown properties with no change record
  // (`insertTableRow` strips them), as the editor alone does; the take-back of Ben's row leaves
  // the file as it was (allowed). Word not checked
  test('a new row beside a row with another person’s pending cell format: the party’s own row', async () => {
    for (const status of ['rejected', 'accepted'] as const) {
      const room = localRoom(TABLE_FILE, { policy: suggestionsOnlyPolicy() });
      const ana = await join(room, 'Ana', { wordFile: TABLE_FILE, review });
      const ben = await join(room, 'Ben', { wordFile: TABLE_FILE, review });
      const olive = await join(room, 'Olive', {
        wordFile: TABLE_FILE,
        mode: 'edit',
        login: ORGANISER,
        review,
      });
      caretIn(ana, 'Fee');
      expect(
        tableCommand(ana, { type: 'setCellFill', color: { kind: 'hex', value: 'FFCC00' } }).ok
      ).toBe(true);
      await settle();
      caretIn(ben, 'Fee');
      expect(tableCommand(ben, { type: 'insertRow', where: 'below' }).ok).toBe(true);
      await settle();
      converged(room, ana, ben, olive);
      const rows = (xml: string) => xml.match(/<w:tr[ >].*?<\/w:tr>/gs) ?? [];
      const [, added] = rows(await savedBody(olive));
      // Ben's row, with no record of Ana's
      expect(revisions(added ?? '')).toEqual(['ins:Ben', 'ins:Ben', 'ins:Ben', 'ins:Ben']);
      await decide(room, olive, 'Ben', status);
      await decide(room, olive, 'Ana', status);
      converged(room, ana, ben, olive);
      const body = await savedBody(olive);
      expect(revisions(body)).toEqual([]);
      expect(rows(body).length).toBe(status === 'accepted' ? 3 : 2);
      closeAll();
    }
  });

  // Today: each new row has its own `w:ins`, and the take-back of Ana's rows leaves Ben's row as
  // it was (allowed)
  test('new rows right above and below another person’s new row: the party’s own rows', async () => {
    for (const status of ['rejected', 'accepted'] as const) {
      const { room, ana, ben, olive } = await people(TABLE_FILE);
      caretIn(ben, 'Fee');
      expect(tableCommand(ben, { type: 'insertRow', where: 'below' }).ok).toBe(true);
      await settle();
      caretIn(ana, 'Fee');
      expect(tableCommand(ana, { type: 'insertRow', where: 'below' }).ok).toBe(true);
      await settle();
      caretIn(ana, '100');
      expect(tableCommand(ana, { type: 'insertRow', where: 'above' }).ok).toBe(true);
      await settle();
      converged(room, ana, ben, olive);
      const rows = (xml: string) => xml.match(/<w:tr[ >].*?<\/w:tr>/gs) ?? [];
      const authors = (xml: string) => rows(xml).map((row) => revisions(row)[0] ?? 'file');
      expect(authors(await savedBody(olive))).toEqual([
        'file',
        'ins:Ana',
        'ins:Ben',
        'ins:Ana',
        'file',
      ]);
      while (cardsOf(room).length > 0) {
        await decide(room, olive, cardsOf(room)[0]![0]!, status);
      }

      converged(room, ana, ben, olive);
      expect(rows(await savedBody(olive)).length).toBe(status === 'accepted' ? 5 : 2);
      closeAll();
    }
  });

  // Today: a deletion leaves a prompt's units (a prompt is not content), so Ben's format on the
  // prompt stays as it was (allowed)
  test('a deletion over a prompt with another person’s pending format: the prompt stays', async () => {
    const { room, ana, ben, olive } = await people(INLINE_PROMPTED);
    caretAt(ben, 0, 12, 18);
    run(ben, { type: 'toggleMark', mark: 'bold' });
    await settle();
    caretAt(ana, 0, 2, 6 + PROMPT.length);
    run(ana, { type: 'deleteText' });
    await settle();
    converged(room, ana, ben, olive);
    expect(editorTexts(ana)).toEqual([`Name: ${PROMPT} end.`]);
    expect(cardsOf(room)).toEqual([['Ana'], ['Ben']]);
    expect(revisions(await savedBody(olive))).toEqual(['del:Ana', 'rPrChange:Ben']);
  });
});

describe('a party edit today’s check refuses', () => {
  // Today: Word keeps one `w:rPrChange` per run, so Ana's format record goes under Ben's name,
  // and her suggestion is gone from today's view (`others_suggestions`)
  test('a format over another person’s pending format: refused by name, the organiser’s replaces it', async () => {
    const { room, ana, ben, olive } = await people(FILE);
    caretAt(ana, 0, 4, 7);
    run(ana, { type: 'toggleMark', mark: 'italic' });
    await settle();
    const before = serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part);
    caretAt(ben, 0, 4, 7);
    run(ben, { type: 'toggleMark', mark: 'bold' });
    await settle();
    expect(refused(ben)).toEqual(['ordered-refused/others_suggestions']);
    converged(room, ana, ben, olive);
    expect(serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part)).toBe(before);
    caretAt(olive, 0, 4, 7);
    run(olive, { type: 'toggleMark', mark: 'bold' });
    await settle();
    converged(room, ana, ben, olive);
  });

  // Today: Word nests Ben's `w:rPrChange` in the runs of Ana's `w:del` and `w:ins`, and the
  // check cannot take it back alone (`nested`), on part or all of her words. A colour her words
  // have already writes no record: allowed
  test('a format over another person’s struck and inserted words: refused by name', async () => {
    const red = docx(p(t('The fee is ') + t('100', '<w:color w:val="C00000"/>') + t(' USD.')));
    const { room, ana, ben, olive } = await people(red);
    caretAt(ana, 0, 11, 14);
    run(ana, { type: 'insertText', text: '200' });
    await settle();
    const before = serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part);
    for (const [from, to] of [
      [11, 17],
      [12, 13],
      [15, 16],
    ]) {
      caretAt(ben, 0, from, to);
      run(ben, { type: 'toggleMark', mark: 'bold' });
      await settle();
      expect(refused(ben)).toEqual(['ordered-refused/nested']);
      converged(room, ana, ben, olive);
      expect(serializeOoxmlPart(composeOrderedPart(room.room.document, BODY).part)).toBe(before);
    }

    caretAt(ben, 0, 11, 14);
    run(ben, { type: 'setMarkAttr', mark: 'color', attr: 'val', value: 'C00000' });
    await settle();
    converged(room, ana, ben, olive);
    expect(cardsOf(room)).toEqual([['Ana']]);
    expect(revisions(await savedBody(olive))).toEqual(['del:Ana', 'ins:Ana']);
  });

  // Today: Word nests Ana's `w:rPrChange` in Ben's deleted run, so taking the deletion back takes
  // her format back too (`package`)
  test('a deletion over another person’s pending format: refused by name', async () => {
    const { room, ana, ben, olive } = await people(FILE);
    caretAt(ana, 0, 4, 7);
    run(ana, { type: 'toggleMark', mark: 'italic' });
    await settle();
    caretAt(ben, 0, 4, 10);
    run(ben, { type: 'deleteText' });
    await settle();
    expect(refused(ben)).toEqual(['ordered-refused/others_suggestions']);
    converged(room, ana, ben, olive);
    // Beside her format, the deletion is Ben's suggestion
    caretAt(ben, 0, 11, 14);
    run(ben, { type: 'deleteText' });
    await settle();
    converged(room, ana, ben, olive);
  });

  // A direct insertion inside a pending insertion goes into it, as Word extends a tracked
  // insertion with untracked typing: Cy's words would go under Ana's name, with no suggestion of
  // Cy's for the room to see (design D9). Refused as direct typing in the file's own text is
  test('a party’s direct typing inside another person’s insertion: refused, the organiser’s joins it', async () => {
    const { room, ana, ben, olive } = await people(FILE);
    const cy = await join(room, 'Cy', { wordFile: FILE, mode: 'edit', login: PARTY, review });
    caretAt(ana, 0, 4);
    run(ana, { type: 'insertText', text: 'abc' });
    await settle();
    caretAt(cy, 0, 5);
    run(cy, { type: 'insertText', text: 'X' });
    await settle();
    expect(refused(cy)).toEqual(['ordered-refused/text']);
    converged(room, ana, ben, cy, olive);
    expect(editorTexts(cy)[0]).toBe('The abcfee is 100 USD.');
    caretAt(olive, 0, 5);
    run(olive, { type: 'insertText', text: 'O' });
    await settle();
    converged(room, ana, ben, cy, olive);
    expect(cardsOf(room)).toEqual([['Ana']]);
    expect(revisions(await savedBody(olive))).toEqual(['ins:Ana']);
    expect(editorTexts(cy)[0]).toBe('The aObcfee is 100 USD.');
  });

  // Today: Enter removes the prompt for good (policy 4), so the take-back of the break leaves the
  // control empty and today's view loses the prompt's words (`text`)
  test('Enter in a block control’s prompt: refused for a party, the organiser’s goes', async () => {
    const room = localRoom(PROMPTED, { policy: suggestionsOnlyPolicy() });
    const ben = await join(room, 'Ben', { wordFile: PROMPTED, review });
    const olive = await join(room, 'Olive', {
      wordFile: PROMPTED,
      mode: 'edit',
      login: ORGANISER,
      review,
    });
    caretAt(ben, 1, 5);
    surface(ben).splitParagraph();
    await settle();
    expect(refused(ben)).toEqual(['ordered-refused/text']);
    converged(room, ben, olive);
    expect(editorTexts(ben)).toEqual(['Name:', PROMPT, 'After.']);
    caretAt(olive, 1, 5);
    surface(olive).splitParagraph();
    await settle();
    converged(room, ben, olive);
    expect(editorTexts(olive)).toEqual(['Name:', '', '', 'After.']);
  });
});
