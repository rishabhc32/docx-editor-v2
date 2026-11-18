// Notes through real editors and one room (group 9.9, task 9.11): a party's new note in a file
// with no notes part, typed in, rejected, accepted and saved; the organiser's deletion,
// conversion and note properties, each undone and made again; a party refused by the host's party
// rule; two new notes at once, alike on every copy; a deletion while another person types in
// the note; a party's take-back; a party's new note and deleted note, each one card that
// decides the note's text with its reference, as Word writes them; and a party's deletion over
// its own new reference, which takes the note too. After each step every copy is the room's,
// and every editor shows what its list composes.

import { afterEach, describe, expect, test } from 'bun:test';
import { readOoxmlPackage, type OoxmlPackage } from '../../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../../store/package/ooxml-tree.ts';
import { settingsPartOf } from '../../../../store/package/note-properties.ts';
import { orderedCardOf } from '../../card-of.ts';
import { orderedCardSeen } from '../../change/apply-decide.ts';
import type { OrderedLocalRoom } from '../../client/local-room.ts';
import { openOrderedWordFile } from '../../client/open.ts';
import { composeOrderedPackage } from '../../compose/package.ts';
import { orderedCards } from '../../decide.ts';
import { suggestionsOnlyPolicy } from '../../room/suggestions-only.ts';
import { docx, MATRIX_EXTRAS, p, t } from '../word-fixtures.ts';
import {
  BODY,
  caretAt,
  closeAll,
  editorXml,
  FILE as PLAIN,
  HeldTransport,
  join,
  listXml,
  localRoom,
  ORGANISER,
  PARTY,
  sameAsRoom,
  settle,
  surface,
  type Person,
} from './editor-harness.ts';
import { reviewTestModule } from './review-module.ts';
import { tableCommand } from './tables-harness.ts';
import { testPerson } from '../test-persons.ts';

afterEach(closeAll);

const review = reviewTestModule();

const ENDNOTES = '/word/endnotes.xml';
const REFERENCE =
  '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>';

/** A footnote the file holds (" Fees are due net of tax."), an endnote, and one section. */
const NOTED = docx(
  p(t('The fee') + REFERENCE + t(' is due.')) +
    p(t('Paid in 30 days.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  MATRIX_EXTRAS
);

/** A date picker after a label, in a file with a notes part. */
const DATED = docx(
  p(
    t('Start date: ') +
      '<w:sdt><w:sdtPr><w:date w:fullDate="2026-01-02T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/>' +
      '<w:lid w:val="en-GB"/></w:date></w:sdtPr><w:sdtContent>' +
      t('2 January 2026') +
      '</w:sdtContent></w:sdt>'
  ) + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  MATRIX_EXTRAS
);

const cellOf = (runs: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(runs)}</w:tc>`;

/** A table whose first row holds the file's footnote reference. */
const ROWED = docx(
  p(t('The fees are below.')) +
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
    `<w:tr>${cellOf(t('Fee') + REFERENCE)}${cellOf(t('Rate'))}</w:tr>` +
    `<w:tr>${cellOf(t('100'))}${cellOf(t('5'))}</w:tr></w:tbl>` +
    p(t('Signed here.')) +
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>',
  MATRIX_EXTRAS
);

async function people(file: Uint8Array, held = false) {
  const room = localRoom(file, { policy: suggestionsOnlyPolicy() });
  const link = new HeldTransport(room.transport({ person: testPerson('Ana'), ...PARTY }));
  const ana = await join(room, 'Ana', { wordFile: file, transport: link, review });
  const olive = await join(room, 'Olive', {
    wordFile: file,
    mode: 'edit',
    login: ORGANISER,
    review,
  });
  link.held = held;
  return { room, ana, olive, link };
}

/** The package an editor shows now. */
const shown = (person: Person): OoxmlPackage => surface(person).session.currentPackage();

/** The notes part of a kind an editor shows, as XML; '' when it shows none. */
function notesXml(person: Person, kind: 'footnotes' | 'endnotes'): string {
  const pkg = shown(person);
  const name = [...pkg.parts.keys()].find((part) => part.startsWith(`/word/${kind}`));
  const part = name === undefined ? undefined : pkg.parts.get(name);
  return part ? serializeOoxmlPart(part) : '';
}

/** Every copy is the room's, with nothing refused, and every editor shows its list. */
function checked(room: OrderedLocalRoom, ...everyone: Person[]): void {
  for (const person of everyone) {
    expect(person.refusals).toEqual([]);
    expect(sameAsRoom(person, room)).toBe(true);
    expect(editorXml(person)).toBe(listXml(person));
    const doc = person.handle.session.orderedDocument();
    // Every notes story of the list shows as the list composes it, with a note or none
    for (const name of [...doc.stories.keys()].filter((n) => /\/(foot|end)notes/.test(n))) {
      const part = shown(person).parts.get(name);
      expect(part ? serializeOoxmlPart(part) : `no ${name}`).toBe(listXml(person, name));
    }
  }
}

const exec = (person: Person, command: Parameters<Person['editor']['exec']>[0]) => {
  const result = person.editor.exec(command);
  expect(result.ok, result.ok ? '' : result.reason).toBe(true);
};

const pendingOf = (room: OrderedLocalRoom, author: string): string[] =>
  [...room.room.document.suggestions.values()]
    .filter((s) => s.status === 'pending' && s.attribution.author === author)
    .map((s) => s.id);

/** A save of a person's editor, read back as a package, and the copy it opens to. */
async function saved(person: Person) {
  const bytes = new Uint8Array(await person.editor.save());
  const read = readOoxmlPackage(bytes, {});
  if (!read.ok) throw new Error(read.reason);
  const reopened = openOrderedWordFile(bytes, { file: 2 });
  return { pkg: read.package, again: composeOrderedPackage(reopened.document, reopened.base) };
}

const xmlOf = (pkg: OoxmlPackage, prefix: string): string => {
  const name = [...pkg.parts.keys()].find((part) => part.startsWith(prefix));
  return name === undefined ? '' : serializeOoxmlPart(pkg.parts.get(name)!);
};

type People = Awaited<ReturnType<typeof people>>;

/** A footnote with words, made by a person at the first paragraph's seventh character. */
async function noteOf(person: Person): Promise<void> {
  caretAt(person, 0, 7);
  expect(surface(person).insertNote('footnote')).toBe(true);
  await settle();
  surface(person).insertPlainText('A note.');
  await settle();
}

/** The `w:id` of every note reference of a kind in a main part, in document order. */
const referenceIds = (xml: string, kind: 'footnote' | 'endnote'): string[] =>
  [...xml.matchAll(new RegExp(`<w:${kind}Reference w:id="(\\d+)"/>`, 'g'))].map((m) => m[1] ?? '');

/** Ana's footnote at the first paragraph's seventh character, typed in, on every copy. */
async function typedNote(room: OrderedLocalRoom, ana: Person, olive: Person): Promise<void> {
  caretAt(ana, 0, 7);
  expect(surface(ana).insertNote('footnote')).toBe(true);
  await settle();
  // The editor is in the new note, as Word leaves it, and types there
  surface(ana).insertPlainText('Ana note.');
  await settle();
  checked(room, ana, olive);
  // Each copy shows the reference as Ana's insertion, and the note with her text
  for (const person of [ana, olive]) {
    const [id = ''] = referenceIds(listXml(person), 'footnote');
    expect(listXml(person)).toMatch(
      new RegExp(`<w:ins [^>]*w:author="Ana"[^>]*><w:r>.*<w:footnoteReference w:id="${id}"/>`)
    );
    expect(notesXml(person, 'footnotes')).toMatch(
      new RegExp(`<w:footnote w:id="${id}">.*Ana note\\.`)
    );
  }
}

/** A notes part that holds Word's separators and no note. */
const SEPARATORS_ONLY = /^(?![\s\S]*<w:(foot|end)noteRef\/>)[\s\S]*w:type="separator"/;

describe('notes in an ordered room (group 9.9)', () => {
  test("a party's footnote in a file without notes: typed and rejected", async () => {
    const { room, ana, olive } = await people(PLAIN);
    await typedNote(room, ana, olive);
    // Rejected: the note goes, and the notes part the suggestion made keeps its separators
    olive.handle.session.requestDecision(pendingOf(room, 'Ana')[0] ?? '', 'rejected');
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).toMatch(SEPARATORS_ONLY);
    }
  });

  // The rejected note's part stays, so the note made again goes into a part the editor holds
  test("a party's footnote made again after a reject: typed, accepted and saved", async () => {
    const { room, ana, olive } = await people(PLAIN);
    await typedNote(room, ana, olive);
    olive.handle.session.requestDecision(pendingOf(room, 'Ana')[0] ?? '', 'rejected');
    await settle();
    await typedNote(room, ana, olive);
    for (const id of pendingOf(room, 'Ana')) olive.handle.session.requestDecision(id, 'accepted');
    await settle();
    checked(room, ana, olive);
    // Saved and opened again: one note, its reference and its text, no revision left
    for (const person of [ana, olive]) {
      const { pkg, again } = await saved(person);
      const ids = referenceIds(xmlOf(pkg, BODY), 'footnote');
      expect(ids).toHaveLength(1);
      expect(xmlOf(pkg, '/word/footnotes')).toMatch(
        new RegExp(`<w:footnote w:id="${ids[0]}">.*Ana note\\.`)
      );
      expect(xmlOf(pkg, BODY)).not.toContain('<w:ins');
      expect(referenceIds(xmlOf(again, BODY), 'footnote')).toEqual(ids);
    }
  });

  // A new note's paragraph had no `w14:paraId` in a notes part the editor already held, so the
  // session could not carry the caret into the note the room numbered, and the next keystroke
  // named the engine's own paragraph: refused as out of step
  test('a new note in a file with a notes part: typed in by each person, on every copy', async () => {
    const { room, ana, olive } = await people(NOTED);
    for (const [person, words] of [
      [ana, 'Ana note.'],
      [olive, 'Olive note.'],
    ] as const) {
      caretAt(person, 1, 4);
      expect(surface(person).insertNote('footnote')).toBe(true);
      await settle();
      surface(person).insertPlainText(words);
      await settle();
      checked(room, ana, olive);
      for (const someone of [ana, olive]) expect(notesXml(someone, 'footnotes')).toContain(words);
    }
  });

  test("the organiser's deletion and conversion, each undone and made again", async () => {
    const { room, ana, olive } = await people(NOTED);
    exec(olive, { type: 'deleteNote', noteKind: 'footnote', noteId: 1 });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) expect(editorXml(person)).not.toContain('footnoteReference');
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive])
      expect(notesXml(person, 'footnotes')).toContain('Fees are due net of tax.');
    exec(olive, { type: 'redo' });
    await settle();
    checked(room, ana, olive);
    exec(olive, { type: 'undo' });
    await settle();
    // Converted: the footnote is the next endnote, and its reference is the endnote's
    exec(olive, { type: 'convertNote', fromKind: 'footnote', noteId: 1 });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(referenceIds(editorXml(person), 'endnote')).toEqual(['2']);
      expect(notesXml(person, 'endnotes')).toMatch(
        /<w:endnote w:id="2">.*Fees are due net of tax\./
      );
      expect(notesXml(person, 'footnotes')).not.toContain('Fees are due');
    }
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive])
      expect(referenceIds(editorXml(person), 'footnote')).toEqual(['1']);
    exec(olive, { type: 'redo' });
    await settle();
    checked(room, ana, olive);
    const { pkg } = await saved(olive);
    expect(referenceIds(xmlOf(pkg, BODY), 'endnote')).toEqual(['2']);
    expect(xmlOf(pkg, ENDNOTES)).toContain('Fees are due net of tax.');
  });

  test("the organiser's note properties for the document and a section, undone", async () => {
    const { room, ana, olive } = await people(NOTED);
    exec(olive, {
      type: 'setNoteProperties',
      scope: 'document',
      footnote: { numFmt: 'lowerRoman' },
    });
    exec(olive, {
      type: 'setNoteProperties',
      scope: 'section',
      sectionIndex: 0,
      footnote: { numRestart: 'eachPage' },
    });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      const settings = settingsPartOf(shown(person));
      expect(settings ? serializeOoxmlPart(settings) : '').toContain(
        '<w:footnotePr><w:numFmt w:val="lowerRoman"/></w:footnotePr>'
      );
      expect(editorXml(person)).toContain(
        '<w:sectPr><w:footnotePr><w:numRestart w:val="eachPage"/></w:footnotePr><w:pgSz'
      );
    }
    exec(olive, { type: 'undo' });
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      const settings = settingsPartOf(shown(person));
      expect(settings ? serializeOoxmlPart(settings) : '').not.toContain('lowerRoman');
      expect(editorXml(person)).not.toContain('footnotePr');
    }
  });

  test("a party's conversion and note properties are refused by the host's party rule", async () => {
    const { room, ana, olive } = await people(NOTED);
    const before = editorXml(olive);
    ana.editor.exec({ type: 'convertNote', fromKind: 'footnote', noteId: 1 });
    ana.editor.exec({ type: 'setNoteProperties', scope: 'document', footnote: { numStart: 3 } });
    await settle();
    // Not a suggestion: today's check gives `package` (the room's `policy-direct`)
    expect(ana.refusals).toEqual([
      expect.objectContaining({ code: 'ordered-refused', reason: 'package' }),
      expect.objectContaining({ code: 'ordered-refused', reason: 'package' }),
    ]);
    ana.refusals.length = 0;
    checked(room, ana, olive);
    expect(editorXml(ana)).toBe(before);
    expect(room.room.order).toBe(0);
  });

  test('two new notes at once: every copy holds both under the same ids, and saves them', async () => {
    const { room, ana, olive, link } = await people(NOTED, true);
    caretAt(ana, 1, 4);
    expect(surface(ana).insertNote('footnote')).toBe(true);
    await settle();
    caretAt(olive, 0, 3);
    expect(surface(olive).insertNote('footnote')).toBe(true);
    await settle();
    link.held = false;
    link.release();
    await settle(16);
    checked(room, ana, olive);
    // The party's note went out first (only the room's frames to her wait): hers is note 2 and
    // the organiser's note 3, on both copies; Word numbers them by their place
    for (const person of [ana, olive]) {
      expect(referenceIds(editorXml(person), 'footnote')).toEqual(['3', '1', '2']);
      expect(notesXml(person, 'footnotes')).toMatch(
        /<w:footnote w:id="2">[\s\S]*<w:footnote w:id="3">/
      );
    }
    const { pkg, again } = await saved(ana);
    expect(referenceIds(xmlOf(pkg, BODY), 'footnote')).toEqual(['3', '1', '2']);
    expect(xmlOf(again, '/word/footnotes')).toBe(xmlOf(pkg, '/word/footnotes'));
  });

  test('one person deletes a note while another types in it: the undo brings both back', async () => {
    const { room, ana, olive, link } = await people(NOTED, true);
    expect(surface(ana).enterNote('footnote:1')).toBe(true);
    surface(ana).insertPlainText('Net. ');
    await settle();
    exec(olive, { type: 'deleteNote', noteKind: 'footnote', noteId: 1 });
    await settle();
    link.held = false;
    link.release();
    await settle(16);
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).not.toContain('Fees are due');
    }
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive])
      expect(notesXml(person, 'footnotes')).toMatch(
        /<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:t xml:space="preserve">Net\. </
      );
  });

  test("a party's undo takes the new note back, and redo makes it again", async () => {
    const { room, ana, olive } = await people(PLAIN);
    caretAt(ana, 1, 3);
    expect(surface(ana).insertNote('endnote')).toBe(true);
    await settle();
    checked(room, ana, olive);
    exec(ana, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('endnoteReference');
      expect(notesXml(person, 'endnotes')).toMatch(SEPARATORS_ONLY);
    }
    // Made again under new ids: past the taken-back note's, which the list keeps hidden
    exec(ana, { type: 'redo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      const [id = ''] = referenceIds(editorXml(person), 'endnote');
      expect(notesXml(person, 'endnotes')).toContain(`<w:endnote w:id="${id}">`);
    }
  });

  test("a party's deleted note: its text struck with its reference, one card, decided and saved", async () => {
    const { room, ana, olive, link } = await people(NOTED, true);
    // Olive types in the note while Ana's deletion is on its way
    exec(ana, { type: 'deleteNote', noteKind: 'footnote', noteId: 1 });
    await settle();
    expect(surface(olive).enterNote('footnote:1')).toBe(true);
    surface(olive).insertPlainText('Net. ');
    await settle();
    link.held = false;
    link.release();
    await settle(16);
    checked(room, ana, olive);
    // The note's own mark and its text, struck by Ana
    const struck = /<w:del [^>]*w:author="Ana"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:footnoteRef\/>/;
    for (const person of [ana, olive]) {
      expect(editorXml(person)).toMatch(/<w:del [^>]*w:author="Ana"[^>]*><w:r>.*footnoteReference/);
      expect(notesXml(person, 'footnotes')).toMatch(struck);
      expect(notesXml(person, 'footnotes')).toMatch(/<w:pPr>.*<w:rPr><w:del [^>]*w:author="Ana"/);
    }
    // One card for the reference and the text
    const cards = () => olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    expect(cards()).toHaveLength(1);
    // Saved and opened again: still one suggestion, drawn the same
    const bytes = new Uint8Array(await olive.editor.save());
    const reopened = openOrderedWordFile(bytes, { file: 2 });
    const pendingAgain = [...reopened.document.suggestions.values()].filter(
      (s) => s.status === 'pending'
    );
    expect(pendingAgain.map((s) => s.attribution.author)).toEqual(['Ana']);
    expect(
      xmlOf(composeOrderedPackage(reopened.document, reopened.base), '/word/footnotes')
    ).toMatch(struck);
    // Rejected: the note shows whole, with Olive's words
    expect(olive.editor.rejectReviewItem(cards()[0]!.key).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(notesXml(person, 'footnotes')).not.toContain('<w:del ');
      expect(notesXml(person, 'footnotes')).toContain('Net. ');
    }
    // Struck again and accepted: the note goes with its reference, and stays gone after a save
    exec(ana, { type: 'deleteNote', noteKind: 'footnote', noteId: 1 });
    await settle();
    expect(cards()).toHaveLength(1);
    expect(olive.editor.acceptReviewItem(cards()[0]!.key).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).not.toContain('Fees are due');
    }
    const { again } = await saved(olive);
    expect(xmlOf(again, '/word/footnotes')).not.toContain('Fees are due');
  });

  test("a party's new note: its runs inserted with its reference, one card, decided and saved", async () => {
    const { room, ana, olive } = await people(PLAIN);
    const cards = () => olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    const suggest = async () => {
      caretAt(ana, 0, 7);
      expect(surface(ana).insertNote('footnote')).toBe(true);
      await settle();
      checked(room, ana, olive);
    };
    // The note's own mark run in Ana's insertion, the paragraph mark unmarked, as Word writes it
    const inserted = /<w:ins [^>]*w:author="Ana"[^>]*><w:r><w:rPr>.*?<\/w:rPr><w:footnoteRef\/>/;
    await suggest();
    for (const person of [ana, olive]) {
      expect(notesXml(person, 'footnotes')).toMatch(inserted);
      expect(notesXml(person, 'footnotes')).not.toContain('<w:rPr><w:ins');
    }
    expect(cards()).toHaveLength(1);
    // Saved and opened again: still one suggestion, drawn the same
    const bytes = new Uint8Array(await olive.editor.save());
    const reopened = openOrderedWordFile(bytes, { file: 2 });
    const pendingAgain = [...reopened.document.suggestions.values()].filter(
      (s) => s.status === 'pending'
    );
    expect(pendingAgain.map((s) => s.attribution.author)).toEqual(['Ana']);
    expect(
      xmlOf(composeOrderedPackage(reopened.document, reopened.base), '/word/footnotes')
    ).toMatch(inserted);
    // Rejected through its card: the note goes with its reference
    expect(olive.editor.rejectReviewItem(cards()[0]!.key).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).toMatch(SEPARATORS_ONLY);
    }
    // Made again and accepted through its card: the note and its reference are plain
    await suggest();
    expect(cards()).toHaveLength(1);
    expect(olive.editor.acceptReviewItem(cards()[0]!.key).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    expect(cards()).toHaveLength(0);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).toContain('<w:footnoteRef/>');
      expect(notesXml(person, 'footnotes')).not.toContain('<w:ins ');
    }
    const { pkg } = await saved(olive);
    expect(xmlOf(pkg, '/word/footnotes') + xmlOf(pkg, BODY)).not.toContain('<w:ins ');
  });

  // Found by the acceptance run: the room showed the note's words, with a card, that the
  // saved file no longer held
  test("a party's deletion over its own new reference takes the note, in the room and the save", async () => {
    const { room, ana, olive } = await people(PLAIN);
    const cards = () => olive.editor.getReviewItems().filter((item) => item.kind === 'revision');
    caretAt(ana, 0, 7);
    expect(surface(ana).insertNote('footnote')).toBe(true);
    await settle();
    surface(ana).insertPlainText('Ana note.');
    await settle();
    checked(room, ana, olive);
    expect(cards()).toHaveLength(1);
    // Back in the body: one character on each side of the reference
    surface(ana).exitNote();
    caretAt(ana, 0, 6, 9);
    surface(ana).deleteSelection();
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).toMatch(SEPARATORS_ONLY);
    }
    // One card, Ana's deletion, in the room and in the editors
    expect(cards()).toHaveLength(1);
    expect(orderedCards(room.room.document).map((card) => card.author)).toEqual(['Ana']);
    // Saved and opened again: the same card, and no note
    const bytes = new Uint8Array(await olive.editor.save());
    const reopened = openOrderedWordFile(bytes, { file: 2 });
    expect(orderedCards(reopened.document).map((card) => card.author)).toEqual(['Ana']);
    const again = composeOrderedPackage(reopened.document, reopened.base);
    expect(xmlOf(again, BODY)).not.toContain('footnoteReference');
    expect(xmlOf(again, '/word/footnotes')).not.toContain('Ana note.');
    // Ana's undo takes back her deletion: every copy agrees, and a save has the room's cards
    exec(ana, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    const undone = openOrderedWordFile(new Uint8Array(await olive.editor.save()), { file: 3 });
    expect(orderedCards(undone.document)).toHaveLength(orderedCards(room.room.document).length);
  });

  // Found by the editor acceptance checks: a date picked in a control replaces its words, and a
  // party's own new reference among them goes with its note. The list composed the notes part
  // without the note, and every editor still showed it
  test("a party's date picked over its own new reference takes the note on every copy", async () => {
    const { room, ana, olive } = await people(DATED);
    caretAt(ana, 0, 14);
    expect(surface(ana).insertNote('footnote')).toBe(true);
    await settle();
    surface(ana).insertPlainText('Ana note.');
    await settle();
    checked(room, ana, olive);
    surface(ana).exitNote();
    caretAt(ana, 0, 16);
    exec(ana, { type: 'setContentControlValue', value: '2026-01-11' } as never);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(editorXml(person)).not.toContain('footnoteReference');
      expect(notesXml(person, 'footnotes')).not.toContain('Ana note.');
    }
  });

  // The same for a row: the organiser's deletion of a row takes the note of the reference in it,
  // and the undo brings both back
  test("the organiser's row deletion takes the note of its reference on every copy", async () => {
    const { room, ana, olive } = await people(ROWED);
    caretAt(olive, 1, 1);
    expect(tableCommand(olive, { type: 'deleteRow' }).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive])
      expect(notesXml(person, 'footnotes')).not.toContain('Fees are due net of tax.');
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive])
      expect(notesXml(person, 'footnotes')).toContain('Fees are due net of tax.');
  });

  // Audit R1 row 31: the same row deletion while another person has words pending in the note.
  // Word and the editor alone remove the note with its reference, the words and their card too;
  // the room kept them shown, with their card, until a reopen. The undo brings all back
  test("the organiser's row deletion takes another person's pending words in its note", async () => {
    const { room, ana, olive } = await people(ROWED);
    const cards = (person: Person) =>
      person.editor.getReviewItems().filter((item) => item.kind === 'revision');
    expect(surface(ana).enterNote('footnote:1')).toBe(true);
    surface(ana).insertPlainText('Net. ');
    await settle();
    checked(room, ana, olive);
    expect(cards(olive)).toHaveLength(1);
    caretAt(olive, 1, 1);
    expect(tableCommand(olive, { type: 'deleteRow' }).ok).toBe(true);
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(notesXml(person, 'footnotes')).not.toContain('Net. ');
      expect(cards(person)).toEqual([]);
    }
    expect(orderedCards(room.room.document)).toEqual([]);
    // The saved file, reopened, holds the same: no note and no card
    const reopened = openOrderedWordFile(new Uint8Array(await olive.editor.save()), { file: 2 });
    expect(orderedCards(reopened.document)).toEqual([]);
    exec(olive, { type: 'undo' });
    await settle();
    checked(room, ana, olive);
    for (const person of [ana, olive]) {
      expect(notesXml(person, 'footnotes')).toContain('Net. ');
      expect(cards(person)).toHaveLength(1);
    }
  });

  // A page that names the words of a note the row deletion took: no copy shows their card, so a
  // decision on it is refused by name, as one on a card the room does not know. Accept All and
  // Reject All pass over it and decide the other cards. The undo brings the words back pending
  for (const status of ['accepted', 'rejected'] as const)
    test(`a decision ${status} on words in a note the row deletion took is refused`, async () => {
      const { room, ana, olive } = await people(ROWED);
      const statusOf = (id: string) => room.room.document.suggestions.get(id)?.status;
      expect(surface(ana).enterNote('footnote:1')).toBe(true);
      surface(ana).insertPlainText('Net. ');
      await settle();
      const [hidden = ''] = pendingOf(room, 'Ana');
      // The card as the organiser's page saw it before the row went
      const seen = orderedCardSeen(olive.handle.session.orderedDocument(), hidden);
      surface(ana).exitNote();
      caretAt(ana, 0, 4);
      surface(ana).insertPlainText('new ');
      await settle();
      const [shownId = ''] = pendingOf(room, 'Ana').filter((id) => id !== hidden);
      caretAt(olive, 1, 1);
      expect(tableCommand(olive, { type: 'deleteRow' }).ok).toBe(true);
      await settle();
      checked(room, ana, olive);
      const doc = olive.handle.session.orderedDocument();
      expect(orderedCards(doc).map((card) => card.id)).toEqual([shownId]);
      expect(orderedCardOf(doc, hidden)).toBeNull();
      expect(olive.handle.session.requestDecision(hidden, status)).toBeNull();
      // The page's view from before the deletion reaches the room
      const actor = {
        person: testPerson('Olive'),
        ...ORGANISER,
        namespace: olive.handle.session.namespace,
      };
      const change = {
        v: 1,
        id: `${actor.namespace}:900`,
        kind: 'decide',
        cards: [{ suggestions: seen }],
        status,
      };
      const sent = room.room.submit(actor, [change]);
      expect(sent.refused.map((refusal) => refusal.reason)).toEqual(['hidden-target']);
      expect(olive.handle.session.requestDecisions([hidden, shownId], status)).not.toBeNull();
      await settle();
      checked(room, ana, olive);
      expect([statusOf(hidden), statusOf(shownId)]).toEqual(['pending', status]);
      exec(olive, { type: 'undo' });
      await settle();
      checked(room, ana, olive);
      for (const person of [ana, olive]) {
        expect(notesXml(person, 'footnotes')).toContain('Net. ');
      }
      expect(statusOf(hidden)).toBe('pending');
      expect(orderedCards(room.room.document).map((card) => card.id)).toEqual([hidden]);
    });

  // Found by the editor acceptance checks: the list kept the new notes part's story after its last
  // note went, and no copy wrote the part. Word keeps a notes part with its separators once its
  // last note goes, and so does the editor alone.
  const endings = {
    "the organiser's deletion over its reference": async ({ olive }: People) => {
      await noteOf(olive);
      surface(olive).exitNote();
      caretAt(olive, 0, 6, 9);
      surface(olive).deleteSelection();
    },
    "the organiser's undo": async ({ olive }: People) => {
      caretAt(olive, 0, 7);
      expect(surface(olive).insertNote('footnote')).toBe(true);
      await settle();
      exec(olive, { type: 'undo' });
    },
    "the party's deletion over its own reference": async ({ ana }: People) => {
      await noteOf(ana);
      surface(ana).exitNote();
      caretAt(ana, 0, 6, 9);
      surface(ana).deleteSelection();
    },
    "the organiser's reject": async ({ room, ana, olive }: People) => {
      await noteOf(ana);
      olive.handle.session.requestDecision(pendingOf(room, 'Ana')[0] ?? '', 'rejected');
    },
  };
  for (const [ending, end] of Object.entries(endings))
    test(`a new notes part stays with its separators after ${ending}`, async () => {
      const met = await people(PLAIN);
      const { room, ana, olive } = met;
      await end(met);
      await settle();
      checked(room, ana, olive);
      for (const person of [ana, olive]) {
        expect(editorXml(person)).not.toContain('footnoteReference');
        expect(notesXml(person, 'footnotes')).toMatch(SEPARATORS_ONLY);
      }
      // Saved and opened again: the same part, and the list holds its story
      const { pkg, again } = await saved(olive);
      expect(xmlOf(pkg, '/word/footnotes')).toMatch(SEPARATORS_ONLY);
      expect(xmlOf(again, '/word/footnotes')).toBe(xmlOf(pkg, '/word/footnotes'));
    });
});
