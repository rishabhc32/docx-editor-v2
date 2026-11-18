// Notes in the ordered model (group 9.9): a first note in a file with no notes part makes the
// part as a resource (design D14), the same on every copy; a rejected or taken-back note leaves
// no part behind; two people's first notes share one part; a note deleted directly goes with its
// reference and comes back with an undo; save and reopen keep the notes and their numbers. A
// new note's runs and a struck note's body are drawn as Word writes them, with their reference.
// A new note goes with its reference when its author takes the reference back.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { orderedCardOf } from '../card-of.ts';
import { orderedIdMinter } from '../change/ids.ts';
import { applyOrderedChange } from '../change/apply.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { orderedViewParagraph } from '../compose/compose.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { orderedCards } from '../decide.ts';
import { orderedPlaceAt } from '../edits/places.ts';
import { isHidden } from '../model/document.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { BODY, DATE } from './change-apply-fixture.ts';
import { docx, p, R, t, W, WML_CT } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

/** Two plain paragraphs and no notes part. */
const PLAIN = docx(p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')));

/** The same paragraphs in a file with settings, as Word writes every file. */
const SETTINGS = docx(p(t('The fee is 100 USD.')) + p(t('Due in 30 days.')), {
  parts: {
    'word/settings.xml': `<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/><w:compat/></w:settings>`,
  },
  contentTypes: `<Override PartName="/word/settings.xml" ContentType="${WML_CT}.settings+xml"/>`,
  relationships: `<Relationship Id="rIdSettings" Type="${R}/settings" Target="settings.xml"/>`,
});

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

function open(bytes = PLAIN): { doc: OrderedDocument; base: OoxmlPackage } {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

function apply(doc: OrderedDocument, id: string, body: Body, author = 'Ana'): string[] {
  const change = { v: 1, id, ...body } as OrderedChange;
  return [
    ...withUndoRecord(doc, () =>
      applyOrderedChange(doc, change, {
        person: testPerson(author),
        date: DATE,
        ids: orderedIdMinter(id),
      })
    ).value.suggestions,
  ];
}

const at = (doc: OrderedDocument, index: number, n: number) => {
  const view = orderedViewParagraph(doc, BODY, index);
  if (!view) throw new Error(`no paragraph ${index}`);
  return orderedPlaceAt(view, n);
};

/** A new note by a person at a paragraph's offset, typed with `text`. */
const newNote = (
  doc: OrderedDocument,
  id: string,
  index: number,
  n: number,
  author = 'Ana',
  note: 'footnote' | 'endnote' = 'footnote',
  mode: 'suggest' | 'edit' = 'suggest'
): string =>
  apply(
    doc,
    id,
    {
      kind: 'insert-note',
      part: BODY,
      mode,
      place: at(doc, index, n),
      note,
      text: `${author}'s note.`,
    },
    author
  )[0] ?? '';

/** The parts of a package by name, as XML, and its relationships and content types. */
function shape(pkg: OoxmlPackage): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, part] of pkg.parts) out[name] = serializeOoxmlPart(part);
  out['#rels'] = JSON.stringify([...pkg.relationships.entries()]);
  return out;
}

/** A new notes part's separators, as the settings list them. */
const SEPARATORS_LISTED =
  '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>';

const settingsOf = (pkg: OoxmlPackage): string => {
  const settings = pkg.parts.get('/word/settings.xml');
  return settings ? serializeOoxmlPart(settings) : '';
};

const notesParts = (pkg: OoxmlPackage) =>
  [...pkg.parts.keys()].filter((name) => /footnotes|endnotes/.test(name)).sort();

/**
 * The notes parts that hold a note past Word's separators: a notes part stays with its
 * separators once its last note goes, as Word and the editor alone keep it.
 */
const partsWithNotes = (pkg: OoxmlPackage) =>
  notesParts(pkg).filter((name) =>
    /<w:(foot|end)noteRef\/>/.test(serializeOoxmlPart(pkg.parts.get(name)!))
  );

function decide(doc: OrderedDocument, suggestion: string, status: 'accepted' | 'rejected') {
  const found = doc.suggestions.get(suggestion);
  if (!found) throw new Error(`no suggestion ${suggestion}`);
  apply(
    doc,
    `olive:${suggestion.length}${status.length}`,
    {
      kind: 'decide',
      status,
      cards: [{ suggestions: [{ suggestion, version: found.version }] }],
    } as Body,
    'Olive'
  );
}

describe('notes (group 9.9)', () => {
  test('a first footnote makes the notes part, the same on two copies, and reopens', () => {
    const one = open(SETTINGS);
    const two = open(SETTINGS);
    newNote(one.doc, 'ana-1:1', 0, 7);
    newNote(two.doc, 'ana-1:1', 0, 7);
    const pkg = composeOrderedPackage(one.doc, one.base);
    expect(shape(pkg)).toEqual(shape(composeOrderedPackage(two.doc, two.base)));
    const [notes = ''] = notesParts(pkg);
    expect(notes).toMatch(/^\/word\/footnotes_ana-1_1_\d+\.xml$/);
    const xml = serializeOoxmlPart(pkg.parts.get(notes) ?? pkg.parts.get(BODY)!);
    expect(xml).toContain('<w:footnote w:id="-1" w:type="separator">');
    expect(xml).toContain('<w:footnote w:id="0" w:type="continuationSeparator">');
    expect(xml).toContain('Ana&apos;s note.');
    // The main document names it by type, and the content type is written
    const rels = pkg.relationships.get(BODY) ?? [];
    expect(rels.some((r) => r.type.endsWith('/footnotes') && notes.endsWith(r.rawTarget))).toBe(
      true
    );
    // The settings list its separators, as the engine alone and Word write them; a file with
    // no settings gets no part for the list, as in the engine alone
    expect(settingsOf(pkg)).toContain(SEPARATORS_LISTED);
    const plain = open();
    newNote(plain.doc, 'ana-1:1', 0, 7);
    expect(composeOrderedPackage(plain.doc, plain.base).parts.has('/word/settings.xml')).toBe(
      false
    );
    // Saved and opened again: the note is the file's, with its reference
    const reopened = open(orderedPackageBytes(one.doc, one.base));
    const again = composeOrderedPackage(reopened.doc, reopened.base);
    expect(notesParts(again)).toEqual([notes]);
    expect(serializeOoxmlPart(again.parts.get(notes)!)).toContain('Ana&apos;s note.');
    expect(serializeOoxmlPart(again.parts.get(BODY)!)).toContain('<w:footnoteReference w:id="1"/>');
    expect(settingsOf(again)).toContain(SEPARATORS_LISTED);
  });

  test('a rejected or taken-back first note leaves its notes part with its separators', () => {
    const rejected = open();
    const sid = newNote(rejected.doc, 'ana-1:1', 0, 7);
    decide(rejected.doc, sid, 'rejected');
    const pkg = composeOrderedPackage(rejected.doc, rejected.base);
    expect(notesParts(pkg)).toHaveLength(1);
    expect(partsWithNotes(pkg)).toEqual([]);
    expect((pkg.relationships.get(BODY) ?? []).some((r) => r.type.endsWith('/footnotes'))).toBe(
      true
    );
    expect(shape(pkg)[BODY]).not.toContain('footnoteReference');
    // Saved and opened again: the same part, which the list holds
    const reopened = open(orderedPackageBytes(rejected.doc, rejected.base));
    expect(shape(composeOrderedPackage(reopened.doc, reopened.base))[notesParts(pkg)[0]!]).toBe(
      shape(pkg)[notesParts(pkg)[0]!]
    );
    const retracted = open();
    const own = newNote(retracted.doc, 'ana-1:1', 0, 7);
    apply(retracted.doc, 'ana-1:2', { kind: 'retract', suggestions: [own] } as Body);
    const taken = composeOrderedPackage(retracted.doc, retracted.base);
    expect([notesParts(taken).length, partsWithNotes(taken)]).toEqual([1, []]);
  });

  test("two people's first notes share one part and number in either decision order", () => {
    const run = (order: 'ana-first' | 'ben-first') => {
      const { doc, base } = open();
      const ana = newNote(doc, 'ana-1:1', 0, 7, 'Ana');
      const ben = newNote(doc, 'ben-1:1', 1, 3, 'Ben', 'footnote');
      const pkg = composeOrderedPackage(doc, base);
      expect(notesParts(pkg)).toHaveLength(1);
      const body = shape(pkg)[BODY] ?? '';
      expect(body).toContain('<w:footnoteReference w:id="1"/>');
      expect(body).toContain('<w:footnoteReference w:id="2"/>');
      // The first note's suggestion rejected: the part stays for the second note
      const [first, second] = order === 'ana-first' ? [ana, ben] : [ben, ana];
      decide(doc, first, 'rejected');
      const after = composeOrderedPackage(doc, base);
      expect(notesParts(after)).toHaveLength(1);
      decide(doc, second, 'rejected');
      expect(partsWithNotes(composeOrderedPackage(doc, base))).toEqual([]);
      return notesParts(after);
    };
    expect(run('ana-first')).toEqual(run('ben-first'));
  });

  test('an endnote and a footnote make their own parts', () => {
    const { doc, base } = open();
    newNote(doc, 'ana-1:1', 0, 7, 'Ana', 'endnote');
    newNote(doc, 'ana-1:2', 1, 3, 'Ana', 'footnote');
    const parts = notesParts(composeOrderedPackage(doc, base));
    expect(parts).toHaveLength(2);
    expect(parts.some((name) => name.includes('/endnotes_'))).toBe(true);
    expect(parts.some((name) => name.includes('/footnotes_'))).toBe(true);
  });

  test("a direct deletion of a note's reference takes the note, and its undo brings it back", () => {
    const { doc, base } = open();
    const sid = newNote(doc, 'olive-1:1', 0, 7, 'Olive', 'footnote', 'edit');
    expect(doc.suggestions.get(sid)?.status).toBe('accepted');
    const [notes = ''] = notesParts(composeOrderedPackage(doc, base));
    const [deletion = ''] = apply(
      doc,
      'olive-1:2',
      { kind: 'delete', part: BODY, mode: 'edit', from: at(doc, 0, 7), to: at(doc, 0, 8) },
      'Olive'
    );
    const gone = composeOrderedPackage(doc, base);
    expect(serializeOoxmlPart(gone.parts.get(BODY)!)).not.toContain('footnoteReference');
    // The part made in this session stays with its separators once its only note goes
    expect(notesParts(gone)).toEqual([notes]);
    expect(partsWithNotes(gone)).toEqual([]);
    apply(
      doc,
      'olive-1:3',
      { kind: 'revert', suggestions: [deletion], status: 'rejected' } as Body,
      'Olive'
    );
    const back = composeOrderedPackage(doc, base);
    expect(notesParts(back)).toEqual([notes]);
    expect(serializeOoxmlPart(back.parts.get(notes)!)).toContain('Olive&apos;s note.');
  });

  test('a change in a note whose reference went is refused by name; a struck one takes it', () => {
    const { doc } = open();
    newNote(doc, 'olive-1:1', 0, 7, 'Olive', 'footnote', 'edit');
    const notes = [...doc.stories.keys()].find((name) => /\/footnotes_/.test(name)) ?? '';
    const typing = (): Body => {
      const view = orderedViewParagraph(doc, notes, 2);
      if (!view) throw new Error('no note paragraph');
      return {
        kind: 'insert',
        part: notes,
        mode: 'suggest',
        place: orderedPlaceAt(view, 1),
        text: 'x',
      };
    };
    const reason = (id: string, change: Body): string => {
      try {
        apply(doc, id, change);
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    const [deletion = ''] = apply(doc, 'ana-1:1', {
      kind: 'delete',
      part: BODY,
      mode: 'suggest',
      from: at(doc, 0, 7),
      to: at(doc, 0, 8),
    });
    // Struck, the note still shows, as Word shows it: typing in it is a suggestion
    expect(reason('ana-1:2', typing())).toBe('applied');
    // A place from before the decision, as a copy that has not seen it yet sends it
    const late = typing();
    // Her words beside her struck note text make one card with the deletion (policy 3)
    const card = orderedCardOf(doc, deletion)?.suggestions ?? [];
    expect(card).toHaveLength(2);
    apply(
      doc,
      'olive-1:4',
      {
        kind: 'decide',
        status: 'accepted',
        cards: [
          {
            suggestions: card.map((suggestion) => ({
              suggestion,
              version: doc.suggestions.get(suggestion)!.version,
            })),
          },
        ],
      } as Body,
      'Olive'
    );
    expect(reason('ana-1:3', late)).toBe('hidden-target');
  });

  test("a struck note's body shows struck with its reference, as Word writes it", () => {
    const pending = (doc: OrderedDocument) =>
      [...doc.suggestions.values()].filter((s) => s.status === 'pending').map((s) => s.id);
    const strike = () => {
      const opened = open();
      newNote(opened.doc, 'olive-1:1', 0, 7, 'Olive', 'footnote', 'edit');
      const [deletion = ''] = apply(opened.doc, 'ana-1:1', {
        kind: 'delete',
        part: BODY,
        mode: 'suggest',
        from: at(opened.doc, 0, 7),
        to: at(opened.doc, 0, 8),
      });
      return { ...opened, deletion };
    };
    const { doc, base, deletion } = strike();
    const pkg = composeOrderedPackage(doc, base);
    const [notes = ''] = notesParts(pkg);
    const body = /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(
      serializeOoxmlPart(pkg.parts.get(notes)!)
    )![0];
    // The paragraph mark and every run, the note's own mark too, struck by Ana at her date
    expect(body).toMatch(/<w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/);
    expect(body).toMatch(/<w:del [^>]*><w:r><w:rPr><w:rStyle [^>]*\/><\/w:rPr><w:footnoteRef\/>/);
    expect(body).toContain('<w:delText>Olive&apos;s note.</w:delText>');
    for (const mark of body.match(/<w:del [^>]*>/g) ?? [])
      expect(mark).toContain(`w:author="Ana" w:date="${DATE}"`);
    expect(body).not.toMatch(/<w:t>[^<]/);
    const story = [...doc.stories.keys()].find((name) => name === notes) ?? notes;
    const view = orderedViewParagraph(doc, story, 2);
    expect(view?.chars.every((c) => c.deleted)).toBe(true);
    // One suggestion; saved and opened again, still one, written the same
    expect(pending(doc)).toEqual([deletion]);
    const reopened = open(orderedPackageBytes(doc, base));
    expect(pending(reopened.doc)).toHaveLength(1);
    // (the relationships' list order is the package's own business)
    const parts = (one: OoxmlPackage) => ({ ...shape(one), '#rels': '' });
    expect(parts(composeOrderedPackage(reopened.doc, reopened.base))).toEqual(parts(pkg));
    // Rejected, the note shows whole; accepted, it goes with its reference
    for (const copy of [{ doc, base }, reopened]) {
      const [sid = ''] = pending(copy.doc);
      decide(copy.doc, sid, 'rejected');
      const xml = serializeOoxmlPart(composeOrderedPackage(copy.doc, copy.base).parts.get(notes)!);
      expect(xml).not.toContain('<w:del ');
      expect(xml).toContain('<w:t>Olive&apos;s note.</w:t>');
    }
    const accepted = strike();
    decide(accepted.doc, accepted.deletion, 'accepted');
    expect(partsWithNotes(composeOrderedPackage(accepted.doc, accepted.base))).toEqual([]);
    // Words typed in the struck note afterwards are not struck, as in the editor alone
    const typed = strike();
    const notesStory = [...typed.doc.stories.keys()].find((name) => /\/footnotes_/.test(name))!;
    const noteView = orderedViewParagraph(typed.doc, notesStory, 2);
    if (!noteView) throw new Error('no note paragraph');
    apply(
      typed.doc,
      'olive-1:9',
      {
        kind: 'insert',
        part: notesStory,
        mode: 'edit',
        place: orderedPlaceAt(noteView, 1),
        text: 'Z',
      },
      'Olive'
    );
    const xml = serializeOoxmlPart(composeOrderedPackage(typed.doc, typed.base).parts.get(notes)!);
    expect(xml).toContain('<w:t>Z</w:t>');
  });

  test("a struck note's paragraph mark has a w:id apart from the body's struck mark", () => {
    const { doc, base } = open();
    newNote(doc, 'olive-1:1', 0, 7, 'Olive', 'footnote', 'edit');
    // Ana strikes the reference, the rest of its paragraph, its mark and words of the next one
    apply(doc, 'ana-1:1', {
      kind: 'delete',
      part: BODY,
      mode: 'suggest',
      from: at(doc, 0, 7),
      to: at(doc, 1, 3),
    });
    const pkg = composeOrderedPackage(doc, base);
    const [notes = ''] = notesParts(pkg);
    // A paragraph mark's deletion: the `w:del` in its run properties
    const markIds = (xml: string) =>
      [...xml.matchAll(/<w:rPr><w:del [^>]*?w:id="(\d+)"[^>]*\/>/g)].map((m) => m[1]);
    const note = /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(
      serializeOoxmlPart(pkg.parts.get(notes)!)
    )![0];
    const [bodyMark] = markIds(serializeOoxmlPart(pkg.parts.get(BODY)!));
    const [noteMark] = markIds(note);
    expect(bodyMark).toBeDefined();
    expect(noteMark).toBeDefined();
    // As Word writes them: one `w:id` would read as one revision across the two parts
    expect(noteMark).not.toBe(bodyMark);
    // Opened again, the note's marks join the reference's suggestion, and the file is the same
    const again = open(orderedPackageBytes(doc, base));
    const parts = (one: OoxmlPackage) => ({ ...shape(one), '#rels': '' });
    expect(parts(composeOrderedPackage(again.doc, again.base))).toEqual(parts(pkg));
    const reference = [...again.doc.suggestions.keys()].find((id) =>
      [...(again.doc.labelled.get(id) ?? [])].some(
        (item) => item.node?.kind !== 'textValue' && item.node?.localName === 'footnoteReference'
      )
    );
    const card =
      reference === undefined ? [] : (orderedCardOf(again.doc, reference)?.suggestions ?? []);
    expect(card.length).toBeGreaterThan(0);
    // Rejected, the note shows whole; the body's struck mark is a decision of its own
    apply(
      again.doc,
      'olive-1:2',
      {
        kind: 'decide',
        status: 'rejected',
        cards: [
          {
            suggestions: card.map((suggestion) => ({
              suggestion,
              version: again.doc.suggestions.get(suggestion)!.version,
            })),
          },
        ],
      } as Body,
      'Olive'
    );
    const after = composeOrderedPackage(again.doc, again.base);
    expect(serializeOoxmlPart(after.parts.get(notes)!)).not.toContain('<w:del ');
    expect(markIds(serializeOoxmlPart(after.parts.get(BODY)!))).toEqual([bodyMark]);
  });

  test("a new note's runs show inserted with its reference, as Word writes it", () => {
    const pending = (doc: OrderedDocument) =>
      [...doc.suggestions.values()].filter((s) => s.status === 'pending').map((s) => s.id);
    const { doc, base } = open();
    const sid = newNote(doc, 'ana-1:1', 0, 7);
    const pkg = composeOrderedPackage(doc, base);
    const [notes = ''] = notesParts(pkg);
    const note = /<w:footnote w:id="1">[\s\S]*?<\/w:footnote>/.exec(
      serializeOoxmlPart(pkg.parts.get(notes)!)
    )![0];
    // The note's mark run and its text in one insertion by Ana at her date; the mark unmarked
    expect(note).toMatch(
      /<w:p [^>]*><w:ins [^>]*><w:r><w:rPr><w:rStyle [^>]*\/><\/w:rPr><w:footnoteRef\/><\/w:r>[\s\S]*Ana&apos;s note\.[\s\S]*<\/w:ins><\/w:p>/
    );
    expect(note).not.toContain('<w:rPr><w:ins');
    for (const mark of note.match(/<w:ins [^>]*>/g) ?? [])
      expect(mark).toContain(`w:author="Ana" w:date="${DATE}"`);
    // Its own `w:id`, apart from the reference's, as Word gives it
    const insIds = (xml: string) =>
      [...xml.matchAll(/<w:ins [^>]*?w:id="(\d+)"/g)].map((m) => m[1]);
    const [referenceId] = insIds(serializeOoxmlPart(pkg.parts.get(BODY)!));
    expect(insIds(note)).toHaveLength(1);
    expect(insIds(note)[0]).not.toBe(referenceId);
    // One suggestion, one card; saved and opened again, the same, written the same
    expect(pending(doc)).toEqual([sid]);
    expect(orderedCardOf(doc, sid)?.suggestions).toEqual([sid]);
    const reopened = open(orderedPackageBytes(doc, base));
    expect(pending(reopened.doc)).toHaveLength(1);
    const parts = (one: OoxmlPackage) => ({ ...shape(one), '#rels': '' });
    expect(parts(composeOrderedPackage(reopened.doc, reopened.base))).toEqual(parts(pkg));
    // Accepted, the note and its reference are plain; rejected, the note goes (and the part the
    // change made goes with it; a reopened file's own part stays with its separators)
    const bytes = orderedPackageBytes(doc, base);
    const made = () => {
      const copy = open();
      newNote(copy.doc, 'ana-1:1', 0, 7);
      return copy;
    };
    for (const status of ['accepted', 'rejected'] as const) {
      for (const copy of [made(), open(bytes)]) {
        const [one = ''] = pending(copy.doc);
        decide(copy.doc, one, status);
        const after = composeOrderedPackage(copy.doc, copy.base);
        const left = after.parts.get(notes);
        if (status === 'rejected')
          expect(left ? serializeOoxmlPart(left) : '').not.toContain('<w:footnote w:id="1">');
        else {
          expect(serializeOoxmlPart(after.parts.get(notes)!)).toContain('Ana&apos;s note.');
          for (const xml of [after.parts.get(notes)!, after.parts.get(BODY)!])
            expect(serializeOoxmlPart(xml)).not.toContain('<w:ins ');
        }
      }
    }
  });

  test("Word's own new note opens as one suggestion; words typed later keep their own", () => {
    // As Word 16 wrote it (sample `a-new-note-word.docx`): ids 0 and 1, one author and date
    const stamp = 'w:author="Ana" w:date="2026-10-29T23:17:00Z"';
    const later =
      '<w:ins w:id="2" w:author="Ana" w:date="2026-10-30T09:00:00Z">' + t(' Later.') + '</w:ins>';
    const bytes = docx(
      p(
        t('The fee') +
          `<w:ins w:id="0" ${stamp}><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
          '<w:footnoteReference w:id="1"/></w:r></w:ins>' +
          t(' is 100 USD.')
      ),
      {
        parts: {
          'word/footnotes.xml':
            `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">` +
            '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
            '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
            `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:ins w:id="1" ${stamp}>` +
            '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>' +
            `${t(' New note.')}</w:ins>${later}</w:p></w:footnote></w:footnotes>`,
        },
        contentTypes: `<Override PartName="/word/footnotes.xml" ContentType="${WML_CT}.footnotes+xml"/>`,
        relationships: `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>`,
      }
    );
    const opened = open(bytes);
    const pending = [...opened.doc.suggestions.values()].filter((s) => s.status === 'pending');
    // The reference and the note's runs are one suggestion; the later words another
    expect(pending.map((s) => s.attribution.date)).toEqual([
      '2026-10-29T23:17:00Z',
      '2026-10-30T09:00:00Z',
    ]);
    const [note, typed] = pending;
    expect(orderedCardOf(opened.doc, note!.id)?.suggestions).toEqual([note!.id]);
    // Written back as Word wrote it: the same ids
    const xml = serializeOoxmlPart(
      composeOrderedPackage(opened.doc, opened.base).parts.get('/word/footnotes.xml')!
    );
    expect(xml).toMatch(
      /<w:ins [^>]*w:id="1"><w:r><w:rPr><w:rStyle [^>]*\/><\/w:rPr><w:footnoteRef\/>/
    );
    expect(xml).toMatch(/<w:ins [^>]*w:id="2">/);
    // Accepted, the note's runs are plain and the later words still pending
    decide(opened.doc, note!.id, 'accepted');
    const after = serializeOoxmlPart(
      composeOrderedPackage(opened.doc, opened.base).parts.get('/word/footnotes.xml')!
    );
    expect(after).not.toMatch(/<w:ins [^>]*w:id="1"/);
    expect(after).toMatch(/<w:ins [^>]*w:id="2"><w:r><w:t xml:space="preserve"> Later\.<\/w:t>/);
    expect(opened.doc.suggestions.get(typed!.id)?.status).toBe('pending');
  });
});

/**
 * Each review card as its author and the shown words it inserts and strikes: the same for a
 * session and its saved file opened again.
 */
function cardWords(doc: OrderedDocument): string[] {
  return orderedCards(doc).map((card) => {
    const ids = new Set(card.suggestions);
    let inserted = '';
    let struck = '';
    for (const story of doc.stories.values())
      for (const item of story.tree.items()) {
        if (item.kind !== 'text' || isHidden(doc, item)) continue;
        if (item.by !== null && ids.has(item.by)) inserted += item.text;
        if (item.dels.some((id) => ids.has(id))) struck += item.text;
      }
    return `${card.author} +${inserted} -${struck}`;
  });
}

/** The shown words of the notes parts. */
const noteWords = (doc: OrderedDocument): string =>
  [...doc.stories.values()]
    .filter((story) => /footnotes|endnotes/.test(story.name))
    .flatMap((story) => [...story.tree.items()])
    .filter((item) => item.kind === 'text' && !isHidden(doc, item))
    .map((item) => (item.kind === 'text' ? item.text : ''))
    .join('');

/** A suggested deletion of the view characters `[from, to)` of a paragraph of the body. */
const deletion = (doc: OrderedDocument, index: number, from: number, to: number): Body => ({
  kind: 'delete',
  part: BODY,
  mode: 'suggest',
  from: at(doc, index, from),
  to: at(doc, index, to),
});

describe('a new note goes with its reference (group 9.9)', () => {
  // Found by the acceptance run (a party's quick pass): a party adds a footnote, then strikes
  // words around its own new reference. The reference is taken back, as Word removes a person's
  // own insertion, and the save writes no note without a reference. The room showed the note's
  // words with a card of their own, which a reopen lost.
  test("a deletion over the author's own new reference takes the note, as the save does", () => {
    const { doc, base } = open();
    newNote(doc, 'ana-1:1', 0, 7);
    expect(cardWords(doc)).toEqual(["Ana +Ana's note. -"]);
    // One unit on each side of the reference
    apply(doc, 'ana-1:2', deletion(doc, 0, 6, 9));
    expect(noteWords(doc)).toBe('');
    expect(cardWords(doc)).toEqual(['Ana + -e ']);
    const pkg = composeOrderedPackage(doc, base);
    expect(partsWithNotes(pkg)).toEqual([]);
    expect(serializeOoxmlPart(pkg.parts.get(BODY)!)).not.toContain('footnoteReference');
    const reopened = open(orderedPackageBytes(doc, base));
    expect(cardWords(reopened.doc)).toEqual(cardWords(doc));
    expect(noteWords(reopened.doc)).toBe('');
  });

  test('the reference alone, and words its author typed in the note later, go with it', () => {
    const { doc, base } = open();
    newNote(doc, 'ana-1:1', 0, 7);
    const notes = [...doc.stories.keys()].find((name) => /\/footnotes_/.test(name)) ?? '';
    const view = orderedViewParagraph(doc, notes, 2);
    if (!view) throw new Error('no note paragraph');
    apply(doc, 'ana-1:2', {
      kind: 'insert',
      part: notes,
      mode: 'suggest',
      place: orderedPlaceAt(view, 1),
      text: 'See ',
    });
    expect(noteWords(doc)).toBe("See Ana's note.");
    apply(doc, 'ana-1:3', deletion(doc, 0, 7, 8));
    expect([noteWords(doc), cardWords(doc)]).toEqual(['', []]);
    const pkg = composeOrderedPackage(doc, base);
    expect(partsWithNotes(pkg)).toEqual([]);
    const reopened = open(orderedPackageBytes(doc, base));
    expect([noteWords(reopened.doc), cardWords(reopened.doc)]).toEqual(['', []]);
  });

  test("a new note that holds another person's suggestion is refused by name, as a new row is", () => {
    const { doc } = open();
    newNote(doc, 'ana-1:1', 0, 7);
    const notes = [...doc.stories.keys()].find((name) => /\/footnotes_/.test(name)) ?? '';
    const view = orderedViewParagraph(doc, notes, 2);
    if (!view) throw new Error('no note paragraph');
    apply(
      doc,
      'olive-1:1',
      {
        kind: 'insert',
        part: notes,
        mode: 'suggest',
        place: orderedPlaceAt(view, 1),
        text: 'Per ',
      },
      'Olive'
    );
    const before = cardWords(doc);
    let reason = 'applied';
    try {
      apply(doc, 'ana-1:2', deletion(doc, 0, 6, 9));
    } catch (error) {
      if (!(error instanceof OrderedRefusal)) throw error;
      reason = error.reason;
    }
    // Taking the note back would take Olive's words back too: nothing changes
    expect(reason).toBe('others-in-row');
    expect([noteWords(doc), cardWords(doc)]).toEqual(["Per Ana's note.", before]);
  });
});
