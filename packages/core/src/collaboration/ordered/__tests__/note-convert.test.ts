// Notes converted to the other kind in the ordered model (group 9.9): the organiser's direct
// edit, as Word converts a note. The copy holds the note's text and every suggestion in it; the
// references become the other kind's; a file with no notes part of that kind gets one; the
// undo (`revert`) gives the file back exactly; and two copies, and either decision order, agree.

import { describe, expect, test } from 'bun:test';
import type { OoxmlPackage } from '../../../store/package/ooxml-package.ts';
import { serializeOoxmlPart } from '../../../store/package/ooxml-tree.ts';
import { applyOrderedChange } from '../change/apply.ts';
import { orderedIdMinter } from '../change/ids.ts';
import type { OrderedChange } from '../change/types.ts';
import { openOrderedWordFile } from '../client/open.ts';
import { composeOrderedPackage, orderedPackageBytes } from '../compose/package.ts';
import { OrderedRefusal } from '../model/refusal.ts';
import type { OrderedDocument } from '../model/types.ts';
import { withUndoRecord } from '../model/undo-record.ts';
import { at, BODY, DATE, FOOTNOTES } from './change-apply-fixture.ts';
import { docx, p, R, t, W, WML_CT } from './word-fixtures.ts';
import { testPerson } from './test-persons.ts';

const reference = (id: number) =>
  `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
const note = (id: number, text: string) =>
  `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>` +
  `<w:footnoteRef/></w:r>${t(text)}</w:p></w:footnote>`;

/** Two footnotes, and no endnotes part. */
const FILE = docx(
  p(t('The fee') + reference(1) + t(' is 100 USD.')) + p(t('Due') + reference(2) + t(' soon.')),
  {
    parts: {
      'word/footnotes.xml':
        `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">` +
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/>' +
        `</w:r></w:p></w:footnote>${note(1, ' Net of tax.')}${note(2, ' In 30 days.')}</w:footnotes>`,
    },
    contentTypes: `<Override PartName="/word/footnotes.xml" ContentType="${WML_CT}.footnotes+xml"/>`,
    relationships: `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>`,
  }
);

type Body = OrderedChange extends infer C
  ? C extends OrderedChange
    ? Omit<C, 'v' | 'id'>
    : never
  : never;

function open(bytes = FILE): { doc: OrderedDocument; base: OoxmlPackage } {
  const opened = openOrderedWordFile(bytes, { file: 1 });
  return { doc: opened.document, base: opened.base };
}

function apply(doc: OrderedDocument, id: string, body: Body, author = 'Olive'): string[] {
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

/** The address of a footnote's open item, by its `w:id`. */
function noteAddress(doc: OrderedDocument, id: number) {
  for (const item of doc.stories.get(FOOTNOTES)?.tree.items() ?? []) {
    const node = item.node;
    if (item.kind !== 'open' || item.token !== 'note' || !node || node.kind === 'textValue')
      continue;
    if (node.attributes.some((a) => a.localName === 'id' && a.value === String(id)))
      return { space: item.space, offset: item.offset };
  }
  throw new Error(`no footnote ${id}`);
}

const convert = (
  doc: OrderedDocument,
  id: string,
  notes: number[],
  mode: 'edit' | 'suggest' = 'edit'
) =>
  apply(doc, id, {
    kind: 'convert-note',
    part: FOOTNOTES,
    mode,
    notes: notes.map((n) => noteAddress(doc, n)),
  })[0] ?? '';

const revert = (doc: OrderedDocument, id: string, suggestion: string) =>
  apply(doc, id, { kind: 'revert', suggestions: [suggestion], status: 'rejected' } as Body);

/** Every part as XML, with the relationships and content types. */
function shape(pkg: OoxmlPackage): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, part] of pkg.parts) out[name] = serializeOoxmlPart(part);
  out['#rels'] = JSON.stringify([...pkg.relationships.entries()]);
  out['#types'] = JSON.stringify(pkg.contentTypes);
  return out;
}

const xmlOf = (pkg: OoxmlPackage, name: string): string => {
  const part = pkg.parts.get(name);
  if (!part) throw new Error(`no part ${name}`);
  return serializeOoxmlPart(part);
};

const endnotesOf = (pkg: OoxmlPackage): string => {
  const name = [...pkg.parts.keys()].find((part) => /\/endnotes_/.test(part));
  return name ? xmlOf(pkg, name) : '';
};

function decide(doc: OrderedDocument, id: string, suggestion: string, status: 'rejected') {
  const found = doc.suggestions.get(suggestion);
  if (!found) throw new Error(`no suggestion ${suggestion}`);
  apply(doc, id, {
    kind: 'decide',
    status,
    cards: [{ suggestions: [{ suggestion, version: found.version }] }],
  } as Body);
}

describe('converted notes (group 9.9)', () => {
  test('a footnote becomes an endnote in a new endnotes part, the same on two copies', () => {
    const one = open();
    const two = open();
    convert(one.doc, 'olive-1:1', [1]);
    convert(two.doc, 'olive-1:1', [1]);
    const pkg = composeOrderedPackage(one.doc, one.base);
    expect(shape(pkg)).toEqual(shape(composeOrderedPackage(two.doc, two.base)));
    const body = xmlOf(pkg, BODY);
    expect(body).toContain(
      '<w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="1"/>'
    );
    expect(body).not.toContain('<w:footnoteReference w:id="1"/>');
    expect(body).toContain('<w:footnoteReference w:id="2"/>');
    const endnotes = endnotesOf(pkg);
    expect(endnotes).toContain('<w:endnote w:id="-1" w:type="separator">');
    expect(endnotes).toMatch(/<w:endnote w:id="1"><w:p[^>]*>.*<w:endnoteRef\/>.*Net of tax\./);
    expect(endnotes).toContain('<w:rStyle w:val="EndnoteReference"/>');
    const footnotes = xmlOf(pkg, FOOTNOTES);
    expect(footnotes).not.toContain('Net of tax.');
    expect(footnotes).toContain('In 30 days.');
    // Saved and opened again, the file composes the same
    const reopened = open(orderedPackageBytes(one.doc, one.base));
    const again = composeOrderedPackage(reopened.doc, reopened.base);
    expect(xmlOf(again, BODY)).toContain('<w:endnoteReference w:id="1"/>');
    expect(endnotesOf(again)).toContain('Net of tax.');
  });

  test('the undo gives the notes back exactly, and the redo the conversion again', () => {
    const { doc, base } = open();
    const before = shape(composeOrderedPackage(doc, base));
    const sid = convert(doc, 'olive-1:1', [1]);
    const converted = shape(composeOrderedPackage(doc, base));
    revert(doc, 'olive-1:2', sid);
    // The endnotes part the conversion made stays with its separators, as Word keeps a notes
    // part once its last note goes; the rest is the file as it was
    const undone = shape(composeOrderedPackage(doc, base));
    const made = Object.keys(undone).filter((name) => !(name in before));
    expect(made).toEqual([expect.stringMatching(/^\/word\/endnotes_/)]);
    expect(undone[made[0]!]).not.toContain('<w:endnoteRef/>');
    expect(undone[BODY]).toBe(before[BODY]);
    expect(undone[FOOTNOTES]).toBe(before[FOOTNOTES]);
    apply(doc, 'olive-1:3', { kind: 'revert', suggestions: [sid], status: 'accepted' } as Body);
    expect(shape(composeOrderedPackage(doc, base))).toEqual(converted);
  });

  test('every footnote, in the order of the notes part, and an endnote back to a footnote', () => {
    const { doc, base } = open();
    convert(doc, 'olive-1:1', [1, 2]);
    const pkg = composeOrderedPackage(doc, base);
    const body = xmlOf(pkg, BODY);
    expect(body.indexOf('<w:endnoteReference w:id="1"/>')).toBeLessThan(
      body.indexOf('<w:endnoteReference w:id="2"/>')
    );
    expect(body).not.toContain('footnoteReference');
    // The footnotes part keeps only Word's separators: no note shows there
    expect(xmlOf(pkg, FOOTNOTES)).not.toMatch(/<w:footnote w:id="[1-9]/);
    const endnotes = endnotesOf(pkg);
    expect(endnotes.indexOf('Net of tax.')).toBeLessThan(endnotes.indexOf('In 30 days.'));
    // Back: the endnote goes to the footnotes part under a new id there
    const name = [...pkg.parts.keys()].find((part) => /\/endnotes_/.test(part)) ?? '';
    const endnote = [...(doc.stories.get(name)?.tree.items() ?? [])].find(
      (item) =>
        item.kind === 'open' &&
        item.node?.kind !== 'textValue' &&
        item.node?.attributes.some((a) => a.localName === 'id' && a.value === '2')
    );
    if (!endnote) throw new Error('no endnote 2');
    apply(doc, 'olive-1:2', {
      kind: 'convert-note',
      part: name,
      mode: 'edit',
      notes: [{ space: endnote.space, offset: endnote.offset }],
    });
    const back = composeOrderedPackage(doc, base);
    expect(xmlOf(back, BODY)).toContain('<w:footnoteReference w:id="3"/>');
    expect(xmlOf(back, FOOTNOTES)).toMatch(/<w:footnote w:id="3">.*In 30 days\./);
  });

  test("another person's typing in the note goes with it; its decision agrees in either order", () => {
    const run = (order: 'convert-first' | 'reject-first') => {
      const { doc, base } = open();
      const [typed = ''] = apply(
        doc,
        'ben-1:1',
        {
          kind: 'insert',
          part: FOOTNOTES,
          mode: 'suggest',
          place: at(doc, 2, 5, FOOTNOTES),
          text: ' all',
        },
        'Ben'
      );
      if (order === 'convert-first') {
        convert(doc, 'olive-1:1', [1]);
        const shown = endnotesOf(composeOrderedPackage(doc, base));
        expect(shown).toMatch(
          /Net<\/w:t><\/w:r><w:ins w:author="Ben"[^>]*><w:r><w:t xml:space="preserve"> all</
        );
        decide(doc, 'olive-1:2', typed, 'rejected');
      } else {
        decide(doc, 'olive-1:2', typed, 'rejected');
        convert(doc, 'olive-1:1', [1]);
      }
      return shape(composeOrderedPackage(doc, base));
    };
    const one = run('convert-first');
    const two = run('reject-first');
    expect(one[BODY]).toEqual(two[BODY]);
    expect(one[FOOTNOTES]).toEqual(two[FOOTNOTES]);
  });

  test("another person's new note, converted: its reference stays their insertion", () => {
    const { doc, base } = open();
    const [added = ''] = apply(
      doc,
      'ben-1:1',
      {
        kind: 'insert-note',
        part: BODY,
        mode: 'suggest',
        place: at(doc, 1, 3),
        note: 'footnote',
        text: 'Ben says.',
      },
      'Ben'
    );
    convert(doc, 'olive-1:1', [3]);
    const body = xmlOf(composeOrderedPackage(doc, base), BODY);
    expect(body).toMatch(/<w:ins w:author="Ben"[^>]*><w:r>.*<w:endnoteReference w:id="1"\/>/);
    decide(doc, 'olive-1:2', added, 'rejected');
    const pkg = composeOrderedPackage(doc, base);
    expect(xmlOf(pkg, BODY)).not.toContain('endnoteReference');
    // The endnotes part the conversion made stays with its separators once its only note goes
    expect(endnotesOf(pkg)).not.toContain('<w:endnoteRef/>');
  });

  test('refused: as a suggestion, a separator, and a note that does not show', () => {
    const { doc } = open();
    const reason = (body: () => unknown): string => {
      try {
        body();
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    };
    expect(reason(() => convert(doc, 'olive-1:1', [1], 'suggest'))).toBe('invalid-edit');
    expect(reason(() => convert(doc, 'olive-1:2', [-1]))).toBe('invalid-edit');
    const sid = convert(doc, 'olive-1:3', [1]);
    expect(sid).not.toBe('');
    expect(reason(() => convert(doc, 'olive-1:4', [1]))).toBe('hidden-target');
  });

  test('typing in a note another change converted is refused by name, never put in the next note', () => {
    const { doc, base } = open();
    // Ana's place in footnote 1, read before Olive's conversion reaches her copy
    const place = at(doc, 2, 5, FOOTNOTES);
    convert(doc, 'olive-1:1', [1]);
    const reason = (() => {
      try {
        apply(
          doc,
          'ana-1:1',
          { kind: 'insert', part: FOOTNOTES, mode: 'suggest', place, text: ' all' },
          'Ana'
        );
      } catch (error) {
        if (error instanceof OrderedRefusal) return error.reason;
        throw error;
      }
      return 'applied';
    })();
    expect(reason).toBe('hidden-target');
    expect(xmlOf(composeOrderedPackage(doc, base), FOOTNOTES)).not.toContain(' all');
  });

  test("new note ids stay within Word's 32767 when the file holds that id", () => {
    // Word reads note ids up to 32767 ([MS-OI29500] 17.11.2(a)), so a file can hold that one;
    // a new note and a converted one take the lowest free id then
    const endnote =
      '<w:endnote w:id="32767"><w:p><w:r><w:endnoteRef/></w:r>' +
      `${t(' Annex.')}</w:p></w:endnote>`;
    const high = docx(
      p(t('The fee') + reference(1) + t(' is 100 USD.')) +
        p(
          t('Due') +
            reference(32767) +
            t(' soon.') +
            '<w:r><w:endnoteReference w:id="32767"/></w:r>'
        ),
      {
        parts: {
          'word/footnotes.xml':
            `<w:footnotes xmlns:w="${W}" xmlns:r="${R}">` +
            `${note(1, ' Net of tax.')}${note(32767, ' In 30 days.')}</w:footnotes>`,
          'word/endnotes.xml': `<w:endnotes xmlns:w="${W}" xmlns:r="${R}">${endnote}</w:endnotes>`,
        },
        contentTypes:
          `<Override PartName="/word/footnotes.xml" ContentType="${WML_CT}.footnotes+xml"/>` +
          `<Override PartName="/word/endnotes.xml" ContentType="${WML_CT}.endnotes+xml"/>`,
        relationships:
          `<Relationship Id="rIdNotes" Type="${R}/footnotes" Target="footnotes.xml"/>` +
          `<Relationship Id="rIdEnd" Type="${R}/endnotes" Target="endnotes.xml"/>`,
      }
    );
    const { doc, base } = open(high);
    apply(doc, 'olive-1:1', {
      kind: 'insert-note',
      part: BODY,
      mode: 'suggest',
      place: at(doc, 0, 3),
      note: 'footnote',
      text: ' New.',
    });
    convert(doc, 'olive-1:2', [1, 32767]);
    const saved = orderedPackageBytes(doc, base);
    const body = xmlOf(composeOrderedPackage(doc, base), BODY);
    expect(body).toContain('<w:footnoteReference w:id="2"/>');
    expect(body).toContain('<w:endnoteReference w:id="1"/>');
    expect(body).toContain('<w:endnoteReference w:id="2"/>');
    const again = composeOrderedPackage(open(saved).doc, open(saved).base);
    for (const [, part] of again.parts) {
      const ids = [
        ...serializeOoxmlPart(part).matchAll(
          /<w:(?:foot|end)note(?:Reference)? [^>]*w:id="(\d+)"/g
        ),
      ];
      for (const [, id] of ids) expect(Number(id)).toBeLessThanOrEqual(32767);
    }
  });
});
